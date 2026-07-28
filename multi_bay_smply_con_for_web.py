import streamlit as st
import openseespy.opensees as ops
import opsvis as opsv
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.ticker as ticker

# ============================================================
# 1. 固定設定與共用函式
# ============================================================
ELEMENTS_PER_SEGMENT = 40  # 每一跨、下部懸臂與頂部懸臂一律切分為 40 個梁元素
POSTPROCESS_POINTS_PER_ELEMENT = 40  # 每個元素後處理取樣點數
WIND_DIRECTION_X = 1.0

# 整合 E (彈性模數) 至材料資料庫，方便擴充不同材質[cite: 1]
MATERIALS = {
    1: {"name": "鋁合金 6105-T5", "E": 710100.0, "Fty": 2450.0, "Ftu": 2660.0, "Fb_allow": 1485.0},
    2: {"name": "鋁合金 6061-T6", "E": 710100.0, "Fty": 2450.0, "Ftu": 2660.0, "Fb_allow": 1485.0},
    3: {"name": "鋁合金 6063-T5", "E": 710100.0, "Fty": 1050.0, "Ftu": 1470.0, "Fb_allow": 682.0},
    4: {"name": "鋁合金 6063-T6", "E": 710100.0, "Fty": 1750.0, "Ftu": 2100.0, "Fb_allow": 1060.0},
    5: {"name": "碳鋼 A36", "E": 2.04e6, "Fty": 2530.0, "Ftu": 4100.0, "Fb_allow": 1518.0},
}

def status_text(passed):
    return "🟢 PASS (通過)" if passed else "🔴 FAIL (不通過)"

def span_allowable_deflection(length_cm):
    """支承間容許變形"""
    if length_cm * 10.0 < 4115.0:
        return (length_cm / 175.0, "L/175")
    return (length_cm / 240.0 + 0.635, "L/240 + 0.635 cm")

# ============================================================
# 2. Streamlit 介面與主程式
# ============================================================
st.set_page_config(page_title="多跨度直料檢核系統", layout="wide")
st.title("🏗️ 帷幕牆多跨直料風力、強軸應力與容許變形檢核")

# --- 左側輸入區 ---
st.sidebar.header("📝 1. 結構與載重參數")
w = st.sidebar.number_input("風力線載重 w (kgf/cm, >0)", min_value=0.001, value=1.5, format="%.4f")

# 懸臂與跨度動態輸入
add_bottom = st.sidebar.checkbox("增加下部懸臂", value=True)
bottom_cantilever_length = st.sidebar.number_input("下部懸臂長度 a (cm)", min_value=0.01, value=30.0) if add_bottom else 0.0

num_spans = st.sidebar.number_input("支承間跨數", min_value=1, max_value=10, value=1, step=1)
span_lengths = []
for i in range(num_spans):
    span_len = st.sidebar.number_input(f"第 {i+1} 跨支承間距 L{i+1} (cm)", min_value=1.0, value=300.0)
    span_lengths.append(span_len)

add_top = st.sidebar.checkbox("增加頂部懸臂", value=True)
top_cantilever_length = st.sidebar.number_input("頂部懸臂長度 b (cm)", min_value=0.01, value=30.0) if add_top else 0.0

st.sidebar.header("📝 2. 材料設定")
material_choice = st.sidebar.selectbox(
    "選擇材料", 
    options=list(MATERIALS.keys()), 
    format_func=lambda x: f"{MATERIALS[x]['name']}"
)

section_mode = st.sidebar.radio("直料區分公料與母料？", options=[1, 2], format_func=lambda x: "否，整體單一斷面" if x == 1 else "是，公料與母料分開輸入")

if section_mode == 1:
    A1 = st.sidebar.number_input("總斷面積 A (cm^2)", min_value=0.001, value=15.0)
    Ix1 = st.sidebar.number_input("總強軸慣性矩 Ix (cm^4)", min_value=0.001, value=120.0)
    Sx1 = st.sidebar.number_input("總強軸斷面係數 Sx (cm^3)", min_value=0.001, value=25.0)
    profiles = [{"name": "整體直料", "A": A1, "Ix": Ix1, "Sx": Sx1}]
else:
    st.sidebar.markdown("**公料**")
    A1 = st.sidebar.number_input("公料 A1 (cm^2)", min_value=0.001, value=8.0)
    Ix1 = st.sidebar.number_input("公料 Ix1 (cm^4)", min_value=0.001, value=60.0)
    Sx1 = st.sidebar.number_input("公料 Sx1 (cm^3)", min_value=0.001, value=12.0)
    st.sidebar.markdown("**母料**")
    A2 = st.sidebar.number_input("母料 A2 (cm^2)", min_value=0.001, value=7.0)
    Ix2 = st.sidebar.number_input("母料 Ix2 (cm^4)", min_value=0.001, value=50.0)
    Sx2 = st.sidebar.number_input("母料 Sx2 (cm^3)", min_value=0.001, value=10.0)
    profiles = [{"name": "公料", "A": A1, "Ix": Ix1, "Sx": Sx1}, {"name": "母料", "A": A2, "Ix": Ix2, "Sx": Sx2}]

# --- 執行分析 ---
if st.sidebar.button("🚀 執行檢核分析", type="primary"):
    mat = MATERIALS[material_choice]
    E_val = mat["E"]  # 動態讀取該材料的彈性模數
    
    Area_mullion = sum(item["A"] for item in profiles)
    I_mullion = sum(item["Ix"] for item in profiles)
    
    supported_height = sum(span_lengths)
    total_height = bottom_cantilever_length + supported_height + top_cantilever_length

    # ============================================================
    # OpenSees 模型建立
    # ============================================================
    ops.wipe()
    ops.model("basic", "-ndm", 2, "-ndf", 3)
    ops.geomTransf("Linear", 1)

    node_tags = []
    element_tags = []
    element_data = {}
    support_node_tags = []
    support_y_coordinates = []
    segment_data = []

    current_node_tag = 1
    current_element_tag = 1
    current_y = 0.0

    ops.node(current_node_tag, 0.0, current_y)
    node_tags.append(current_node_tag)

    # 1. 下部懸臂
    if bottom_cantilever_length > 0.0:
        segment_start_y = current_y
        segment_start_node = current_node_tag
        element_length = bottom_cantilever_length / ELEMENTS_PER_SEGMENT
        segment_element_tags = []
        previous_node_tag = current_node_tag

        for local_ele in range(1, ELEMENTS_PER_SEGMENT + 1):
            current_node_tag += 1
            current_y = segment_start_y + local_ele * element_length
            ops.node(current_node_tag, 0.0, current_y)
            node_tags.append(current_node_tag)
            # 建立元素時代入動態 E_val
            ops.element("elasticBeamColumn", current_element_tag, previous_node_tag, current_node_tag, Area_mullion, E_val, I_mullion, 1)
            element_tags.append(current_element_tag)
            segment_element_tags.append(current_element_tag)
            element_data[current_element_tag] = {"node_i": previous_node_tag, "node_j": current_node_tag, "segment_type": "lower_cantilever", "segment_index": 1}
            previous_node_tag = current_node_tag
            current_element_tag += 1

        segment_data.append({"type": "lower_cantilever", "index": 1, "name": "下部懸臂", "start_y": segment_start_y, "end_y": current_y, "length": bottom_cantilever_length, "start_node": segment_start_node, "end_node": current_node_tag, "element_tags": segment_element_tags})

    support_node_tags.append(current_node_tag)
    support_y_coordinates.append(current_y)

    # 2. 支承間跨度
    for span_index, span_length in enumerate(span_lengths, start=1):
        segment_start_y = current_y
        segment_start_node = current_node_tag
        element_length = span_length / ELEMENTS_PER_SEGMENT
        segment_element_tags = []
        previous_node_tag = current_node_tag

        for local_ele in range(1, ELEMENTS_PER_SEGMENT + 1):
            current_node_tag += 1
            current_y = segment_start_y + local_ele * element_length
            ops.node(current_node_tag, 0.0, current_y)
            node_tags.append(current_node_tag)
            # 建立元素時代入動態 E_val
            ops.element("elasticBeamColumn", current_element_tag, previous_node_tag, current_node_tag, Area_mullion, E_val, I_mullion, 1)
            element_tags.append(current_element_tag)
            segment_element_tags.append(current_element_tag)
            element_data[current_element_tag] = {"node_i": previous_node_tag, "node_j": current_node_tag, "segment_type": "span", "segment_index": span_index}
            previous_node_tag = current_node_tag
            current_element_tag += 1

        support_node_tags.append(current_node_tag)
        support_y_coordinates.append(current_y)
        segment_data.append({"type": "span", "index": span_index, "name": f"第 {span_index} 跨", "start_y": segment_start_y, "end_y": current_y, "length": span_length, "start_node": segment_start_node, "end_node": current_node_tag, "element_tags": segment_element_tags})

    # 3. 頂部懸臂
    if top_cantilever_length > 0.0:
        segment_start_y = current_y
        segment_start_node = current_node_tag
        element_length = top_cantilever_length / ELEMENTS_PER_SEGMENT
        segment_element_tags = []
        previous_node_tag = current_node_tag

        for local_ele in range(1, ELEMENTS_PER_SEGMENT + 1):
            current_node_tag += 1
            current_y = segment_start_y + local_ele * element_length
            ops.node(current_node_tag, 0.0, current_y)
            node_tags.append(current_node_tag)
            # 建立元素時代入動態 E_val
            ops.element("elasticBeamColumn", current_element_tag, previous_node_tag, current_node_tag, Area_mullion, E_val, I_mullion, 1)
            element_tags.append(current_element_tag)
            segment_element_tags.append(current_element_tag)
            element_data[current_element_tag] = {"node_i": previous_node_tag, "node_j": current_node_tag, "segment_type": "upper_cantilever", "segment_index": 1}
            previous_node_tag = current_node_tag
            current_element_tag += 1

        segment_data.append({"type": "upper_cantilever", "index": 1, "name": "頂部懸臂", "start_y": segment_start_y, "end_y": current_y, "length": top_cantilever_length, "start_node": segment_start_node, "end_node": current_node_tag, "element_tags": segment_element_tags})

    # 4. 邊界條件與載重
    for support_node_tag in support_node_tags:
        ops.fix(support_node_tag, 1, 1, 0)

    ops.timeSeries("Linear", 1)
    ops.pattern("Plain", 1, 1)
    q_local_y = -w * WIND_DIRECTION_X
    for element_tag in element_tags:
        ops.eleLoad("-ele", element_tag, "-type", "-beamUniform", q_local_y, 0.0)

    ops.constraints("Plain")
    ops.numberer("RCM")
    ops.system("BandGeneral")
    ops.algorithm("Linear")
    ops.integrator("LoadControl", 1.0)
    ops.analysis("Static")

    if ops.analyze(1) != 0:
        st.error("OpenSees 靜力分析失敗！")
        st.stop()

    # ============================================================
    # 後處理與檢核邏輯
    # ============================================================
    ux = {tag: float(ops.nodeDisp(tag, 1)) for tag in node_tags}
    uy = {tag: float(ops.nodeDisp(tag, 2)) for tag in node_tags}
    rz = {tag: float(ops.nodeDisp(tag, 3)) for tag in node_tags}

    all_y, all_ux, all_moment, all_shear = [], [], [], []
    segment_response = { (item["type"], item["index"]): {"y": [], "ux": [], "moment": [], "shear": []} for item in segment_data }

    for element_sequence, element_tag in enumerate(element_tags):
        data = element_data[element_tag]
        node_i = data["node_i"]
        node_j = data["node_j"]
        y_i, y_j = float(ops.nodeCoord(node_i, 2)), float(ops.nodeCoord(node_j, 2))
        length = y_j - y_i

        local_v_i, local_v_j = -ux[node_i], -ux[node_j]
        theta_i, theta_j = rz[node_i], rz[node_j]
        xi = np.linspace(0.0, 1.0, POSTPROCESS_POINTS_PER_ELEMENT)

        N1 = 1.0 - 3.0 * xi**2 + 2.0 * xi**3
        N2 = xi - 2.0 * xi**2 + xi**3
        N3 = 3.0 * xi**2 - 2.0 * xi**3
        N4 = -xi**2 + xi**3

        # 後處理計算公式代入動態 E_val
        load_bubble = q_local_y * length**4 / (24.0 * E_val * I_mullion) * xi**2 * (1.0 - xi)**2
        local_v = N1 * local_v_i + N2 * length * theta_i + N3 * local_v_j + N4 * length * theta_j + load_bubble
        global_ux = -local_v
        y_values = y_i + xi * length

        curvature = ((-6.0 + 12.0 * xi) / length**2 * local_v_i + (-4.0 + 6.0 * xi) / length * theta_i +
                     (6.0 - 12.0 * xi) / length**2 * local_v_j + (-2.0 + 6.0 * xi) / length * theta_j +
                     q_local_y * length**2 / (24.0 * E_val * I_mullion) * (2.0 - 12.0 * xi + 12.0 * xi**2))
        moment_values = E_val * I_mullion * curvature

        third_derivative = (12.0 / length**3 * local_v_i + 6.0 / length**2 * theta_i - 12.0 / length**3 * local_v_j +
                            6.0 / length**2 * theta_j + q_local_y * length / (24.0 * E_val * I_mullion) * (-12.0 + 24.0 * xi))
        shear_values = E_val * I_mullion * third_derivative

        start_index = 0 if element_sequence == 0 else 1
        all_y.extend(y_values[start_index:].tolist())
        all_ux.extend(global_ux[start_index:].tolist())
        all_moment.extend(moment_values[start_index:].tolist())
        all_shear.extend(shear_values[start_index:].tolist())

        key = (data["segment_type"], data["segment_index"])
        response = segment_response[key]
        segment_start_index = 0 if len(response["y"]) == 0 else 1
        response["y"].extend(y_values[segment_start_index:].tolist())
        response["ux"].extend(global_ux[segment_start_index:].tolist())

    all_y = np.asarray(all_y, dtype=float)
    all_ux = np.asarray(all_ux, dtype=float)
    all_moment = np.asarray(all_moment, dtype=float)
    all_shear = np.asarray(all_shear, dtype=float)

    # 提取極值與發生位置
    max_disp_idx = int(np.argmax(np.abs(all_ux)))
    max_disp, max_disp_y = float(all_ux[max_disp_idx]), float(all_y[max_disp_idx])
    
    max_moment_idx = int(np.argmax(np.abs(all_moment)))
    max_moment_signed = float(all_moment[max_moment_idx])
    max_moment_y = float(all_y[max_moment_idx])
    max_abs_moment = abs(max_moment_signed)
    
    max_shear_idx = int(np.argmax(np.abs(all_shear)))
    max_shear_signed = float(all_shear[max_shear_idx])
    max_shear_y = float(all_y[max_shear_idx])
    max_abs_shear = abs(max_shear_signed)

    # 變形檢核
    span_check_results = []
    for segment in segment_data:
        key = (segment["type"], segment["index"])
        y_vals, ux_vals = np.asarray(segment_response[key]["y"]), np.asarray(segment_response[key]["ux"])
        control_idx = int(np.argmax(np.abs(ux_vals)))
        control_disp, control_y = abs(float(ux_vals[control_idx])), float(y_vals[control_idx])
        
        if segment["type"] == "span":
            allowable, formula = span_allowable_deflection(segment["length"])
        else:
            allowable, formula = 2.0 * segment["length"] / 175.0, "2Lc/175"
            
        span_check_results.append({
            "區段名稱": segment["name"], "長度 (cm)": segment["length"], "最大變形位置 Y": control_y, 
            "實際變形 (cm)": control_disp, "容許公式": formula, "容許限制 (cm)": allowable, "passed": control_disp <= allowable
        })

    deflection_pass = all(item["passed"] for item in span_check_results)

    # 應力檢核
    stress_results = []
    for profile in profiles:
        moment_share = max_abs_moment * profile["Ix"] / I_mullion
        fb = moment_share / profile["Sx"]
        utilization = fb / mat["Fb_allow"]
        stress_results.append({
            "斷面名稱": profile["name"], "Ix (cm^4)": f"{profile['Ix']:.4f}", "Sx (cm^3)": f"{profile['Sx']:.4f}",
            "分配彎矩 (kgf-cm)": f"{moment_share:.4f}", "撓曲應力 fb": f"{fb:.4f}", "使用率": f"{utilization:.4f}",
            "檢核結果": status_text(utilization <= 1.0), "passed": utilization <= 1.0
        })

    stress_pass = all(item["passed"] for item in stress_results)
    overall_pass = stress_pass and deflection_pass

    # ============================================================
    # UI 呈現區：將內容分頁
    # ============================================================
    tab1, tab2, tab3 = st.tabs(["📊 檢核總結 & 應力變形", "📈 結構與內力圖表", "📋 詳細計算與反力數據"])

    with tab1:
        st.subheader("🏁 最終判定結果")
        col1, col2, col3 = st.columns(3)
        col1.metric("強軸應力檢核", "通過" if stress_pass else "不通過", delta="OK" if stress_pass else "FAIL", delta_color="normal" if stress_pass else "inverse")
        col2.metric("容許變形檢核", "通過" if deflection_pass else "不通過", delta="OK" if deflection_pass else "FAIL", delta_color="normal" if deflection_pass else "inverse")
        col3.metric("整體結構判定", "通過" if overall_pass else "不通過", delta="OK" if overall_pass else "FAIL", delta_color="normal" if overall_pass else "inverse")
        
        st.divider()
        st.subheader("1️⃣ 直料風力強軸應力檢核")
        st.write(f"**選用材料:** {mat['name']} (E={E_val:.0f}, Fty={mat['Fty']:.2f}, Ftu={mat['Ftu']:.2f}, 容許 Fb={mat['Fb_allow']:.2f} kgf/cm²)")
        st.write(f"**總面積 A:** {Area_mullion:.4f} cm² | **總慣性矩 Ix:** {I_mullion:.4f} cm⁴")
        st.write(f"**最大絕對彎矩 |M|:** {max_abs_moment:.4f} kgf-cm")
        df_stress = pd.DataFrame(stress_results).drop(columns=["passed"])
        st.dataframe(df_stress, use_container_width=True)

        st.divider()
        st.subheader("2️⃣ 各區段變形檢核明細")
        df_disp = pd.DataFrame(span_check_results)
        df_disp["檢核結果"] = df_disp["passed"].apply(status_text)
        st.dataframe(df_disp.drop(columns=["passed"]), use_container_width=True)

    with tab2:
        st.subheader("圖表視覺化 (OpsVis)")
        fmt_main = {"color": "blue", "linestyle": "solid", "linewidth": 2.0, "marker": "", "markersize": 0}
        fmt_ref = {"color": "blue", "linestyle": "solid", "linewidth": 0.8, "marker": "", "markersize": 0}

        y_coords = [ops.nodeCoord(n)[1] for n in ops.getNodeTags()]
        y_min, y_max = min(y_coords), max(y_coords)
        total_H = y_max - y_min
        load_width = max(total_H * 0.08, 1.0)
        target_w = total_H * 0.4
        
        sfac_M = target_w / max_abs_moment if max_abs_moment > 1e-3 else 1.0
        sfac_V = target_w / max_abs_shear if max_abs_shear > 1e-3 else 1.0

        fig1, ax1 = plt.subplots(figsize=(6, 8))
        fig2, ax2 = plt.subplots(figsize=(6, 8))
        fig3, ax3 = plt.subplots(figsize=(6, 8))
        fig4, ax4 = plt.subplots(figsize=(6, 8))

        # --- 圖1：原始模型 ---
        ax1.set_title("Vertical Multi-span Beam Model")
        opsv.plot_model(ax=ax1, node_supports=True, node_labels=0, element_labels=0)
        
        # 畫均布風載背景塊與箭頭
        rect = patches.Rectangle((-load_width, y_min), load_width, total_H, linewidth=1, edgecolor='steelblue', facecolor='dodgerblue', alpha=0.2)
        ax1.add_patch(rect)
        for y_pos in np.linspace(y_min, y_max, 15):
            ax1.annotate('', xy=(0, y_pos), xytext=(-load_width, y_pos), arrowprops=dict(arrowstyle='->', color='steelblue', lw=1.5))
            
        # 標示各支承 (S1, S2...)
        for support_number, support_y in enumerate(support_y_coordinates, start=1):
            ax1.annotate(
                f"S{support_number}", xy=(0.0, support_y), xytext=(8, 0),
                textcoords="offset points", ha="left", va="center", fontsize=9,
                bbox={"boxstyle": "round,pad=0.18", "facecolor": "white", "alpha": 0.80}
            )

        # 標示下部懸臂自由端與虛線
        if bottom_cantilever_length > 0.0:
            ax1.annotate(
                "Lower free tip", xy=(0.0, 0.0), xytext=(8, -8),
                textcoords="offset points", ha="left", va="top", fontsize=9,
                bbox={"boxstyle": "round,pad=0.18", "facecolor": "white", "alpha": 0.80}
            )
            ax1.axhline(bottom_cantilever_length, linestyle="--", linewidth=1.0, color="gray")
            ax1.text(load_width * 0.15, bottom_cantilever_length / 2.0, "Bottom cantilever", rotation=90, va="center")

        # 標示頂部懸臂自由端與虛線
        if top_cantilever_length > 0.0:
            ax1.annotate(
                "Upper free tip", xy=(0.0, total_height), xytext=(8, 8),
                textcoords="offset points", ha="left", va="bottom", fontsize=9,
                bbox={"boxstyle": "round,pad=0.18", "facecolor": "white", "alpha": 0.80}
            )
            supported_end_y = bottom_cantilever_length + sum(span_lengths)
            ax1.axhline(supported_end_y, linestyle="--", linewidth=1.0, color="gray")
            ax1.text(load_width * 0.15, supported_end_y + top_cantilever_length / 2.0, "Top cantilever", rotation=90, va="center")

        ax1.set_xlim(-load_width * 2, load_width * 2)
        ax1.set_ylim(y_min - total_H * 0.1, y_max + total_H * 0.1)
        ax1.grid(True, linestyle="--", alpha=0.3)

        # --- 圖2：放大變形圖 ---
        target_defo_w = total_H * 0.15
        max_abs_disp = abs(max_disp)
        sfac_defo = target_defo_w / max_abs_disp if max_abs_disp > 1e-5 else 1.0
        ax2.set_title(f"Deformed Shape\nMax |UX| = {max_abs_disp:.4f} cm")
        opsv.plot_defo(sfac=sfac_defo, unDefoFlag=1, node_supports=True, ax=ax2)
        
        # 標示最大變形點
        ax2.plot(max_disp * sfac_defo, max_disp_y, marker="o", markersize=6, zorder=5)
        ax2.annotate(
            f"Max |UX|\nUX={max_disp:.5f} cm\nY={max_disp_y:.2f} cm",
            xy=(max_disp * sfac_defo, max_disp_y), xytext=(10, 8),
            textcoords="offset points", arrowprops={"arrowstyle": "->", "linewidth": 0.8},
            fontsize=9, bbox={"boxstyle": "round,pad=0.20", "facecolor": "white", "alpha": 0.85}
        )
        
        ax2.set_xlim(-target_defo_w * 2, target_defo_w * 2) 
        ax2.set_ylim(y_min - total_H * 0.1, y_max + total_H * 0.1) 
        ax2.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f"{x / sfac_defo:.2f}"))
        ax2.set_xlabel("True Displacement UX (cm)")
        ax2.grid(True, linestyle="--", alpha=0.3)

        # --- 圖3：彎矩圖 ---
        ax3.set_title(f"Bending Moment Diagram\nMax |M| = {max_abs_moment:.1f} kgf-cm")
        opsv.section_force_diagram_2d("M", sfac=sfac_M, nep=17, ax=ax3, fmt_secforce1=fmt_main, fmt_secforce2=fmt_ref, ref_vert_lines=False, end_max_values=False, node_supports=True, number_format=".1f")
        ax3.plot([0, 0], [y_min, y_max], color='black', linewidth=1.5, zorder=3)
        
        # 標示最大彎矩點
        ax3.axhline(max_moment_y, linewidth=0.8, linestyle=":", alpha=0.65)
        ax3.plot(0.0, max_moment_y, marker="o", markersize=6, zorder=5)
        ax3.annotate(
            f"Max |M|\nM={max_moment_signed:.1f} kgf-cm\nY={max_moment_y:.2f} cm",
            xy=(0.0, max_moment_y), xytext=(10, 8),
            textcoords="offset points", arrowprops={"arrowstyle": "->", "linewidth": 0.8},
            fontsize=9, bbox={"boxstyle": "round,pad=0.20", "facecolor": "white", "alpha": 0.85}
        )
        
        ax3.set_xlim(-target_w * 1.5, target_w * 1.5)              
        ax3.set_ylim(y_min - total_H * 0.1, y_max + total_H * 0.1) 
        ax3.grid(True, linestyle="--", alpha=0.3)

        # --- 圖4：剪力圖 ---
        ax4.set_title(f"Shear Force Diagram\nMax |V| = {max_abs_shear:.1f} kgf")
        opsv.section_force_diagram_2d("V", sfac=sfac_V, nep=17, ax=ax4, fmt_secforce1=fmt_main, fmt_secforce2=fmt_ref, ref_vert_lines=False, end_max_values=False, node_supports=True, number_format=".1f")
        ax4.plot([0, 0], [y_min, y_max], color='black', linewidth=1.5, zorder=3)
        
        # 標示最大剪力點
        ax4.axhline(max_shear_y, linewidth=0.8, linestyle=":", alpha=0.65, color="steelblue")
        ax4.plot(0.0, max_shear_y, marker="o", markersize=6, zorder=5)
        ax4.annotate(
            f"Max |V|\nV={max_shear_signed:.1f} kgf\nY={max_shear_y:.2f} cm",
            xy=(0.0, max_shear_y), xytext=(10, 8),
            textcoords="offset points", arrowprops={"arrowstyle": "->", "linewidth": 0.8},
            fontsize=9, bbox={"boxstyle": "round,pad=0.20", "facecolor": "white", "alpha": 0.85}
        )
        
        ax4.set_xlim(-target_w * 1.5, target_w * 1.5)              
        ax4.set_ylim(y_min - total_H * 0.1, y_max + total_H * 0.1) 
        ax4.grid(True, linestyle="--", alpha=0.3)
        for fig in (fig1, fig2, fig3, fig4):
                    fig.tight_layout()

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st.pyplot(fig1)
            st.pyplot(fig3)
        with col_f2:
            st.pyplot(fig2)
            st.pyplot(fig4)
                    
            plt.close('all')
    with tab3:
        st.subheader("⚖️ 支承反力與水平力平衡")
        ops.reactions()
        support_reactions = []
        for i, node_tag in enumerate(support_node_tags, start=1):
            support_reactions.append({
                "支承編號": f"S{i}", "節點 ID": node_tag, "Y (cm)": support_y_coordinates[i - 1],
                "RX (kgf)": float(ops.nodeReaction(node_tag, 1)), "RY (kgf)": float(ops.nodeReaction(node_tag, 2)), "MZ (kgf-cm)": float(ops.nodeReaction(node_tag, 3))
            })
            
        df_react = pd.DataFrame(support_reactions)
        st.table(df_react)
        
        sum_rx = sum(item["RX (kgf)"] for item in support_reactions)
        external_force_x = w * total_height
        balance_error = sum_rx + external_force_x
        
        st.write(f"- 風載總水平力: {external_force_x:.3f} kgf")
        st.write(f"- 支承水平反力總和: {sum_rx:.3f} kgf")
        st.write(f"- 水平力平衡誤差: {balance_error:.3e} kgf")
