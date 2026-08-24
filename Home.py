from __future__ import annotations

import streamlit as st


# ============================================================
# 頁面基本設定
# ============================================================

st.set_page_config(
    page_title="帷幕牆結構分析",
    page_icon="🏗️",
    layout="wide",
)


# ============================================================
# 首頁
# ============================================================

def show_home() -> None:
    st.title("帷幕牆設計與結構分析")
    st.caption("設計風壓、玻璃、直料與橫料整合分析平台")

    st.markdown("---")
    st.subheader("選擇分析項目")

    # --------------------------------------------------------
    # 第一排：設計風壓、玻璃
    # --------------------------------------------------------
    col1, col2 = st.columns(2, gap="large")

    with col1:
        with st.container(border=True):
            st.subheader("🌬️ 設計風壓")
            st.write(
                "依工址基本設計風速、地況、用途係數、地形係數、"
                "內外風壓係數及 Zone 4 / Zone 5 分區，"
                "計算帷幕牆局部構材與外部被覆物之設計風壓。"
            )

            if st.button(
                "進入設計風壓分析",
                key="go_wind_pressure",
                use_container_width=True,
                type="primary",
            ):
                st.switch_page(wind_pressure_page)

    with col2:
        with st.container(border=True):
            st.subheader("🪟 玻璃分析")
            st.write(
                "進行玻璃面材相關設計與檢核，"
                "依玻璃尺寸、厚度、設計風壓等條件計算所需結果，"
                "並提供玻璃面材設計資訊。"
            )

            if st.button(
                "進入玻璃分析",
                key="go_glass",
                use_container_width=True,
                type="primary",
            ):
                st.switch_page(glass_page)

    st.markdown("")

    # --------------------------------------------------------
    # 第二排：直料、橫料
    # --------------------------------------------------------
    col3, col4 = st.columns(2, gap="large")

    with col3:
        with st.container(border=True):
            st.subheader("↕️ 直料分析")
            st.write(
                "直立多段構件之風壓分析，可建立多跨、懸臂、"
                "Hinge、Pin 與 Free end，並進行位移、剪力、彎矩、"
                "支承反力、強軸應力及容許變形檢核。"
            )

            if st.button(
                "進入直料分析",
                key="go_mullion",
                use_container_width=True,
                type="primary",
            ):
                st.switch_page(mullion_page)

    with col4:
        with st.container(border=True):
            st.subheader("↔️ 橫料分析")
            st.write(
                "簡支橫料之風壓與玻璃自重分析，可處理單片玻璃、"
                "上下同尺寸玻璃及上下玻璃獨立輸入，並進行強弱軸變形、"
                "彎矩、剪力與雙軸應力檢核。"
            )

            if st.button(
                "進入橫料分析",
                key="go_transom",
                use_container_width=True,
                type="primary",
            ):
                st.switch_page(transom_page)

    st.markdown("---")
    st.info(
        "也可以直接使用左側導覽列切換："
        "首頁、設計風壓、玻璃分析、直料分析與橫料分析。"
    )


# ============================================================
# 頁面定義
# ============================================================

home_page = st.Page(
    show_home,
    title="首頁",
    icon="🏠",
    default=True,
)

wind_pressure_page = st.Page(
    "pages/01_設計風壓.py",
    title="設計風壓",
    icon="🌬️",
)

glass_page = st.Page(
    "pages/02_玻璃分析.py",
    title="玻璃分析",
    icon="🪟",
)

mullion_page = st.Page(
    "pages/03_直料分析.py",
    title="直料分析",
    icon="↕️",
)

transom_page = st.Page(
    "pages/04_橫料分析.py",
    title="橫料分析",
    icon="↔️",
)


# ============================================================
# 左側導覽
# ============================================================

navigation = st.navigation(
    [
        home_page,
        wind_pressure_page,
        glass_page,
        mullion_page,
        transom_page,
    ],
    position="sidebar",
)

navigation.run()
