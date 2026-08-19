from __future__ import annotations

import streamlit as st


st.set_page_config(
    page_title="帷幕牆結構分析",
    page_icon="🏗️",
    layout="wide",
)

st.title("帷幕牆結構分析")

st.markdown("### 選擇分析項目")

col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.subheader("直料分析")
        st.write("直立多段梁／帷幕直料之風載、變形、彎矩、剪力、應力與支承反力分析。")
        st.page_link("pages/1_直料分析.py", label="進入直料分析", icon="↕️", use_container_width=True)

with col2:
    with st.container(border=True):
        st.subheader("橫料分析")
        st.write("橫料之風壓 Tributary Load、玻璃自重、強弱軸變形與雙軸應力分析。")
        st.page_link("pages/2_橫料分析.py", label="進入橫料分析", icon="↔️", use_container_width=True)

st.info("左側頁面選單也可直接切換直料與橫料分析。")
