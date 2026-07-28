import streamlit as st

st.set_page_config(
    page_title="汇发数据管理系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2.2rem; padding-bottom: 3rem;}
    .hero {text-align:center; padding:26px 20px 24px; border-bottom:1px solid #e7ebf0; margin-bottom:26px;}
    .hero-title {font-size:30px; font-weight:800; color:#16243a; line-height:1.5;}
    .hero-subtitle {font-size:16px; color:#6b7280; margin-top:8px;}
    .card {border:1px solid #dfe5ec; border-radius:16px; padding:22px; background:#fbfcfe; min-height:190px;}
    .card-title {font-size:20px; font-weight:750; color:#172033; margin-bottom:10px;}
    .card-text {font-size:15px; color:#5d6675; line-height:1.8;}
    .note {margin-top:28px; padding:16px 18px; border-radius:12px; background:#f4f7fb; color:#4b5563; line-height:1.8;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <div class="hero-title">📊 汇发数据管理系统</div>
      <div class="hero-subtitle">销售数据更新 · 服务站汇总 · 应收账款跟进</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.subheader("请选择左侧功能")

col1, col2, col3 = st.columns(3, gap="large")
with col1:
    st.markdown(
        """
        <div class="card">
          <div class="card-title">📊 销售数据更新</div>
          <div class="card-text">上传康总基础数据、销货单、销售订单和应收总账，完成客户匹配并生成最新乡镇明细和总汇总。</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with col2:
    st.markdown(
        """
        <div class="card">
          <div class="card-title">🏪 服务站数据更新</div>
          <div class="card-text">从最新康总数据中提取全部 F 开头服务站，保留原顺序和样式，重新计算服务站合计。</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with col3:
    st.markdown(
        """
        <div class="card">
          <div class="card-title">💰 应收账款更新</div>
          <div class="card-text">筛选正数应收客户，并保留负责人、联系日期、付款状态和下次跟进日期等历史信息。</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    '<div class="note">数据均在当前运行环境的内存中处理。生成文件后请及时下载保存。左侧菜单可以在三个功能之间切换。</div>',
    unsafe_allow_html=True,
)
