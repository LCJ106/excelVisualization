import streamlit as st
import pandas as pd
import os
from pathlib import Path
import openpyxl

# ===================== 【1】页面配置 =====================
st.set_page_config(
    page_title="政策补贴展示系统",
    layout="wide"
)

# ===================== 【2】全局 CSS 样式（写在这里最安全！） =====================
st.markdown("""
<style>
/* 隐藏 Streamlit 默认元素 */
.stDeployButton, #MainMenu, footer {
    visibility: hidden;
}

/* 核心：强制移除 Streamlit 所有限制 → 全屏 */
section[data-testid="stSidebar"] ~ .main {
    max-width: 100% !important;
    padding: 0 !important;
    margin: 0 !important;
}

.block-container {
    max-width: 100% !important;
    padding-left: 0 !important;
    padding-right: 0 !important;
}

/* 外层全屏行区块 */
.row-wrap {
    width: 100%;
    padding: 30px 0;
}


/* 卡片通用样式 */
.card {
    border-radius: 12px;
    padding: 24px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
    margin:16px auto !important;
    max-width:1000px;
    width: 70%;            /* 小屏幕自动缩到 95% 宽度 */
}

/* 让页面大标题居中 */
h1 {
    text-align: center !important;
}

/* 长文本展示框（和卡片完全融合） */
.text-area-custom {
    padding: 12px;
    border-radius: 8px;
    font-size: 15px;
    line-height: 1.6;
    max-height: 120px;
    overflow-y: auto;
    margin-top: 6px;
}
</style>
""", unsafe_allow_html=True)

# ===================== 【3】卡片函数（可读性极高） =====================
def show_card(policy_row, dept_name, is_blue=True):
    # 配色
    if is_blue:
        bg_color = "#f2f8fc"
        border_color = "#3674b9"
        title_color = "#285b94"
        box_bg = "#e6f1fc"
        # 卡片所在的层背景
        row_bg = "#f5f5f5"
    else:
        bg_color = "#f3faf6"
        border_color = "#38a169"
        title_color = "#2d8556"
        box_bg = "#e6f5ed"
        row_bg = "transparent"

    print(policy_row)
    print(type(policy_row))
    policy_content = policy_row["政策内容"]
    policy_file = policy_row["政策文件"]
    policy_object = policy_row["政策适用对象"]
    allowance_standard = policy_row["补贴标准/金额"]
    capital_source = policy_row["资金来源"]
    implement_period = policy_row["实施期限"]
    # 卡片整体（一个完整 div）
    # <div style="font-size:24px; font-weight:bold; color:{title_color}; margin-bottom:18px;">{dept_name}</div>
    st.markdown(f"""
    <div class="row-wrap" style="background:{row_bg};">
        <div class="card" style="background:{bg_color}; ">
            <div style="margin:8px 0; font-size:16px;"><b>政策内容：</b> {policy_content}</div>
            <div style="margin:8px 0; font-size:16px;"><b>政策文件：</b> {policy_file}</div>
            <div style="margin-top:16px;">
                <div style="font-size:16px; font-weight:bold;">政策适用对象：</div>
                <div class="text-area-custom" style="background:{box_bg};">
                    {policy_object}</div>
            </div>
            <div style="margin-top:16px;">
                <div style="font-size:16px; font-weight:bold;">补贴标准/金额：</div>
                <div class="text-area-custom" style="background:{box_bg};">{allowance_standard}</div>
            </div>
            <div style="margin-top:16px; font-size:16px;"><b>资金来源：</b> {capital_source}</div>
            <div style="margin-top:8px; font-size:16px;"><b>实施期限：</b> {implement_period}</div>
        </div></div>
    """, unsafe_allow_html=True)

# ===================== 【4】主页面逻辑 =====================
def show_visual(df):
    st.title("🏛️ 各科室政策补贴展示页面")
    # policy
    sheets = list(df.keys())
    print(sheets)
    policy = df[sheets[0]]
    policy_group = policy.groupby("科室/单位")
    for id, (dept_name, group) in enumerate(policy_group):
        st.markdown(f'<div id="dept_{dept_name}"></div>', unsafe_allow_html=True)
        for index, row in group.iterrows():
            show_card(row, dept_name, is_blue=(id % 2 == 0))
    departments = policy["科室/单位"].drop_duplicates()

    # # 自动交替颜色
    # for i, dept in enumerate(departments):
    #     show_card(dept, is_blue=(i % 2 == 0))

DATA_SAVE_FILE = "latest_data.xlsx"
# 优先加载已保存的最新表格
df = None
BASE_DIR = Path(__file__).parent

file_path = BASE_DIR / DATA_SAVE_FILE
print("文件是否存在：", os.path.exists(file_path))
if file_path.exists():
    df = pd.read_excel(file_path, header=1, sheet_name=None)
    # st.success("已自动加载当前最新版本表格数据")
# 有数据就展示+做可视化
if df is not None:
    # st.subheader("数据明细")
    # st.dataframe(df, use_container_width=True)
    # 下方直接粘贴你的ECharts可视化代码即可
    show_visual(df)

