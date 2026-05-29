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

# ===================== 【2】全局 CSS 样式=====================
st.markdown("""
<style>
/* 隐藏 Streamlit 默认元素 */
#MainMenu, div[data-testid="stToolbar"] .stDeployButton, footer {
    visibility: hidden !important;
}

/* 推荐做法：让整个 Header 固定并变色 */
header[data-testid="stHeader"] {
    background-color: #222222 !important;
    top: 0 !important;
    z-index: 9999 !important;
    height: 3rem !important;      /* 压缩高度，原值 3.75rem，可按需微调 */
    min-height: 3rem !important;  /* 必须同步修改，否则 min-height 会撑住不让它变小 */
    position: fixed !important;
    
}

/* 2. 隐藏 Header 里面原本的空白区域和部署按钮（如果有的话），但不影响导航栏 */
[data-testid="stHeader"] > div:not([data-testid="stTopNav"]) {
    /* 如果需要隐藏某些特定元素，可以在这里做精细控制 */
}

/* 1. 去除单个导航链接的椭圆边框和默认背景 */
[data-testid="stTopNavLink"] {
    border: none !important;             /* 去除椭圆边框 */
    box-shadow: none !important;         /* 去除可能存在的阴影 */
    background-color: transparent !important; /* 背景透明 */
    border-radius: 0 !important;         /* 确保没有圆角 */
    padding: 6px 6px !important;       /* 你的自定义间距 */
    font-weight: normal !important;
}

/* 2. 强制修改导航链接内部的文字颜色 */
[data-testid="stTopNavLink"] [data-testid="stMarkdownContainer"] {
    color: white !important;             /* 强制白色字体 */
    font-size: 15px !important;          /* 你的自定义字号 */
    text-align: center !important;
}

/* 3. 鼠标悬停时的效果（可选，但建议加上，否则鼠标放上去没反应） */
[data-testid="stTopNavLink"]:hover {
    border: none !important;
    box-shadow: none !important;
    background-color: rgba(255, 255, 255, 0.1) !important; /* 悬停时轻微变亮 */
    border-radius: 0 !important;
}

/* 4. 当前被选中的页面高亮（非常重要，否则不知道选的是哪个） */
[data-testid="stTopNavLink"][aria-current="page"] {
    background-color: #080808 !important;
}


.block-container {
    max-width: 100% !important;
    padding-left: 0 !important;
    padding-right: 0 !important;
    padding-top: 1rem !important;
}

/* 外层全屏行区块 */
.row-wrap {
    /*width: 100%;*/
    padding: 30px 0;
     flex-grow: 0;        /* 不要拉伸！如果一行有剩余空间，不要把卡片撑大，保持原样 */
    flex-shrink: 0;      /* 不要压缩！保证卡片不会被压扁变形 */
    flex-basis: 600px;   /* 理想宽度 300px。这是触发换行的关键依据！ */
    
    /* --- 2. 兼容小屏幕设备 (手机/平板) --- */
    min-width: 280px;    /* 最小宽度，允许在极窄屏幕下稍微缩一点点 */
    max-width: 80%;     /* 关键！如果屏幕宽度小于 300px，卡片最多占满屏幕，绝不超出产生横向滚动条 */
    
    /* --- 3. 盒模型与美化 --- */
    width: 600px;        /* 显式设定宽度，双重保险 */
    box-sizing: border-box; /* 确保 padding 和 border 包含在 300px 内，不会撑破布局 */
    height: 100%;     /* 等高 */
}

/*侧边栏*/
section[data-testid="stSidebar"] {
    top : 3rem;
    width : 260px !important;
}
[data-testid="stSidebarCollapseButton"] {
        display: none !important; 
}
[data-testid="stSidebarHeader"] {
    height : 0; 
}
/* 选中侧边栏内的所有 a 标签 */
section[data-testid="stSidebar"] a {
    text-decoration: none;  /* 去除下划线 */
    color: black;           /* 修改字体颜色 */
    padding-left: 16px ; 
}

@media (max-width: 768px) {
    section[data-testid="stSidebar"] a {
    padding-left: 10px !important;
    }
}
    
/* 可选：鼠标悬停时的颜色 */
section[data-testid="stSidebar"] a:hover {
    color: gray;  
}

[data-testid="stHorizontalBlock"]{
    padding: 50px 60px 60px 60px; /* 上下左右都留空隙 */
}

/* 卡片通用样式 */
.card {
    flex-grow: 0;        /* 不要拉伸！如果一行有剩余空间，不要把卡片撑大，保持原样 */
    flex-shrink: 0;      /* 不要压缩！保证卡片不会被压扁变形 */
    flex-basis: 460px;   /* 理想宽度 这是触发换行的关键依据！ */
    
    /* --- 2. 兼容小屏幕设备 (手机/平板) --- */
    min-width: 280px;    /* 最小宽度，允许在极窄屏幕下稍微缩一点点 */
    
    /* --- 3. 盒模型与美化 --- */
    width: 460px;        /* 显式设定宽度，双重保险 */
    box-sizing: border-box; /* 确保 padding 和 border 包含在 300px 内，不会撑破布局 */
    height: 100%;     /* 等高 */
    
    border-radius: 12px;
    padding: 36px 10px 24px 24px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08); 
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

    policy_content = policy_row["政策内容"]
    policy_file = policy_row["政策文件"]
    policy_object = policy_row["政策适用对象"]
    allowance_standard = policy_row["补贴标准/金额"]
    capital_source = policy_row["资金来源"]
    implement_period = policy_row["实施期限"]
    # 卡片整体（一个完整 div）
    # <div class="row-wrap" style="background:{row_bg};"></div>
    st.markdown(f"""
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
        </div>
    """, unsafe_allow_html=True)

def page_home():
    st.markdown('<div id="top"></div>', unsafe_allow_html=True)
    st.title("🏛️ 各科室政策补贴展示页面")
    df = st.session_state.df
    # policy
    sheets = list(df.keys())
    policy = df[sheets[0]]
    departments = policy["科室/单位"].drop_duplicates()
    # 导航栏
    # st.sidebar.markdown("## [ 📜 科室政策](#top)", unsafe_allow_html=True)
    st.sidebar.markdown("## 📜 科室导航")
    print("科室导航")
    icons = ["📁", "📋", "📊", "📍", "📁", "📎", "✅", "📌", "📋"]
    policy_group = policy.groupby("科室/单位")
    for id, (dept_name, group) in enumerate(policy_group):
        icon = icons[id % len(icons)]
        # div id，用于绑定; #为函数要求的前缀在 div id之前
        sid_link = f"#dept_{dept_name}"
        st.sidebar.markdown(f"[{icon}&nbsp;{dept_name}]({sid_link})")
        st.markdown(f'<div id="dept_{dept_name}"></div>', unsafe_allow_html=True)
        with st.container(horizontal=True, gap="medium"):
            for index, row in group.iterrows():
                show_card(row, dept_name, is_blue=(id % 2 == 0))

def page_second():
    st.markdown('<div id="top"></div>', unsafe_allow_html=True)

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
if df is not None and "df" not in st.session_state:
    st.session_state.df = df  # 你的表格


pages = [
        st.Page(page_home, title="🏠 首页", default=True),
        st.Page(page_second, title="🔖 标签"),
]
# 官方导航
pg = st.navigation(pages, position="top")
pg.run()