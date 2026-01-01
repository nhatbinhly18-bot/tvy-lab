import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime

# ==========================================
# 1. 网页基础配置
# ==========================================
st.set_page_config(page_title="体卫艺办公助手", page_icon="🏫", layout="centered")

# --- 🎨 全局 CSS 美化 ---
st.markdown("""
<style>
    /* 全局字体与背景 */
    html, body, [class*="css"] { font-family: 'Helvetica Neue', 'PingFang SC', 'Microsoft YaHei', sans-serif; }
    .stApp { background-image: linear-gradient(135deg, #fdfbfb 0%, #ebedee 100%); }
    
    /* 侧边栏 */
    [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e1e4e8; box-shadow: 2px 0 10px rgba(0,0,0,0.01); }
    
    /* 卡片容器 */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        background-color: #ffffff; border: 1px solid #e1e4e8 !important;
        border-radius: 12px !important; box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        padding: 1.5rem !important;
    }
    
    /* 按钮美化 */
    div.stButton > button { border-radius: 20px !important; font-weight: 600 !important; }
    div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #1976d2 0%, #1565c0 100%) !important; color: white !important;
    }
    
    /* 🚀 简报助手专用传送门按钮 */
    div.stButton > a { 
        display: inline-block; width: 100%; text-align: center;
        padding: 1.2rem !important; font-size: 1.3rem !important;
        font-weight: bold !important; border-radius: 12px !important;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white !important; text-decoration: none;
        box-shadow: 0 4px 15px rgba(118, 75, 162, 0.4);
        transition: transform 0.2s;
    }
    div.stButton > a:hover { transform: scale(1.02); }
    
    /* 隐藏页眉 */
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    @media (min-width: 769px) { header {visibility: hidden;} }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 核心配置区
# ==========================================
CONTACT_PASSWORD = "lhjy" 
MY_API_KEY = "sk-dzsawqzsktjximglmkzyezbtyhqbysvenoxublemcgertlqp"
BASE_URL = "https://api.siliconflow.cn/v1"

# 状态初始化
if "contacts_authenticated" not in st.session_state: st.session_state.contacts_authenticated = False
if "parseddata_doc" not in st.session_state: st.session_state.parseddata_doc = None
if "step" not in st.session_state: st.session_state.step = 1
if "original_input" not in st.session_state: st.session_state.original_input = ""

# ==========================================
# 3. 侧边栏导航
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/school-building.png", width=60)
    st.header("体卫艺办公助手")
    st.success("● 系统运行正常") 
    st.markdown("---")
    
    mode = st.radio("功能导航：", [
        "✨ 体卫艺简报助手", 
        "📝 领导公务单生成", 
        "🔍 学校查号台"
    ])
    
    st.markdown("---")
    st.caption("Ver 3.5 | 龙华区教育局体卫艺科")
    st.caption("维护者：孙沛")

# ==========================================
# 模块一：✨ 体卫艺简报助手 (Portal Mode)
# ==========================================
if mode == "✨ 体卫艺简报助手":
    st.markdown("# ✨ 体卫艺简报助手")
    st.caption("DeepSeek-V3 满血版引擎驱动")
    
    with st.container(border=True):
        st.subheader("📝 智能润色专家")
        st.info("👇 **请直接发送：会议通知 + 参会名单 + 杂乱语音稿**")
        st.write("") 
        
        # ✅ 这里填入测试通过的 DeepSeek 链接
        DEEPSEEK_LINK = "https://www.coze.cn/store/agent/7587031903597985832?from=store_search_suggestion&bid=6ilacph8g3009"
        
        st.link_button("🚀 启动 DeepSeek 助手", DEEPSEEK_LINK)
        
        st.write("")
        st.markdown("""
        <small style='color:gray'>
        💡 <b>使用说明：</b><br>
        1. 点击按钮将跳转至体卫艺专用 AI 页面。<br>
        2. 支持超长文本处理与 DeepSeek 深度思考。<br>
        3. <b>无需配置 Key，永久免费使用。</b>
        </small>
        """, unsafe_allow_html=True)

# ==========================================
# 模块二：📝 领导公务单自动生成器
# ==========================================
elif mode == "📝 领导公务单生成":
    st.markdown("# 📋 领导公务单生成器")
    st.info("💡 **提示：** 请一次性说清：时间、地点、会议名称、人数、对接人、领导、参加部门及议程。")

    if st.session_state.step == 1:
        with st.container(border=True):
            st.subheader("1️⃣ 描述活动信息")
            user_input = st.text_area("请在此输入...", height=150, placeholder="例如：明天上午10点在二楼多功能厅有个座谈会...", label_visibility="collapsed")
        
        st.write("") 
        if st.button("✨ 立即生成 Word", type="primary", use_container_width=True):
            if not user_input:
                st.warning("⚠️ 内容不能为空")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                current_date = datetime.now().strftime("%Y年%m月%d日")
                
                with st.spinner("🤖 正在解析要素并润色..."):
                    prompt = f"""
                    你现在是龙华教育局资深笔杆子。请根据用户输入解析公文要素。
                    当前日期：{current_date}。
                    用户输入：{user_input}
                    标准人名库：杨灵芝、尹泽利、文良方、孙沛、刘冰、杨帆、陈海万、路旭阳、王轩、王燕、李桂情、甘月琴、方梦懿、吴正光、李长生、梁永誉、刘喜菊
                    要求：JSON格式输出，字段包含 title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。
                    """
                    try:
                        res = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct", 
                            messages=[{"role": "user", "content": prompt}], 
                            response_format={'type': 'json_object'}
                        )
                        st.session_state.parseddata_doc = json.loads(res.choices[0].message.content)
                        st.session_state.step = 2
                        st.rerun()
                    except Exception as e:
                        st.error(f"解析出错：{e}")

    elif st.session_state.step == 2:
        d = st.session_state.parseddata_doc
        with st.container(border=True):
            st.subheader("2️⃣ 预览与微调")
            t = st.text_input("活动名称", d.get("title", ""))
            c = st.text_area("背景理由", d.get("content", ""), height=100)
            a = st.text_area("议程", d.get("agenda", ""), height=100)
            
            c1, c2 = st.columns(2)
            with c1: tm = st.text_input("时间", d.get("time", ""))
            with c2: pl = st.text_input("地点", d.get("place", ""))
            
            c3, c4 = st.columns(2)
            with c3: dist_l = st.text_input("区领导", d.get("dist_leader", ""))
            with c4: bur_l = st.text_input("局领导", d.get("bur_leader", ""))
            
            st.write("")
            col_b1, col_b2 = st.columns([1, 2])
            with col_b1:
                if st.button("⬅️ 返回修改"):
                    st.session_state.step = 1
                    st.rerun()
            with col_b2:
                try:
                    tpl = DocxTemplate("申报单模板.docx")
                    tpl.render(d)
                    bio = io.BytesIO()
                    tpl.save(bio)
                    fname = f"{datetime.now().strftime('%m%d')}_体卫艺_{t}.docx"
                    st.download_button("💾 导出 Word", bio.getvalue(), fname, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"生成失败：{e}")

# ==========================================
# 模块三：🔍 学校查号台
# ==========================================
else:
    st.markdown("# 🔍 龙华学校查号台")
    
    if not st.session_state.contacts_authenticated:
        with st.container(border=True):
            pwd = st.text_input("请输入授权密码", type="password")
            if st.button("验证登录", type="primary"):
                if pwd == CONTACT_PASSWORD:
                    st.session_state.contacts_authenticated = True
                    st.rerun()
                else:
                    st.error("密码错误")
        st.stop()

    @st.cache_data
    def load_contacts():
        try: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='utf-8-sig').fillna('无')
        except: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='gbk').fillna('无')

    df = load_contacts()
    with st.container(border=True):
        q = st.text_input("🔎 搜索学校或姓名", placeholder="支持模糊搜索...")
    if q:
        mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
        st.dataframe(df[mask], use_container_width=True, hide_index=True)
