import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime

# 1. 网页基础配置
st.set_page_config(page_title="体卫艺办公助手", page_icon="📋", layout="centered")

# --- 🎨 旗舰级商业 SaaS 设计语言 ---
st.markdown("""
<style>
    /* 全局重置与字体 */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "PingFang SC", "Microsoft YaHei", sans-serif;
        color: #1f2937;
    }
    
    /* 页面背景 - 深邃商务灰 */
    .stApp {
        background-color: #f8fafc;
        background-image: radial-gradient(#cbd5e1 0.5px, transparent 0.5px);
        background-size: 24px 24px;
    }

    /* 侧边栏样式 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    
    /* 隐藏多余元素 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    @media (min-width: 769px) { header {visibility: hidden;} }
    @media (max-width: 768px) { 
        header {visibility: visible !important; background-color: transparent !important;}
        .block-container { padding-top: 2rem !important; }
    }

    /* ---------------- 👑 高级感：进化版超级卡片 ---------------- */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        background-color: #ffffff;
        border: 1px solid #d1d5db !important; /* 强化边框 */
        border-radius: 16px !important; /* 更圆润 */
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04) !important; /* 深度悬浮阴影 */
        padding: 2.5rem !important;
        margin-top: 1rem;
    }

    /* 顶部标题区 - 商业软件 Header */
    .saas-header {
        background: linear-gradient(135deg, #eff6ff 0%, #ffffff 100%);
        padding: 2rem;
        border-left: 6px solid #2563eb; /* 侧边品牌蓝条 */
        border-radius: 8px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border: 1px solid #e5e7eb;
    }

    h1 {
        font-size: 2rem !important;
        font-weight: 850 !important;
        color: #1e3a8a !important; /* 深蓝 */
        letter-spacing: -0.03em;
        margin: 0 !important;
    }
    
    /* 自定义徽章 */
    .saas-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 6px;
        font-size: 0.7rem;
        font-weight: 700;
        text-transform: uppercase;
        background-color: #2563eb;
        color: white;
        margin-bottom: 0.75rem;
    }
    
    /* ---------------- 交互组件 ---------------- */
    /* 输入框 - 极简白风格 */
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 10px !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #fcfcfc !important;
        transition: all 0.2s;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #3b82f6 !important;
        background-color: #ffffff !important;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1) !important;
    }
    
    /* 按钮 - 扁平商务 */
    div.stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 0.6rem 1.2rem !important;
    }
    
    /* 进度条美化 */
    .step-container {
        display: flex;
        justify-content: space-between;
        margin: 2rem 0;
        padding: 0 2rem;
    }
    .step-item {
        text-align: center;
        flex: 1;
    }
    .step-circle {
        width: 36px;
        height: 36px;
        border-radius: 10px; /* 方圆感更现代 */
        background-color: #f1f5f9;
        color: #94a3b8;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
        font-size: 1.1rem;
        margin: 0 auto 0.5rem;
        transition: all 0.3s;
    }
    .step-active .step-circle {
        background-color: #2563eb;
        color: white;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
    }
    .step-text {
        font-size: 0.85rem;
        font-weight: 600;
        color: #64748b;
    }
    .step-active .step-text {
        color: #1e3a8a;
    }

    @media (max-width: 768px) {
        div[data-testid="stVerticalBlockBorderWrapper"] > div { padding: 1.2rem !important; }
        .step-container { padding: 0; }
    }
</style>
""", unsafe_allow_html=True)

# --- 🔒 通讯录专属密码 ---
CONTACT_PASSWORD = "lhjy" 
MY_API_KEY = "sk-dzsawqzsktjximglmkzyezbtyhqbysvenoxublemcgertlqp"
BASE_URL = "https://api.siliconflow.cn/v1"

# 初始化状态
if "contacts_authenticated" not in st.session_state:
    st.session_state.contacts_authenticated = False
if "parseddata_doc" not in st.session_state:
    st.session_state.parseddata_doc = None
if "step" not in st.session_state:
    st.session_state.step = 1
if "polished_text" not in st.session_state:
    st.session_state.polished_text = None
if "original_input" not in st.session_state:
    st.session_state.original_input = ""

# 3. 侧边栏导航
with st.sidebar:
    st.header("⚙️ 体卫艺办公助手")
    st.success("● AI 核心引擎已连接") 
    
    st.markdown("---")
    mode = st.radio("功能切换：", ["📝 领导公务单自动生成器", "🔍 龙华学校查号台"])
    st.markdown("---")
    st.info("**💡 帮助中心**\n\n如需支持，请联系体卫艺科。")
    st.caption("维护者：孙沛 | 龙华区教育局体卫艺专用")
    st.write("") 
    if st.button("🔒 安全退出"):
        st.session_state.contacts_authenticated = False
        st.session_state.parseddata_doc = None
        st.rerun()

# ----------------- 模块一：领导公务单生成器 (企业版) -----------------
if mode == "📝 领导公务单自动生成器":
    
    # 顶部导航指引
    st.caption("↖️ **导航：** 点击左上角 **>** 可切换功能")
    
    # 动态渲染进度条 HTML
    s1_class = "step-active" if st.session_state.step == 1 else ""
    s2_class = "step-active" if st.session_state.step == 2 else ""
    
    step_html = f"""
    <div class="step-container">
        <div class="step-item {s1_class}">
            <div class="step-circle">1</div>
            <div class="step-text">智能填报</div>
        </div>
        <div class="step-item {s2_class}">
            <div class="step-circle">2</div>
            <div class="step-text">确认与生成</div>
        </div>
    </div>
    """
    
    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #f0f7ff 0%, #ffffff 100%);
        padding: 2rem;
        border-radius: 12px;
        border-left: 8px solid #2563eb;
        border-right: 1px solid #e2e8f0;
        border-top: 1px solid #e2e8f0;
        border-bottom: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 2rem;
    ">
        <div class="saas-badge" style="background-color: #2563eb; color: white; padding: 2px 8px; border-radius: 4px; font-size: 10px; font-weight: bold; margin-bottom: 10px; display: inline-block;">
            ENTERPRISE EDITION V2.5
        </div>
        <h1 style="margin: 0; color: #1e3a8a; font-size: 24px; font-weight: 800;">📋 体卫艺领导公务单自动生成器</h1>
        <p style="color: #475569; margin-top: 8px; font-size: 14px; font-weight: 500;">
            龙华教育局政务专用 · 智能公文系统 | <span style="color: #2563eb;">Tech Support by Peipei</span>
        </p>
    </div>
    {step_html}
    """, unsafe_allow_html=True)
    
    # --- BLUE INFO BOX ---
    st.info("""
    **💡 智能指令范例：**
    “明天上午10点在二楼多功能厅有个生涯教育座谈会，大概20人，孙沛对接，1小时，邀请灵芝主任参加。”
    """)

    # --- Step 1: Input ---
    if st.session_state.step == 1:
        with st.container(border=True):
            st.subheader("✍️ 描述活动")
            st.caption("请直接粘贴语音转文字内容，AI 助手将自动提取关键要素。")
            
            user_input = st.text_area(
                "input_area", 
                height=160, 
                placeholder="在此输入...", 
                key="input_doc", 
                label_visibility="collapsed"
            )
        
        st.write("")
        col_btn, _ = st.columns([1, 0.2])
        if col_btn.button("✨ 开始智能分析", type="primary", use_container_width=True):
            if not user_input:
                st.warning("⚠️ 请输入内容")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                st.session_state.original_input = user_input
                current_date_str = datetime.now().strftime("%Y年%m月%d日")
                weekday = datetime.now().strftime("%w")
                
                with st.spinner("🔄 AI 正在分析语义并生成公文..."):
                    name_corrections = {
                        "林芝": "杨灵芝", "杨林芝": "杨灵芝", "陈海湾": "陈海万", "陈海完": "陈海万",
                        "尹泽力": "尹泽利", "尹则利": "尹泽利", "文量方": "文良方", "温良方": "文良方",
                        "刘兵": "刘冰", "梁永育": "梁永誉", "方梦仪": "方梦懿"
                    }
                    full_prompt = f"""
                    你现在是龙华教育局资深笔杆子。请根据以下用户的大白话描述，解析出公文要素，并对【理由背景】和【议程】部分进行专业润色。
                    【当前日期参考】：今天是 {current_date_str} (星期{weekday})。
                    【用户输入】：{user_input}
                    【解析与润色要求】：
                    1. **人名纠错**：如果用户输入的人名与标准人名库相似（如"林芝"应为"杨灵芝"），请自动纠正。
                    2. **content (理由背景)**：将用户的背景描述转化为"为落实...要求，推进...发展"等公文规范用语。
                    3. **agenda (详细议程)**：**固定输出以下三项**：["专题汇报", "座谈交流", "领导讲话"]。
                    4. **time (时间)**：必须将"明天"等相对时间**计算为具体的年月日**。
                    5. **duration (时长)**：统一计算为"X小时"。
                    6. **contact (公务对接人)**：提取人名，若无则默认为"孙沛"。
                    7. **dist_leader/bur_leader**：准确提取拟请出席的领导，不加部门前缀。
                    必须以 JSON 格式输出: title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。
                    """
                    try:
                        chat_completion = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct", 
                            messages=[{"role": "user", "content": full_prompt}], 
                            response_format={'type': 'json_object'},
                            timeout=30 
                        )
                        result = json.loads(chat_completion.choices[0].message.content)
                        # 字段健壮性
                        for f in ["title", "content", "agenda", "time", "contact"]:
                            if f not in result: result[f] = ""
                        st.session_state.parseddata_doc = result
                        st.session_state.step = 2
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 解析出错：{str(e)}")

    # --- Step 2: Confirmation ---
    elif st.session_state.step == 2 and st.session_state.parseddata_doc:
        d = st.session_state.parseddata_doc
        
        with st.container(border=True):
            st.subheader("📝 确认详情")
            
            # Form Layout
            t = st.text_input("活动名称", d.get("title", ""))
            c = st.text_area("背景/理由", d.get("content", ""), height=100)
            
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list):
                agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            if not agenda_val: agenda_val = "1. 专题汇报\n2. 座谈交流\n3. 领导讲话"
            a = st.text_area("会议议程", agenda_val, height=120)
            
            st.write("---")
            c1, c2 = st.columns(2)
            with c1:
                tm = st.text_input("开始时间", d.get("time", ""))
                dr_val = str(d.get("duration", "1小时"))
                if "小时" not in dr_val: dr_val += "小时"
                dr = st.text_input("会议时长", dr_val)
            with c2:
                st.text_input("仅供参考", "时间不可调整", disabled=True)
                ct = st.text_input("对接人", d.get("contact", "孙沛"))
                
            c3, c4, c5 = st.columns([2, 1, 1])
            with c3: pl = st.text_input("地点", d.get("place", ""))
            with c4: nm = st.text_input("人数", d.get("num", ""))
            with c5: pj = st.selectbox("投影", ["☑使用", "☐不使用"], index=0 if "是" in str(d.get("projector")) else 1)
            
            st.write("---")
            st.markdown("**领导出席**")
            dl = st.text_input("区领导", d.get("dist_leader", ""))
            bl = st.text_input("局领导", d.get("bur_leader", ""))
            oth = st.text_input("建议参加部门", d.get("others") or "体卫艺劳科")
            st.caption("ℹ️ 说明：请于活动前一周周四下班前提交。")

        # Action Buttons
        col_b, col_d = st.columns([1, 2])
        with col_b:
            if st.button("⬅️ 修改信息"):
                 st.session_state.step = 1
                 st.rerun()

        with col_d:
            try:
                # 绿色下载按钮样式注入
                st.markdown("""
                <style>
                div.stButton > button:nth-last-child(1) {
                    background-color: #10b981 !important;
                    border-color: #10b981 !important;
                    color: white !important;
                }
                </style>
                """, unsafe_allow_html=True)
                
                final_data = {
                    "title": t, "content": c, "agenda": a, "time": tm, 
                    "duration": dr, "place": pl, "num": nm, "contact": ct, 
                    "projector": pj, "dist_leader": dl, "bur_leader": bl, "others": oth
                }
                tpl = DocxTemplate("申报单模板.docx")
                tpl.render(final_data)
                bio = io.BytesIO()
                tpl.save(bio)
                
                mmdd = datetime.now().strftime("%m%d")
                leader_name = bl.strip() if bl.strip() else (dl.strip() if dl.strip() else "领导")
                leader_name = leader_name.split('、')[0] if '、' in leader_name else leader_name
                filename = f"{mmdd}_{leader_name}_体卫艺劳科_{t}.docx"
                
                st.download_button(
                    label="📥 导出正式公文 (Word)",
                    data=bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error: {e}")

# ----------------- 模块二：龙华学校查号台 -----------------
else:
    st.caption("↖️ **导航：** 点击左上角 **>** 可切换功能")
    st.markdown("### 🔍 龙华学校查号台")
    st.caption("全区学校通讯录快速查询系统")
    
    if not st.session_state.contacts_authenticated:
        with st.container(border=True):
            st.info("🔒 访问权限验证")
            pwd = st.text_input("请输入访问密码", type="password")
            if st.button("解锁系统", type="primary", use_container_width=True):
                if pwd == CONTACT_PASSWORD:
                    st.session_state.contacts_authenticated = True
                    st.rerun()
                else:
                    st.error("❌ 密码错误")
        st.stop()

    @st.cache_data
    def load_contacts():
        try: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='utf-8-sig').fillna('无')
        except: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='gbk').fillna('无')

    df = load_contacts()
    
    with st.container(border=True):
        q = st.text_input("🔎 全文检索", placeholder="搜索学校、姓名...")
        
    if q:
        mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
        st.success(f"找到 {len(df[mask])} 条相关结果")
        st.dataframe(df[mask], use_container_width=True, hide_index=True)
    else:
        st.caption("👆 在上方输入关键词开始搜索")
