import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime

# 1. 网页基础配置
st.set_page_config(page_title="体卫艺办公助手", page_icon="🚀", layout="centered")

# --- 🎨 深度美化 / CSS 设计 ---
st.markdown("""
<style>
    /* 全局字体 */
    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', 'PingFang SC', 'Microsoft YaHei', sans-serif;
    }
    
    /* 页面背景 - 商务风云雾白 */
    .stApp {
        background-color: #f7f9fc;
        background-image: linear-gradient(135deg, #f7f9fc 0%, #eceff4 100%);
    }

    /* 侧边栏样式 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e1e4e8;
        box-shadow: 2px 0 10px rgba(0,0,0,0.01);
    }
    
    /* 隐藏顶部红线 */
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* ---------------- 卡片式容器设计 ---------------- */
    /* 所有的 st.container(border=True) 都会应用这个样式 */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        background-color: #ffffff;
        border: 1px solid #e1e4e8 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        padding: 1.5rem !important;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    
    /* ---------------- 标题与文字 ---------------- */
    h1 {
        color: #0d47a1; /* 商务深蓝 */
        font-weight: 700 !important;
        letter-spacing: -0.5px;
        margin-bottom: 0.5rem !important;
    }
    
    h2, h3 {
        color: #1565c0;
        font-weight: 600 !important;
    }
    
    .stMarkdown p {
        color: #424242;
        line-height: 1.6;
    }

    /* ---------------- 交互组件美化 ---------------- */
    
    /* 输入框优化 */
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 8px !important;
        border: 1px solid #cfd8dc !important;
        background-color: #fcfcfc !important;
        transition: all 0.3s ease;
    }
    
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #1976d2 !important;
        background-color: #ffffff !important;
        box-shadow: 0 0 0 3px rgba(25, 118, 210, 0.1) !important;
    }
    
    /* 按钮美化 - 圆角 + 阴影 */
    div.stButton > button {
        border-radius: 20px !important; /* 圆角胶囊样式 */
        font-weight: 600 !important;
        padding: 0.5rem 1.5rem !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }
    
    /* 主要按钮 (Primary) - 商务蓝 */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #1976d2 0%, #1565c0 100%) !important;
        color: white !important;
        box-shadow: 0 4px 6px rgba(21, 101, 192, 0.2) !important;
    }
    
    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(21, 101, 192, 0.3) !important;
    }
    
    /* 下载按钮 (Secondary) - 保持醒目但和谐 */
    div.stButton > button[kind="secondary"] {
        background-color: #ffffff !important;
        color: #1565c0 !important;
        border: 1px solid #1565c0 !important;
    }

    /* ---------------- 表格与手机端优化 ---------------- */
    /* 查号台表格优化 */
    [data-testid="stDataFrame"] {
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #e0e0e0;
    }
    
    /* 手机端间距调整 */
    @media (max-width: 768px) {
        .block-container {
            padding-top: 2rem !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        h1 { font-size: 1.5rem !important; }
        div[data-testid="stVerticalBlockBorderWrapper"] > div {
            padding: 1rem !important; /* 手机端卡片内边距减小 */
        }
    }
</style>
""", unsafe_allow_html=True)

# --- 🔒 通讯录专属密码 ---
CONTACT_PASSWORD = "lhjy" 

# 2. 核心配置
MY_API_KEY = "sk-dzsawqzsktjximglmkzyezbtyhqbysvenoxublemcgertlqp"
BASE_URL = "https://api.siliconflow.cn/v1"

# 初始化状态
if "contacts_authenticated" not in st.session_state:
    st.session_state.contacts_authenticated = False
if "parseddata_doc" not in st.session_state:
    st.session_state.parseddata_doc = None
# 新增：两步流程的状态管理
if "step" not in st.session_state:
    st.session_state.step = 1  # 1=输入, 2=确认润色, 3=确认字段
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
    st.info("""
    **💡 助手功能说明：**
    
    1. **公务单生成**：
       语音口语 → 规范公文Word
       
    2. **学校查号台**：
       全区通讯录一键查询
    """)
    st.caption("维护者：孙沛 | 龙华区教育局体卫艺专用")
    
    st.write("") # Spacer
    if st.button("🔒 退出并锁定系统"):
        st.session_state.contacts_authenticated = False
        st.session_state.parseddata_doc = None
        st.rerun()

# ----------------- 模块一：领导公务单生成器 -----------------
if mode == "📝 领导公务单自动生成器":
    
    # 使用容器包裹标题区域，打造卡片感
    st.markdown("# 🚀 领导公务单自动生成器")
    st.caption("Intelligent Official Document Generator")
    
    # 蓝色提示框 - 提示语
    st.info("""
    **💡 智能提示：** 请一次性说清：时间、地点、会议名称、人数、对接人、领导、参加部门及议程。
    
    **🗣️ 参考范例：** “明天上午10点在二楼多功能厅有个生涯教育座谈会，大概20人，孙沛对接，1小时，邀请灵芝主任参加。”
    """)

    # --- 第一步：输入与润色 ---
    if st.session_state.step == 1:
        # 输入区卡片
        with st.container(border=True):
            st.subheader("1️⃣ 描述活动信息")
            st.caption("支持直接粘贴语音转文字的内容，AI 将自动提取要素。")
            
            user_input = st.text_area(
                "请在此输入...", 
                height=150, 
                placeholder="请点击此处粘贴或输入内容...", 
                key="input_doc", 
                label_visibility="collapsed"
            )
        
        st.write("") # 间距
        if st.button("✨ 立即智能填表并生成 Word", type="primary", use_container_width=True):
            if not user_input:
                st.warning("⚠️ 内容不能为空，请输入活动描述。")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                st.session_state.original_input = user_input
                
                # 获取当前日期用于计算相对时间
                current_date_str = datetime.now().strftime("%Y年%m月%d日")
                weekday = datetime.now().strftime("%w")
                
                with st.spinner("🤖 正在解析要素并润色公文语言..."):
                    
                    # 标准人名库（用于纠正语音转文字的谐音错误）
                    name_corrections = {
                        "林芝": "杨灵芝", "杨林芝": "杨灵芝",
                        "陈海湾": "陈海万", "陈海完": "陈海万",
                        "尹泽力": "尹泽利", "尹则利": "尹泽利",
                        "文量方": "文良方", "温良方": "文良方",
                        "刘兵": "刘冰",
                        "梁永育": "梁永誉",
                        "方梦仪": "方梦懿"
                    }
                    
                    full_prompt = f"""
                    你现在是龙华教育局资深笔杆子。请根据以下用户的大白话描述，解析出公文要素，并对【理由背景】和【议程】部分进行专业润色。
                    
                    【当前日期参考】：今天是 {current_date_str} (星期{weekday})。
                    【用户输入】：{user_input}
                    
                    【标准人名库】（请优先匹配）：
                    杨灵芝、尹泽利、文良方、孙沛、刘冰、杨帆、陈海万、路旭阳、王轩、王燕、李桂情、甘月琴、方梦懿、吴正光、李长生、梁永誉、刘喜菊
                    
                    【解析与润色要求】：
                    1. **人名纠错**：如果用户输入的人名与标准人名库相似（如"林芝"应为"杨灵芝"，"陈海湾"应为"陈海万"），请自动纠正为标准名字。
                    2. **content (理由背景)**：将用户的背景描述转化为"为落实...要求，推进...发展"等公文规范用语。只有动宾结构和语序调整，严禁杜撰。
                    3. **agenda (详细议程)**：**固定输出以下三项，顺序不可变**：["专题汇报", "座谈交流", "领导讲话"]。**严禁添加、删除或修改这三项**，无论用户输入什么。
                    4. **time (时间)**：必须将"明天"、"后天"、"周三"等相对时间**计算为具体的年月日**（格式：YYYY年MM月DD日 HH:MM）。禁止直接写"明天"或"下周"。
                    5. **duration (时长)**：统一计算为"X小时"或"X.5小时"（如1.5小时），**不要用分钟**。
                    6. **contact (公务对接人)**：提取人名（优先从标准人名库匹配），若无则默认为"孙沛"。
                    7. **dist_leader (区领导)** / **bur_leader (局领导)**：准确提取拟请出席的领导职务/姓名（如"灵芝主任"应识别为"杨灵芝"）。**严禁添加"教育发展中心"等部门前缀**，直接写姓名加职务即可。
                    8. **others (参加单位)**：提取建议参加的部门或单位。
                    9. **其他字段**：title(活动名称), place(地点), num(人数), projector(投影仪: ☑是/☐否)。
                    
                    必须以 JSON 格式严格输出，包含以下字段：
                    title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。
                    """
                    
                    try:
                        chat_completion = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct", 
                            messages=[{"role": "user", "content": full_prompt}], 
                            response_format={'type': 'json_object'},
                            timeout=30 
                        )
                        result = json.loads(chat_completion.choices[0].message.content)
                        
                        # 字段健壮性处理
                        required_fields = ["title", "content", "agenda", "time", "place", "num", "contact"]
                        for field in required_fields:
                            if field not in result:
                                result[field] = ""
                        
                        st.session_state.parseddata_doc = result
                        st.session_state.step = 2  # 跳到确认表单
                        st.rerun()
                        
                    except json.JSONDecodeError:
                         st.error("❌ AI 解析返回格式有误，请尝试补充细节后重试。")
                    except TimeoutError:
                        st.error("⏱️ 请求超时，网络可能较慢。请稍后重试或简化输入内容。")
                    except Exception as e:
                        st.error(f"❌ 解析出错：{str(e)}")

    # --- 第二步：确认所有字段 ---
    elif st.session_state.step == 2 and st.session_state.parseddata_doc:
        d = st.session_state.parseddata_doc
        
        # 预览区卡片
        with st.container(border=True):
            st.subheader("2️⃣ 核心要素预览与微调")
            st.markdown("**📌 申报部门：体卫艺劳科**") 
            
            t = st.text_input("📝 政务活动名称", d.get("title", ""))
            c = st.text_area("📄 政务活动申请理由、背景", d.get("content", ""), height=100)
            
            # 处理 agenda
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list):
                agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            if not agenda_val:
                agenda_val = "1. 专题汇报\n2. 座谈交流\n3. 领导讲话"
            a = st.text_area("📋 议程", agenda_val, height=120)
            
            st.divider() # 分割线
            
            col1, col2 = st.columns(2)
            with col1:
                tm = st.text_input("⏰ 时间", d.get("time", ""))
                
                # 确保时长有单位
                duration_val = d.get("duration", "1小时")
                duration_val = str(duration_val) if duration_val else "1小时"
                if "小时" not in duration_val:
                    duration_val = f"{duration_val}小时"
                dr = st.text_input("⏳ 会议时长", duration_val)
                
            with col2:
                st.text_input("🚫 时间调整", "不可调整", disabled=True) 
                ct = st.text_input("👤 公务对接人", d.get("contact", "孙沛"))

            col3, col4, col5 = st.columns([2, 1, 1])
            with col3:
                pl = st.text_input("📍 地点", d.get("place", ""))
            with col4:
                nm = st.text_input("👥 人数", d.get("num", ""))
            with col5:
                pj = st.selectbox("📽️ 投影仪", ["☑使用", "☐不使用"], index=0 if "是" in str(d.get("projector")) else 1)
            
            st.divider()
            st.markdown("**👑 领导出席**")
            dist_l = st.text_input("1. 拟请出席的区领导", d.get("dist_leader", ""))
            bur_l = st.text_input("2. 拟请办公室协调出席的局领导", d.get("bur_leader", ""))
            
            st.divider()
            oth = st.text_input("🏛️ 建议参加单位(部门)", d.get("others") or "体卫艺劳科")
            
            st.caption("ℹ️ 说明：此表请于政务活动前一周星期四下班前交办公室登记汇总。")

        col_final_back, col_final_down = st.columns([1, 2])
        with col_final_back:
            if st.button("⬅️ 返回修改"):
                 st.session_state.step = 1
                 st.rerun()

        with col_final_down:
            try:
                final_data = {
                    "title": t, "content": c, "agenda": a, "time": tm, 
                    "duration": dr, "place": pl, "num": nm, "contact": ct, 
                    "projector": pj, "dist_leader": dist_l, "bur_leader": bur_l, "others": oth
                }
                tpl = DocxTemplate("申报单模板.docx")
                tpl.render(final_data)
                bio = io.BytesIO()
                tpl.save(bio)
                
                mmdd = datetime.now().strftime("%m%d")
                leader_name = bur_l.strip() if bur_l.strip() else (dist_l.strip() if dist_l.strip() else "领导")
                leader_name = leader_name.split('、')[0] if '、' in leader_name else leader_name
                filename = f"{mmdd}_{leader_name}_体卫艺劳科_{t}.docx"
                
                # 注入自定义样式，让下载按钮在不改变原生type的情况下变色
                st.markdown("""
                <style>
                    /* 定位最后一个按钮（通常是下载按钮，因为返回按钮在它前面） */
                    div.stButton > button:nth-last-child(1) {
                         background-color: #2e7d32 !important; /* 绿色 */
                         color: white !important;
                         border: none !important;
                    }
                </style>
                """, unsafe_allow_html=True)
                
                # 核心：直接使用原生按钮，触发微信的系统拦截机制 - 绝对不动
                st.download_button(
                    label="💾 确认无误，导出 Word",
                    data=bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"生成失败：{e}")

# ----------------- 模块二：龙华学校查号台 -----------------
else:
    st.markdown("### 🔍 龙华学校查号台")
    st.caption("全区学校通讯录快速查询系统")
    
    if not st.session_state.contacts_authenticated:
        # 登录卡片
        with st.container(border=True):
            st.info("🔒 内部数据访问受限")
            pwd = st.text_input("请输入授权密码", type="password", help="请向管理员获取密码")
            if st.button("验证登录", type="primary", use_container_width=True):
                if pwd == CONTACT_PASSWORD:
                    st.session_state.contacts_authenticated = True
                    st.rerun()
                else:
                    st.error("❌ 密码错误，请重试。")
        st.stop()

    @st.cache_data
    def load_contacts():
        try:
            return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='utf-8-sig').fillna('无')
        except:
            return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='gbk').fillna('无')

    df = load_contacts()
    
    # 搜索框卡片
    with st.container(border=True):
        q = st.text_input("🔎 快速搜索", placeholder="输入学校名或人名关键词（如：龙华中学 或 张三）...")
        
    if q:
        mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
        st.write(f"📊 搜索结果：找到 {len(df[mask])} 条记录")
        st.dataframe(df[mask], use_container_width=True, hide_index=True)
    else:
        st.caption("👆 在上方输入关键词开始搜索，支持模糊匹配。")
