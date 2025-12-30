import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime
import streamlit.components.v1 as components
# 1. 网页基础配置
APP_VERSION = "V1.2.1230.01"
st.set_page_config(page_title="体卫艺办公助手", page_icon="🚀", layout="centered")

# --- Mobile Optimization / Custom CSS ---
st.markdown("""
<style>
    /* 隐藏顶部菜单和页脚，但保留移动端侧边栏按钮 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 桌面端隐藏header，移动端保留以便访问侧边栏 */
    @media (min-width: 769px) {
        header {visibility: hidden;}
    }

    /* 调整移动端内边距 */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 0.75rem;
        padding-right: 0.75rem;
    }
    
    /* 侧边栏优化 */
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
    }
    
    /* 移动端字体优化 */
    @media (max-width: 768px) {
        /* 基础字体减小 */
        html, body, [class*="css"] {
            font-size: 14px !important;
        }
        
        /* 标题字体调整 */
        h1 {
            font-size: 1.5rem !important;
            margin-bottom: 0.5rem !important;
        }
        
        h2 {
            font-size: 1.25rem !important;
            margin-bottom: 0.5rem !important;
        }
        
        h3 {
            font-size: 1.1rem !important;
            margin-bottom: 0.5rem !important;
        }
        
        /* 输入框标签字体 */
        [data-testid="stWidgetLabel"] {
            font-size: 0.85rem !important;
        }
        
        /* 按钮字体 */
        button {
            font-size: 0.9rem !important;
            padding: 0.4rem 0.8rem !important;
        }
        
        /* 文本区域和输入框 */
        textarea, input {
            font-size: 0.9rem !important;
        }
        
        /* 侧边栏字体 */
        [data-testid="stSidebar"] {
            font-size: 0.85rem !important;
        }
        
        /* 减小元素间距 */
        .element-container {
            margin-bottom: 0.5rem !important;
        }
        
        /* Expander 标题 */
        [data-testid="stExpander"] summary {
            font-size: 0.9rem !important;
        }
    }

    /* 绿色按钮 (涵盖主要按钮和下载按钮) */
    div.stButton > button[kind="primary"], div.stButton > button:last-child {
        background-color: #28a745 !important;
        border-color: #28a745 !important;
        color: white !important;
    }
    div.stButton > button[kind="primary"]:hover, div.stButton > button:last-child:hover {
        background-color: #218838 !important;
        border-color: #1e7e34 !important;
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
   # --- 从这里开始替换 (对应原 129 行) ---
    st.success(f"● AI 核心引擎已连接 ({APP_VERSION})") 
    
    st.info(f"""
    **💡 使用小提示：** 本助手集成三大核心功能：
    
    1. **公务单生成**：智能解析文字生成 Word。
    2. **学校查号台**：全区学校通讯录快速查询。
    3. **智能简报生成**：AI 驱动的文稿创作中心。
    
    您可以通过下方的 **“功能切换”** 选项随时跳转。
    """)
    
    st.caption(f"当前版本：{APP_VERSION}")
    st.caption("维护者：孙沛 | 龙华区教育局体卫艺专用")
    st.divider()
    # --- 替换结束 (衔接到原 142 行) ---
    
    mode = st.radio("功能切换：", ["📝 领导公务单自动生成器", "🔍 龙华学校查号台", "✨ 智能简报生成"])
    
    st.write("") # Spacer
    if st.button("🔒 退出并锁定"):
        st.session_state.contacts_authenticated = False
        st.session_state.parseddata_doc = None
        st.rerun()

# ----------------- 模块一：领导公务单生成器 -----------------
if mode == "📝 领导公务单自动生成器":
    # Custom CSS for compact layout
    # st.markdown("""
    # <style>
    #     /* 完全去除所有间距 */
    #     .main .block-container {
    #         padding-top: 0.5rem;
    #         padding-bottom: 0.5rem;
    #     }
    #     
    #     /* 标题完全无间距 */
    #     h1, h2, h3 {
    #         margin-top: 0 !important;
    #         margin-bottom: 0 !important;
    #         padding-top: 0 !important;
    #         padding-bottom: 0 !important;
    #     }
    #     
    #     /* 段落完全无间距 */
    #     p {
    #         margin-top: 0 !important;
    #         margin-bottom: 0 !important;
    #         padding-top: 0 !important;
    #         padding-bottom: 0 !important;
    #     }
    #     
    #     /* info/warning 框最小间距 */
    #     .stAlert {
    #         margin-top: 0.2rem !important;
    #         margin-bottom: 0.2rem !important;
    #         padding: 0.5rem 1rem !important;
    #     }
    #     
    #     /* 所有元素间距为0 */
    #     .element-container {
    #         margin-top: 0 !important;
    #         margin-bottom: 0 !important;
    #         padding-top: 0 !important;
    #         padding-bottom: 0 !important;  
    #     }
    #     
    #     /* 绿色按钮样式 */
    #     div.stButton > button:first-child[kind="primary"] {
    #         background-color: #28a745;
    #         border-color: #28a745;
    #         color: white;
    #     }
    #     div.stButton > button:first-child[kind="primary"]:hover {
    #         background-color: #218838;
    #         border-color: #1e7e34;
    #     }
    # </style>
    # """, unsafe_allow_html=True)
    # 醒目的功能切换提示（方便年长用户）
    st.warning("👆 点击左上角 **>>** 可切换到「查号台」")
    st.markdown("# 🚀 领导公务单自动生成器")
    st.markdown("<div style='font-size: 18px; margin: 0.3rem 0; line-height: 1.4;'>欢迎使用！本工具旨在帮您一键完成体卫艺政务活动申报。</div>", unsafe_allow_html=True)

    # --- 蓝色提示框（固定显示）---
    st.info("""
    **💡 请一次性说清：** 时间、地点、会议名称、人数、对接人、领导、参加部门、背景及议程。
    
    **参考范例：** 明天上午10点在二楼多功能厅有个生涯教育座谈会，大概20人，孙沛对接，时长1小时，邀请灵芝主任参加
    """)
    # Custom CSS for Green Button
    st.markdown("""
    <style>
        div.stButton > button:first-child[kind="primary"] {
            background-color: #28a745;
            border-color: #28a745;
            color: white;
        }
        div.stButton > button:first-child[kind="primary"]:hover {
            background-color: #218838;
            border-color: #1e7e34;
        }
    </style>
    """, unsafe_allow_html=True)

    # --- 第一步：输入与润色 ---
    if st.session_state.step == 1:
        user_input = st.text_area("✍️ 请输入活动描述（支持语音转文字复制粘贴）：", height=150, placeholder="请在此粘贴或输入内容...", key="input_doc")
        
        if st.button("✨ 立即智能填表并生成 Word", type="primary"):
            if not user_input:
                st.warning("内容不能为空。")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                st.session_state.original_input = user_input
                
                # 获取当前日期用于计算相对时间
                current_date_str = datetime.now().strftime("%Y年%m月%d日")
                weekday = datetime.now().strftime("%w")
                
                with st.spinner("正在解析要素并润色公文语言...（通常需要 5-15 秒，请耐心等待）"):
                    
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
                            model="Qwen/Qwen2.5-7B-Instruct",  # 更快的模型，速度提升 3-5 倍 
                            messages=[{"role": "user", "content": full_prompt}], 
                            response_format={'type': 'json_object'},
                            timeout=30  # 30秒超时
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
                        st.error("⏱️ 请求超时（超过30秒），网络可能较慢。请稍后重试或简化输入内容。")
                    except Exception as e:
                        st.error(f"❌ 解析出错：{str(e)}")
                        st.info("💡 **建议**：检查网络连接 / 简化输入内容 / 稍后重试")

    # --- 第二步：确认所有字段 ---
    elif st.session_state.step == 2 and st.session_state.parseddata_doc:
        d = st.session_state.parseddata_doc
        with st.container(border=True):
            st.markdown("### 🧐 核心要素预览与微调")
            st.markdown("**申报部门：体卫艺劳科**") # 假设固定
            
            t = st.text_input("📝 政务活动名称", d.get("title", ""))
            c = st.text_area("📄 政务活动申请理由、背景", d.get("content", ""), height=80)
            
            # 处理 agenda 可能是 list 的情况
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list):
                agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            # 如果没有议程，使用默认值
            if not agenda_val:
                agenda_val = "1. 专题汇报\n2. 座谈交流\n3. 领导讲话"
            a = st.text_area("📋 议程", agenda_val, height=120)
            
            col1, col2 = st.columns(2)
            with col1:
                tm = st.text_input("⏰ 时间", d.get("time", ""))
                
                # 确保时长有单位
                duration_val = d.get("duration", "1小时")
                # 转换为字符串以避免类型错误
                duration_val = str(duration_val) if duration_val else "1小时"
                if "小时" not in duration_val:
                    duration_val = f"{duration_val}小时"
                dr = st.text_input("⏳ 会议时长", duration_val)
                
            with col2:
                # Mockup checkbox visualization for "可否调整" (visual only for now)
                st.caption("时间可否调整：☑否") 
                ct = st.text_input("👤 公务对接人", d.get("contact", "孙沛"))

            col3, col4, col5 = st.columns([2, 1, 1])
            with col3:
                pl = st.text_input("📍 地点", d.get("place", ""))
            with col4:
                nm = st.text_input("👥 人数", d.get("num", ""))
            with col5:
                pj = st.selectbox("📽️ 投影仪", ["☑使用", "☐不使用"], index=0 if "是" in str(d.get("projector")) else 1)
            
            st.markdown("---")
            dist_l = st.text_input("1. 拟请出席的区领导", d.get("dist_leader", ""))
            bur_l = st.text_input("2. 拟请办公室协调出席的局领导", d.get("bur_leader", ""))
            
            st.markdown("---")
            oth = st.text_input("建议参加单位(部门)", d.get("others") or "体卫艺劳科")
            
            st.caption("说明：此表请于政务活动前一周星期四下班前交办公室登记汇总。")

        col_final_back, col_final_down = st.columns([1, 2])
        with col_final_back:
            if st.button("⬅️ 返回上一步"):
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
                
                # 核心：直接使用原生按钮，触发微信的系统拦截机制
                st.download_button(
                    label="💾 确认无误，导出 Word",
                    data=bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"生成失败：{e}")


# ----------------- 模块三：智能简报生成 -----------------
elif mode == "✨ 智能简报生成":
    st.markdown("# ✨ 体卫艺 AI 简报创作中心")
    st.info("💡 正在连接 DeepSeek V3 智慧大脑... 审核通过后，此处将直接显示对话窗口。")
    components.iframe("https://www.baidu.com", height=750, scrolling=True)

# ----------------- 模块二：龙华学校查号台 -----------------
else:
    st.markdown("### 🔍 龙华学校查号台")
    if not st.session_state.contacts_authenticated:
        st.info("🔒 为了数据安全，访问通讯录需要授权。")
        pwd = st.text_input("请输入授权密码", type="password", help="请向管理员获取密码")
        if st.button("验证登录", type="primary"):
            if pwd == CONTACT_PASSWORD:
                st.session_state.contacts_authenticated = True
                st.rerun()
            else:
                st.error("密码错误，请重试。")
        st.stop()

    @st.cache_data
    def load_contacts():
        try:
            return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='utf-8-sig').fillna('无')
        except:
            return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='gbk').fillna('无')

    df = load_contacts()
    q = st.text_input("请输入学校名或人名关键词：", placeholder="例如：龙华中学 或 张三")
    if q:
        mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
        st.dataframe(df[mask], use_container_width=True, hide_index=True)
    else:
        st.caption("👆 在上方输入关键词开始搜索")
