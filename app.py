import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
import requests
import time
from openai import OpenAI
from datetime import datetime
import streamlit.components.v1 as components

# ==========================================
# 1. 核心配置与全局样式 (V1.2.1230.06)
# ==========================================
APP_VERSION = "V1.2.1230.06"
COZE_PAT = "pat_RX45BITCHQSbpPVvXaIIycBhfLrPrfdjJhGIokKJAVt8XJDTQmTmlCi8BvXOlWPZ"
BOT_ID = "7586932127112577033"
MY_API_KEY = "sk-dzsawqzsktjximglmkzyezbtyhqbysvenoxublemcgertlqp"
BASE_URL = "https://api.siliconflow.cn/v1"
CONTACT_PASSWORD = "lhjy"

st.set_page_config(page_title="体卫艺办公助手", page_icon="📋", layout="centered")

# --- 🎨 深度美化 CSS ---
st.markdown("""
<style>
    .stApp { background-image: linear-gradient(135deg, #f7f9fc 0%, #eceff4 100%); }
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        background-color: #ffffff; border: 1px solid #e1e4e8 !important;
        border-radius: 12px !important; box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        padding: 1.5rem !important;
    }
    h1 { color: #0d47a1; font-weight: 700 !important; }
    div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #1976d2 0%, #1565c0 100%) !important;
        border-radius: 20px !important; color: white !important;
    }
    /* 额外保留：绿色确认按钮样式 (用于下载) */
    div.stButton > button.green-button {
         background-color: #2e7d32 !important;
         color: white !important;
         border: none !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 核心函数：扣子 API (简报专家) ---
def call_briefing_expert(content):
    headers = {"Authorization": f"Bearer {COZE_PAT}", "Content-Type": "application/json"}
    payload = {
        "bot_id": BOT_ID, "user_id": "peipei_admin", "stream": False,
        "additional_messages": [{"role": "user", "content": content, "content_type": "text"}]
    }
    try:
        print(f"🚀 发送请求中...") 
        res = requests.post("https://api.coze.cn/v3/chat", headers=headers, json=payload, timeout=10).json()
        chat_id, conv_id = res['data']['id'], res['data']['conversation_id']

        for i in range(60):
            status_url = f"https://api.coze.cn/v3/chat/retrieve?chat_id={chat_id}&conversation_id={conv_id}"
            status_res = requests.get(status_url, headers=headers).json()
            curr_status = status_res.get('data', {}).get('status')
            
            print(f"⏳ 第 {i+1} 秒，状态: {curr_status}") 

            if curr_status == 'completed':
                print("✅ AI 创作完成！正在抓取...")
                msg_url = f"https://api.coze.cn/v3/chat/message/list?conversation_id={conv_id}&chat_id={chat_id}"
                msg_res = requests.get(msg_url, headers=headers).json()
                
                # --- 新增：把 AI 所有的消息类型都打印出来看看 ---
                messages = msg_res.get('data', [])
                print(f"🔍 调试信息：AI 一共返回了 {len(messages)} 条消息")
                for m in messages:
                    print(f"👉 消息类型: {m.get('type')}, 内容片段: {m.get('content')[:20]}...")

                for m in reversed(messages):
                    # 修改：不仅找 answer，也尝试抓取任何可能有内容的类型
                    if m.get('type') in ['answer', 'verbose']: 
                        print("🎉 成功抓取到内容！")
                        return m.get('content')
            
            time.sleep(1)
        return "⚠️ 超时：AI 写完了但代码没抓到内容"
    except Exception as e:
        return f"❌ 崩溃报错: {str(e)}"

# 初始化状态
if "contacts_authenticated" not in st.session_state:
    st.session_state.contacts_authenticated = False
if "parseddata_doc" not in st.session_state:
    st.session_state.parseddata_doc = None
if "step" not in st.session_state:
    st.session_state.step = 1
if "original_input" not in st.session_state:
    st.session_state.original_input = ""

# ==========================================
# 2. 侧边栏导航
# ==========================================
with st.sidebar:
    st.header("⚙️ 体卫艺办公助手")
    st.success(f"● AI 核心引擎已连接 ({APP_VERSION})")
    st.markdown("---")
    mode = st.radio("功能切换：", ["✨ 体卫艺简报助手", "📝 领导公务单自动生成器", "🔍 龙华学校查号台"])
    st.markdown("---")
    st.caption("维护者：孙沛 | 龙华区教育局体卫艺专用")
    
    st.write("")
    if st.button("🔒 退出并锁定系统"):
        st.session_state.contacts_authenticated = False
        st.session_state.parseddata_doc = None
        st.rerun()

# ==========================================
# 3. 核心功能逻辑
# ==========================================

if mode == "✨ 体卫艺简报助手":
    st.markdown("# ✨ 体卫艺简报助手")
    st.info("您好！我是擅长将杂乱信息转化为规范政务简讯的小助手，能为您打造高质量的体卫艺相关简报。👇 请直接发送：会议通知 + 参会名单 + 杂乱语音稿")
    
    u_content = st.text_area("✍️ 输入活动信息...", height=150, placeholder="例如：今天下午3点在教育局二楼会议室... （输入后点击下方生成按钮）")
    
    if st.button("🚀 生成润色简报", type="primary", use_container_width=True):
        if not u_content.strip():
            st.warning("⚠️ 请先输入内容")
        else:
            with st.spinner("🤖 笔杆子正在斟酌辞令..."):
                result = call_briefing_expert(u_content)
                if "⚠️" in result or "❌" in result:
                    st.error(result)
                else:
                    st.success("✅ 生成成功！")
                    st.info(result)
                    
                    # 复制/下载
                    st.download_button(
                        label="📋 复制简报内容",
                        data=result,
                        file_name=f"简报_{datetime.now().strftime('%m%d')}.txt",
                        mime="text/plain"
                    )

elif mode == "📝 领导公务单自动生成器":
    st.markdown("# 📋 公务单生成器")
    st.info("💡 请一次性说清：时间、地点、会议名称、人数、对接人、领导、参加部门及议程。")

    # --- Step 1: 输入与润色 ---
    if st.session_state.step == 1:
        with st.container(border=True):
            user_input = st.text_area("✍️ 描述活动信息", height=150, key="input_doc", placeholder="例：明天上午10点在会议室...")
        
        if st.button("✨ 立即智能填表", type="primary", use_container_width=True):
            if not user_input:
                st.warning("⚠️ 内容不能为空。")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                st.session_state.original_input = user_input
                
                current_date_str = datetime.now().strftime("%Y年%m月%d日")
                
                with st.spinner("🤖 正在解析要素并润色公文语言..."):
                    full_prompt = f"""
                    你现在是龙华教育局资深笔杆子。请解析以下描述：{user_input}
                    当前参考日期：{current_date_str}
                    
                    1. 人名纠错：杨灵芝、尹泽利、文良方、孙沛、刘冰、杨帆、陈海万等。
                    2. agenda 必须固定输出，严格包含这三项：["专题汇报", "座谈交流", "领导讲话"]。
                    3. content (理由背景)：转化为公文规范用语。
                    4. time (时间)："明天"等相对时间转为具体日期。
                    5. duration (时长)：统一计算为"X小时"。
                    
                    必须输出 JSON 格式，包含字段：title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。
                    """
                    try:
                        chat_completion = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct", 
                            messages=[{"role": "user", "content": full_prompt}], 
                            response_format={'type': 'json_object'},
                            timeout=30 
                        )
                        result = json.loads(chat_completion.choices[0].message.content)
                        st.session_state.parseddata_doc = result
                        st.session_state.step = 2
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 解析出错：{str(e)}")

    # --- Step 2: 确认与导出 ---
    elif st.session_state.step == 2 and st.session_state.parseddata_doc:
        d = st.session_state.parseddata_doc
        with st.container(border=True):
            st.subheader("2️⃣ 确认信息")
            t = st.text_input("📝 活动名称", d.get("title", ""))
            c = st.text_area("📄 申请理由/背景", d.get("content", ""), height=100)
            
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list): agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            if not agenda_val: agenda_val = "1. 专题汇报\n2. 座谈交流\n3. 领导讲话"
            a = st.text_area("📋 议程", agenda_val, height=120)
            
            c1, c2 = st.columns(2)
            with c1: tm = st.text_input("⏰ 时间", d.get("time", ""))
            with c2: dr = st.text_input("⏳ 时长", str(d.get("duration", "1小时")))
            
            c3, c4, c5 = st.columns([2, 1, 1])
            with c3: pl = st.text_input("📍 地点", d.get("place", ""))
            with c4: nm = st.text_input("👥 人数", d.get("num", ""))
            with c5: ct = st.text_input("👤 对接人", d.get("contact", "孙沛"))
            
            dist_l = st.text_input("👑 区领导", d.get("dist_leader", ""))
            bur_l = st.text_input("👑 局领导", d.get("bur_leader", ""))
            oth = st.text_input("🏛️ 参加单位", d.get("others") or "体卫艺劳科")

        col_b, col_d = st.columns([1, 2])
        if col_b.button("⬅️ 返回", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
            
        with col_d:
            # Word 生成逻辑
            try:
                final_data = {
                    "title": t, "content": c, "agenda": a, "time": tm, 
                    "duration": dr, "place": pl, "num": nm, "contact": ct, 
                    "dist_leader": dist_l, "bur_leader": bur_l, "others": oth,
                    "projector": "☑使用" # 简化处理
                }
                tpl = DocxTemplate("申报单模板.docx")
                tpl.render(final_data)
                bio = io.BytesIO()
                tpl.save(bio)
                
                filename = f"{datetime.now().strftime('%m%d')}_{t}.docx"
                
                # 特别样式注入，确保下载按钮变色
                st.markdown("""<style>div.stButton > button:nth-last-child(1) {background-color: #2e7d32 !important; color: white !important;}</style>""", unsafe_allow_html=True)
                
                st.download_button(
                    label="💾 确认并导出 Word",
                    data=bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"生成失败: {e}")

else: # 查号台
    st.markdown("### 🔍 龙华学校查号台")
    if not st.session_state.contacts_authenticated:
        pwd = st.text_input("🔒 请输入密码", type="password")
        if st.button("验证登录", type="primary"):
            if pwd == CONTACT_PASSWORD:
                st.session_state.contacts_authenticated = True
                st.rerun()
            else:
                st.error("❌ 密码错误")
    else:
        @st.cache_data
        def load_contacts():
            try: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='utf-8-sig').fillna('无')
            except: return pd.read_csv('龙华中小学校通讯录（含幼儿园）.csv', encoding='gbk').fillna('无')
            
        df = load_contacts()
        q = st.text_input("🔎 搜学校 / 搜人名")
        if q:
            mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
            st.dataframe(df[mask], use_container_width=True, hide_index=True)
        else:
            st.caption("👆 输入关键词开始搜索")
