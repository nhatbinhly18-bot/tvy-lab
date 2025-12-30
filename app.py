   </div>
   {step_html}
   """, unsafe_allow_html=True)
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

# ----------------- 模块一：领导公务单生成器 (SaaS 版) -----------------
if mode == "📝 领导公务单自动生成器":
    
    # 顶部导航指引
    st.caption("↖️ **导航：** 点击左上角 **>** 可切换功能")
    
    # 动态渲染进度条 HTML
    s1_class = "step-active" if st.session_state.step == 1 else ""
    s2_class = "step-active" if st.session_state.step == 2 else ""
    
    step_html = f"""
    <div class="step-container">
        <div class="step-line"></div>
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
    
    # Hero Header Area
    st.markdown(f"""
    <div style="margin-bottom: 2rem;">
        <span class="saas-badge badge-primary">AI Powered V2.0</span>
        <h1>📋 体卫艺领导公务单自动生成器</h1>
        <p style="color: #6b7280; margin-top: -10px;">Technical Support Provided by Peipei</p>
</div>
{step_html}
""", unsafe_allow_html=True)
