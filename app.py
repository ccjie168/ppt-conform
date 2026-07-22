import streamlit as st
import tempfile
import os
import json
import subprocess
from pathlib import Path
from pptx import Presentation

from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.precheck.analyzer import PreCheckAnalyzer
from core.qa.reporter import QAReporter
from core.models import PreCheckResult


# 内置模板路径
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", "2026 se template eng.pptx")

# 4 种背景风格
BACKGROUND_STYLES = [
    {"key": "white", "name": "白色简约", "color": "#FFFFFF", "text_color": "#0A2F24", "icon": "⬜"},
    {"key": "light_green", "name": "浅绿色清新", "color": "#E7FFD9", "text_color": "#0A2F24", "icon": "🟩"},
    {"key": "dark_green", "name": "深绿色商务", "color": "#0A2F24", "text_color": "#FFFFFF", "icon": "🟢"},
    {"key": "gradient", "name": "渐变科技", "color": "linear-gradient(135deg, #0A2F24, #3DCD58)", "text_color": "#FFFFFF", "icon": "🎨"},
]


def _get_git_commit() -> str:
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=os.path.dirname(__file__),
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return "unknown"


def _get_git_commit_date() -> str:
    try:
        result = subprocess.run(
            ["git", "log", "-1", "--format=%cd", "--date=format:%Y-%m-%d"],
            cwd=os.path.dirname(__file__),
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return ""


APP_VERSION = _get_git_commit()
APP_VERSION_DATE = _get_git_commit_date()


st.set_page_config(
    page_title="PPT 标准模板转换",
    page_icon="⚡",
    layout="wide"
)


def _load_css():
    css_path = Path(__file__).parent / "static" / "schneider_style.css"
    if css_path.exists():
        with open(css_path, "r", encoding="utf-8") as f:
            css = f.read()
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


_load_css()


PERSIST_DIR = os.path.join(os.path.dirname(__file__), ".persist")


def _ensure_persist_dir():
    os.makedirs(PERSIST_DIR, exist_ok=True)


def _load_history():
    """加载转换历史记录"""
    try:
        history_file = os.path.join(PERSIST_DIR, "history.json")
        if os.path.exists(history_file):
            with open(history_file, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []


def _save_history(history):
    """保存转换历史记录"""
    _ensure_persist_dir()
    try:
        history_file = os.path.join(PERSIST_DIR, "history.json")
        with open(history_file, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _add_conversion_record(input_name, output_name, background_style, slide_count):
    """添加一条转换记录"""
    from datetime import datetime
    history = _load_history()
    style_name = next((s["name"] for s in BACKGROUND_STYLES if s["key"] == background_style), "未知风格")
    record = {
        "id": datetime.now().strftime("%Y%m%d%H%M%S"),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "input_name": input_name,
        "output_name": output_name,
        "background_style": background_style,
        "style_name": style_name,
        "slide_count": slide_count,
    }
    history.insert(0, record)
    history = history[:50]
    _save_history(history)
    return record


def _run_precheck(pptx_file) -> PreCheckResult | None:
    """执行 Pre-check 分析"""
    if pptx_file is None:
        return None
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        tmp.write(pptx_file.getbuffer())
        tmp_path = tmp.name
    try:
        analyzer = PreCheckAnalyzer()
        return analyzer.analyze(tmp_path)
    finally:
        os.unlink(tmp_path)


def _render_precheck(result: PreCheckResult):
    """渲染 Pre-check 报告"""
    st.markdown('<div class="config-card" style="margin-bottom: 20px;">', unsafe_allow_html=True)
    st.markdown('<div class="config-card-title">PPT 质量分析报告</div>', unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("页面数量", result.slide_count)
    col2.metric("母版数量", result.master_count)
    col3.metric("字体种类", len(result.fonts_used))
    col4.metric("越界对象", result.overflow_objects_count)

    st.markdown("---")

    checks = [
        ("4:3 比例", result.is_4_3_ratio, "warning"),
        ("含动画", result.has_animation, "info"),
        ("含嵌入图表", result.has_embedded_chart, "warning"),
        ("含 SmartArt", result.has_smartart, "warning"),
        ("含媒体", result.has_media, "warning"),
        ("旧 SE 模板", result.has_old_se_template, "info"),
    ]

    icon_cols = st.columns(6)
    for i, (label, value, level) in enumerate(checks):
        with icon_cols[i]:
            icon = "✅" if not value else ("⚠️" if level == "warning" else "ℹ️")
            st.markdown(
                f"<div style='text-align: center; padding: 8px;'>"
                f"<div style='font-size: 20px;'>{icon}</div>"
                f"<div style='font-size: 11px; color: #718096; margin-top: 4px;'>{label}</div>"
                f"</div>",
                unsafe_allow_html=True
            )

    if result.fonts_used:
        st.markdown("---")
        st.markdown(f"**使用字体：** {', '.join(result.fonts_used)}")

    if result.issues:
        st.markdown("---")
        st.markdown("**发现的问题：**")
        for issue in result.issues:
            icon = "❌" if issue.level == "error" else ("⚠️" if issue.level == "warning" else "ℹ️")
            st.markdown(f"- {icon} {issue.message}")

    st.markdown('</div>', unsafe_allow_html=True)


# ============================================
#   侧边栏
# ============================================
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="sidebar-brand-icon">✓</div>
        <div>
            <div class="sidebar-brand-text">Schneider</div>
            <div class="sidebar-brand-sub">PPT 转换工具</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="sidebar-nav">', unsafe_allow_html=True)

    nav_items = [
        {"icon": "🔄", "label": "模板转换", "key": "convert"},
        {"icon": "📋", "label": "历史记录", "key": "history"},
        {"icon": "📖", "label": "使用说明", "key": "help"},
    ]

    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "convert"

    for item in nav_items:
        is_active = st.session_state.get("current_page") == item["key"]
        if st.button(
            f"{item['icon']}  {item['label']}",
            key=f"nav_{item['key']}",
            use_container_width=True,
            type="primary" if is_active else "secondary",
        ):
            st.session_state["current_page"] = item["key"]
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="sidebar-user">
        <div class="sidebar-user-avatar" style="background: linear-gradient(135deg, #3DCD58 0%, #2EAE4A 100%);">JC</div>
        <div class="sidebar-user-info">
            <div class="sidebar-user-name">JC</div>
            <div class="sidebar-user-role">设计部</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


current_page = st.session_state.get("current_page", "convert")


# ============================================
#   页面 1: 模板转换
# ============================================
if current_page == "convert":
    st.markdown('<h1 class="page-title">PPT 标准模板转换</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">上传 PPT 文件，智能转换为标准企业模板格式</p>', unsafe_allow_html=True)

    # ---------- 上传区域 ----------
    uploaded_file = st.file_uploader(
        "点击或拖拽上传 PPT 文件",
        type=["pptx"],
        key="input_ppt",
        label_visibility="visible",
    )

    if uploaded_file is not None:
        st.markdown(f"""
        <div class="upload-card-wrapper">
            <div class="upload-card uploaded">
                <div class="upload-card-icon">✅</div>
                <div class="upload-card-title">{uploaded_file.name}</div>
                <div class="upload-card-desc">{uploaded_file.size / 1024:.1f} KB · 已上传</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ---------- Pre-check 分析 ----------
    if uploaded_file is not None:
        precheck_result = _run_precheck(uploaded_file)
        if precheck_result:
            _render_precheck(precheck_result)

    # ---------- 转换配置卡片 ----------
    st.markdown('<div class="config-card animate-fade-in">', unsafe_allow_html=True)
    st.markdown('<div class="config-card-title">转换配置</div>', unsafe_allow_html=True)

    # 背景风格选择（4 种固定选项）
    st.markdown("**选择背景风格**")
    style_cols = st.columns(4)
    selected_style = st.session_state.get("background_style", "dark_green")

    for i, style in enumerate(BACKGROUND_STYLES):
        with style_cols[i]:
            is_selected = selected_style == style["key"]
            bg = style["color"]
            txt = style["text_color"]
            border = "2px solid #3DCD58" if is_selected else "2px solid #E2E8F0"
            st.markdown(f"""
            <div style="
                background: {bg};
                border: {border};
                border-radius: 10px;
                padding: 16px 8px;
                text-align: center;
                cursor: pointer;
                color: {txt};
                margin-bottom: 8px;
            ">
                <div style="font-size: 24px;">{style['icon']}</div>
                <div style="font-size: 13px; font-weight: 600; margin-top: 4px;">{style['name']}</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("选择", key=f"style_{style['key']}", use_container_width=True,
                         type="primary" if is_selected else "secondary"):
                st.session_state["background_style"] = style["key"]
                st.rerun()

    st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)

    # 补充说明
    notes = st.text_area(
        "补充说明",
        placeholder="可选：输入补充转换要求",
        height=80,
        key="notes_textarea",
    )

    st.markdown('</div>', unsafe_allow_html=True)

    # ---------- 底部操作栏 ----------
    st.markdown('<div class="action-bar">', unsafe_allow_html=True)

    reset_col, convert_col = st.columns([1, 1])
    with reset_col:
        if st.button("重置", use_container_width=True, key="reset_btn"):
            for key in list(st.session_state.keys()):
                if key not in ["current_page", "background_style"]:
                    del st.session_state[key]
            st.rerun()
    with convert_col:
        can_convert = uploaded_file is not None
        convert_button = st.button(
            "开始转换",
            disabled=not can_convert,
            type="primary",
            use_container_width=True,
            key="convert_btn",
        )

    st.markdown('</div>', unsafe_allow_html=True)

    # ---------- 转换过程 ----------
    if convert_button and uploaded_file:
        with st.spinner("正在转换中，请稍候..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, uploaded_file.name)
                output_filename = f"转换后的_{uploaded_file.name}"
                output_path = os.path.join(tmpdir, output_filename)
                report_path = os.path.join(tmpdir, "conversion_report.xlsx")

                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                try:
                    registry = TemplateRegistry()
                    replayer = ContentReplayer(registry, template_path=TEMPLATE_PATH)

                    out_path, qa_items = replayer.convert_with_classification(
                        source_path=input_path,
                        output_path=output_path,
                        background_style=selected_style,
                    )

                    # 生成 QA 报告
                    style_name = next((s["name"] for s in BACKGROUND_STYLES if s["key"] == selected_style), "未知")
                    reporter = QAReporter()
                    reporter.generate(qa_items, report_path, summary={
                        "源文件": uploaded_file.name,
                        "背景风格": style_name,
                        "总页数": len(qa_items),
                        "Migration 页数": sum(1 for q in qa_items if q.migration_mode == "migration"),
                        "Adaptation 页数": sum(1 for q in qa_items if q.migration_mode == "adaptation"),
                        "需人工检查页数": sum(1 for q in qa_items if q.need_manual_review),
                    })

                    st.success(f"转换成功！共 {len(qa_items)} 页")

                    # 统计信息
                    m_count = sum(1 for q in qa_items if q.migration_mode == "migration")
                    a_count = sum(1 for q in qa_items if q.migration_mode == "adaptation")
                    review_count = sum(1 for q in qa_items if q.need_manual_review)
                    st.info(f"Migration: {m_count} 页 | Adaptation: {a_count} 页 | 需人工检查: {review_count} 页")

                    if review_count > 0:
                        st.warning(f"有 {review_count} 页建议人工检查，请查看 QA 报告")

                    # 下载按钮
                    dl_col1, dl_col2 = st.columns(2)
                    with dl_col1:
                        with open(output_path, "rb") as f:
                            st.download_button(
                                label="下载转换后的 PPT",
                                data=f.read(),
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                                type="primary",
                            )
                    with dl_col2:
                        with open(report_path, "rb") as f:
                            st.download_button(
                                label="下载 QA 报告",
                                data=f.read(),
                                file_name="conversion_report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                            )

                    # 保存历史记录
                    _add_conversion_record(
                        input_name=uploaded_file.name,
                        output_name=output_filename,
                        background_style=selected_style,
                        slide_count=len(qa_items),
                    )

                except Exception as e:
                    st.error(f"转换出错: {str(e)}")
                    import traceback
                    with st.expander("查看详细错误"):
                        st.code(traceback.format_exc())


# ============================================
#   页面 2: 历史记录
# ============================================
elif current_page == "history":
    st.markdown('<h1 class="page-title">历史记录</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">查看和管理您的转换历史</p>', unsafe_allow_html=True)

    history = _load_history()

    if not history:
        st.markdown("""
        <div class="config-card">
            <div style="text-align: center; padding: 40px 20px;">
                <div style="font-size: 48px; margin-bottom: 16px; opacity: 0.4;">📋</div>
                <div style="font-size: 16px; color: #4A5568; font-weight: 500; margin-bottom: 8px;">暂无转换记录</div>
                <div style="font-size: 13px; color: #718096;">完成第一次转换后，记录将显示在这里</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("去转换 PPT", type="primary", key="go_convert"):
            st.session_state["current_page"] = "convert"
            st.rerun()
    else:
        total_count = len(history)
        st.markdown(f"""
        <div class="config-card" style="margin-bottom: 20px;">
            <div style="display: flex; gap: 24px; align-items: center;">
                <div>
                    <div style="font-size: 12px; color: #718096; margin-bottom: 4px;">总转换次数</div>
                    <div style="font-size: 24px; font-weight: 700; color: #3DCD58;">{total_count}</div>
                </div>
                <div>
                    <div style="font-size: 12px; color: #718096; margin-bottom: 4px;">最近转换</div>
                    <div style="font-size: 14px; font-weight: 500; color: #1A1A1A;">{history[0]['timestamp']}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns([6, 1])
        with col2:
            if st.button("清空记录", type="secondary", key="clear_history"):
                _save_history([])
                st.rerun()

        for idx, record in enumerate(history):
            with st.container():
                st.markdown(f"""
                <div class="config-card" style="margin-bottom: 12px; padding: 16px;">
                    <div style="display: flex; justify-content: space-between; align-items: flex-start; gap: 16px;">
                        <div style="flex: 1; min-width: 0;">
                            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                                <span style="font-size: 20px;">📄</span>
                                <span style="font-size: 15px; font-weight: 600; color: #1A1A1A; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">{record['input_name']}</span>
                            </div>
                            <div style="display: flex; flex-wrap: wrap; gap: 16px; font-size: 13px; color: #718096;">
                                <span>⏰ {record['timestamp']}</span>
                                <span>🎨 {record.get('style_name', '未知风格')}</span>
                                <span>📊 {record['slide_count']} 页</span>
                            </div>
                            <div style="margin-top: 8px; font-size: 12px; color: #A0AEC0;">
                                输出文件: {record['output_name']}
                            </div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 4])
                with btn_col1:
                    if st.button("重新转换", key=f"reconvert_{record['id']}"):
                        st.session_state["current_page"] = "convert"
                        st.session_state["background_style"] = record.get("background_style", "dark_green")
                        st.rerun()
                with btn_col2:
                    if st.button("删除", key=f"delete_{record['id']}"):
                        history.pop(idx)
                        _save_history(history)
                        st.rerun()


# ============================================
#   页面 3: 使用说明
# ============================================
elif current_page == "help":
    st.markdown('<h1 class="page-title">使用说明</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">了解如何使用 PPT 标准模板转换工具</p>', unsafe_allow_html=True)

    st.markdown("""
    <div class="step-list">
        <div class="step-item">
            <div class="step-number">1</div>
            <div class="step-content">
                <div class="step-title">上传待转换 PPT</div>
                <div class="step-desc">在「模板转换」页面上传需要转换格式的 PPT 文件（.pptx 格式）</div>
            </div>
        </div>
        <div class="step-item">
            <div class="step-number">2</div>
            <div class="step-content">
                <div class="step-title">查看质量分析</div>
                <div class="step-desc">系统自动分析 PPT 质量，包括页面数量、母版数量、字体、动画、SmartArt 等</div>
            </div>
        </div>
        <div class="step-item">
            <div class="step-number">3</div>
            <div class="step-content">
                <div class="step-title">选择背景风格</div>
                <div class="step-desc">从白色、浅绿色、深绿色、渐变色 4 种背景风格中选择一种</div>
            </div>
        </div>
        <div class="step-item">
            <div class="step-number">4</div>
            <div class="step-content">
                <div class="step-title">开始转换</div>
                <div class="step-desc">点击「开始转换」按钮，系统自动识别页面类型并映射到对应 Layout</div>
            </div>
        </div>
        <div class="step-item">
            <div class="step-number">5</div>
            <div class="step-content">
                <div class="step-title">下载结果</div>
                <div class="step-desc">下载转换后的 PPT 和 QA 报告（conversion_report.xlsx），检查需要人工复核的页面</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="config-card" style="margin-top: 24px;">', unsafe_allow_html=True)
    st.markdown('<div class="config-card-title">关于混合模式转换</div>', unsafe_allow_html=True)

    st.markdown("""
    <div style="font-size: 14px; color: #4A5568; line-height: 1.8;">
        <p style="margin: 0 0 12px 0;">本工具采用<strong>混合模式转换</strong>，根据页面类型自动选择最优路径：</p>
        <ul style="margin: 0; padding-left: 20px;">
            <li style="margin-bottom: 8px;"><strong>Content Migration</strong>：封面、章节页、结尾页等简单页面从零创建，只迁移文本内容，彻底避免旧格式残留</li>
            <li style="margin-bottom: 8px;"><strong>Technical Adaptation</strong>：含图表、表格、SmartArt 的复杂页面保留内容，应用品牌样式标准化</li>
            <li style="margin-bottom: 8px;"><strong>字体标准化</strong>：中文字体替换为汉仪旗黑，英文字体替换为 Poppins</li>
            <li style="margin-bottom: 8px;"><strong>水印清除</strong>：自动检测并删除文本水印和图片水印</li>
            <li><strong>QA 报告</strong>：生成 Excel 报告，标记需要人工检查的页面</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    version_text = f"版本: <code>{APP_VERSION}</code>"
    if APP_VERSION_DATE:
        version_text += f" &nbsp;|&nbsp; 更新日期: {APP_VERSION_DATE}"

    st.markdown(f"""
    <div style="margin-top: 24px; padding: 16px 20px; background: #F7FAFC; border-radius: 10px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
        <div style="font-size: 12px; color: #718096;">{version_text}</div>
        <div style="font-size: 12px; color: #A0AEC0;">Schneider Electric · PPT 标准模板转换工具</div>
    </div>
    """, unsafe_allow_html=True)


# ============================================
#   全局底部版本信息
# ============================================
version_bar_text = f"Commit: <code>{APP_VERSION}</code>"
if APP_VERSION_DATE:
    version_bar_text += f" &nbsp;|&nbsp; {APP_VERSION_DATE}"

st.markdown(f"""
<style>
.version-bar {{
    position: fixed;
    bottom: 0;
    left: var(--sidebar-width, 240px);
    right: 0;
    padding: 10px 24px;
    background: #FFFFFF;
    border-top: 1px solid #E2E8F0;
    font-size: 12px;
    color: #718096;
    display: flex;
    justify-content: center;
    align-items: center;
    z-index: 100;
    box-shadow: 0 -1px 3px rgba(0,0,0,0.04);
}}

@media (max-width: 768px) {{
    .version-bar {{
        left: 0;
        padding: 8px 16px;
    }}
}}
</style>
<div class="version-bar">
    {version_bar_text}
</div>
""", unsafe_allow_html=True)
