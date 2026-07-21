import streamlit as st
import tempfile
import os
import json
import subprocess
from pathlib import Path
from pptx import Presentation

from core.extractor.pptx_extractor import PptxExtractor
from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.validator.validator import Validator
from core.models import UserConfig
from core.analyzer import TemplateAnalyzer


# 标准16:9宽屏比例（容差）
WIDESCREEN_16_9_RATIO = 16 / 9
ASPECT_RATIO_TOLERANCE = 0.02


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


def _is_light_color(hex_color: str) -> bool:
    try:
        if not hex_color:
            return True
        if hex_color.startswith("#"):
            hex_color = hex_color[1:]
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return luminance > 0.5
    except Exception:
        return True


PERSIST_DIR = os.path.join(os.path.dirname(__file__), ".persist")
LAST_TEMPLATE_FILE = os.path.join(PERSIST_DIR, "last_template.pptx")
LAST_CONFIG_FILE = os.path.join(PERSIST_DIR, "last_config.json")


def _ensure_persist_dir():
    os.makedirs(PERSIST_DIR, exist_ok=True)


def _save_template(template_file):
    _ensure_persist_dir()
    content = template_file.getbuffer()
    with open(LAST_TEMPLATE_FILE, "wb") as f:
        f.write(content)
    st.session_state["_template_content"] = content
    st.session_state["_template_name"] = template_file.name
    st.session_state["_template_size"] = template_file.size


def _save_config(config_data):
    _ensure_persist_dir()
    with open(LAST_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)


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


def _add_conversion_record(input_name, output_name, template_name, master_index, master_name, slide_count, config):
    """添加一条转换记录"""
    from datetime import datetime
    history = _load_history()
    record = {
        "id": datetime.now().strftime("%Y%m%d%H%M%S"),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "input_name": input_name,
        "output_name": output_name,
        "template_name": template_name,
        "master_index": master_index,
        "master_name": master_name,
        "slide_count": slide_count,
        "config": config,
    }
    history.insert(0, record)
    # 最多保留50条记录
    history = history[:50]
    _save_history(history)
    return record


def _load_last_config():
    try:
        with open(LAST_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _has_last_template():
    return os.path.exists(LAST_TEMPLATE_FILE) or "_template_content" in st.session_state


def _get_last_template_info():
    if "_template_name" in st.session_state:
        return {
            "name": st.session_state["_template_name"],
            "size": st.session_state["_template_size"],
        }
    try:
        config = _load_last_config()
        if config:
            return {
                "name": config.get("template_name", "上次使用的模板"),
                "size": os.path.getsize(LAST_TEMPLATE_FILE) if os.path.exists(LAST_TEMPLATE_FILE) else 0,
            }
    except Exception:
        pass
    return None


def _clear_last_template():
    if os.path.exists(LAST_TEMPLATE_FILE):
        os.remove(LAST_TEMPLATE_FILE)
    if os.path.exists(LAST_CONFIG_FILE):
        os.remove(LAST_CONFIG_FILE)
    for key in ["_template_content", "_template_name", "_template_size", "_template_cache_key", "_template_analysis", "_template_file"]:
        if key in st.session_state:
            del st.session_state[key]


def _check_template_aspect_ratio(template_file) -> tuple[bool, str, str]:
    try:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            tmp_path = tmp.name
            if hasattr(template_file, "getbuffer"):
                tmp.write(template_file.getbuffer())
            else:
                tmp.write(template_file)

        prs = Presentation(tmp_path)
        width = prs.slide_width
        height = prs.slide_height
        ratio = width / height if height else 0

        os.unlink(tmp_path)

        ratio_str = f"{width / 914400:.2f} x {height / 914400:.2f} 英寸 (比例 {ratio:.3f})"
        return True, ratio_str, ""
    except Exception as e:
        return False, "", f"模板尺寸检查失败: {str(e)}"


def _analyze_template_file(template_file, is_reload=False) -> dict | None:
    if template_file is None:
        return None

    cache_key = f"{template_file.name}_{template_file.size}"

    if not is_reload and st.session_state.get("_template_cache_key") == cache_key:
        return st.session_state.get("_template_analysis")

    with tempfile.TemporaryDirectory() as tmpdir:
        tpl_path = os.path.join(tmpdir, "shared_template.pptx")
        with open(tpl_path, "wb") as f:
            if hasattr(template_file, "getbuffer"):
                f.write(template_file.getbuffer())
            else:
                f.write(template_file)
        try:
            analyzer = TemplateAnalyzer()
            result = analyzer.analyze(tpl_path)
            st.session_state["_template_cache_key"] = cache_key
            st.session_state["_template_analysis"] = result
            st.session_state["_template_file"] = {
                "name": template_file.name if hasattr(template_file, "name") else "上次使用的模板",
                "size": template_file.size if hasattr(template_file, "size") else os.path.getsize(LAST_TEMPLATE_FILE),
            }
            if not is_reload:
                _save_config({
                    "template_name": st.session_state["_template_file"]["name"],
                    "template_size": st.session_state["_template_file"]["size"],
                    "selected_master": st.session_state.get("selected_master", 0),
                    "include_header": st.session_state.get("include_header", False),
                    "include_footer": st.session_state.get("include_footer", False),
                    "include_icon": st.session_state.get("include_icon", False),
                })
            return result
        except Exception as e:
            st.session_state["_template_cache_key"] = None
            st.session_state["_template_analysis"] = None
            st.session_state["_template_file"] = None
            st.error(f"模板分析失败: {str(e)}")
            return None


def _load_last_template():
    if not _has_last_template():
        return None, None

    try:
        config = _load_last_config()
        
        if "_template_content" in st.session_state:
            file_content = st.session_state["_template_content"]
            template_name = st.session_state.get("_template_name", "上次使用的模板")
            template_size = st.session_state.get("_template_size", len(file_content))
        elif os.path.exists(LAST_TEMPLATE_FILE):
            with open(LAST_TEMPLATE_FILE, "rb") as f:
                file_content = f.read()
            template_name = config.get("template_name", "上次使用的模板") if config else "上次使用的模板"
            template_size = os.path.getsize(LAST_TEMPLATE_FILE)
        else:
            return None, None

        class MockFile:
            name = template_name
            size = template_size

            def getbuffer(self):
                return file_content

        mock_file = MockFile()

        if config:
            st.session_state["selected_master"] = config.get("selected_master", 0)
            st.session_state["include_header"] = config.get("include_header", False)
            st.session_state["include_footer"] = config.get("include_footer", False)
            st.session_state["include_icon"] = config.get("include_icon", False)

        return mock_file, config
    except Exception as e:
        st.warning(f"加载上次模板失败: {str(e)}")
        return None, None


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
    
    # 导航菜单
    st.markdown('<div class="sidebar-nav">', unsafe_allow_html=True)
    
    nav_items = [
        {"icon": "🔄", "label": "模板转换", "key": "convert"},
        {"icon": "📋", "label": "历史记录", "key": "history"},
        {"icon": "📁", "label": "模板管理", "key": "templates"},
        {"icon": "📖", "label": "使用说明", "key": "help"},
    ]
    
    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "convert"
    
    for item in nav_items:
        is_active = st.session_state.get("current_page") == item["key"]
        active_class = "active" if is_active else ""
        
        if st.button(
            f"{item['icon']}  {item['label']}",
            key=f"nav_{item['key']}",
            use_container_width=True,
            type="primary" if is_active else "secondary",
        ):
            st.session_state["current_page"] = item["key"]
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # 底部用户区
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
    
    # ---------- 转换配置卡片 ----------
    st.markdown('<div class="config-card animate-fade-in">', unsafe_allow_html=True)
    st.markdown('<div class="config-card-title">转换配置</div>', unsafe_allow_html=True)
    
    # 风格选择
    master_options = []
    template_analysis = st.session_state.get("_template_analysis")
    
    if template_analysis is not None:
        for m in template_analysis["masters"]:
            bg = m["background"]
            display_color = bg.get("display_color", "")
            if display_color and isinstance(display_color, list):
                display_color = display_color[0] if display_color else ""
            master_options.append({
                "index": m["index"],
                "name": m["name"],
                "style_id": m.get("style_id", "?"),
                "style_name": m.get("style_name", "未知风格"),
                "style_desc": m.get("style_desc", ""),
                "display_color": display_color,
                "bg_type": bg.get("type", ""),
                "bg_color": bg.get("color", ""),
                "bg_gradient": bg.get("gradient", []),
            })
    
    selected_master_index = st.session_state.get("selected_master", 0)
    
    if master_options:
        option_labels = [f"{m['style_name']}" for m in master_options]
        selected_idx = st.selectbox(
            "风格选择",
            options=range(len(option_labels)),
            format_func=lambda i: option_labels[i],
            index=selected_master_index if selected_master_index < len(option_labels) else 0,
            key="style_select",
        )
        selected_master_index = master_options[selected_idx]["index"]
        st.session_state["selected_master"] = selected_master_index
    else:
        st.selectbox(
            "风格选择",
            options=["请先上传模板"],
            index=0,
            disabled=True,
            key="style_select_disabled",
        )
        selected_master_index = 0
    
    # 模板上传
    st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)
    
    template_col1, template_col2 = st.columns([3, 1])
    with template_col1:
        global_template = st.file_uploader(
            "模板选择",
            type=["pptx"],
            key="global_template",
            help="上传公司标准模板 PPT 文件",
        )
    with template_col2:
        st.write("")
        st.write("")
        use_last = st.button(
            "使用上次",
            disabled=not _has_last_template(),
            key="use_last_template",
            use_container_width=True,
        )
    
    if global_template is not None:
        is_valid, ratio_str, err_msg = _check_template_aspect_ratio(global_template)
        if not is_valid:
            st.error(err_msg)
            _clear_last_template()
            global_template = None
        else:
            st.success(f"✅ 模板已加载（{ratio_str}）")
            _save_template(global_template)
            _analyze_template_file(global_template)
    elif use_last:
        mock_file, config = _load_last_template()
        if mock_file:
            is_valid, ratio_str, err_msg = _check_template_aspect_ratio(mock_file)
            if not is_valid:
                st.error(err_msg)
                _clear_last_template()
            else:
                st.success(f"✅ 已加载上次模板（{ratio_str}）")
                _analyze_template_file(mock_file, is_reload=True)
                global_template = mock_file
    else:
        last_info = _get_last_template_info()
        if last_info and _has_last_template():
            st.info(f"📋 上次模板：{last_info['name']}")
            if st.session_state.get("_template_cache_key") is None:
                mock_file, config = _load_last_template()
                if mock_file:
                    _analyze_template_file(mock_file, is_reload=True)
                    global_template = mock_file
    
    # 开关选项
    st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)
    
    include_header = st.checkbox(
        "保留原始图片",
        value=st.session_state.get("include_header", False),
        key="include_header",
    )
    include_footer = st.checkbox(
        "应用页眉与页脚",
        value=st.session_state.get("include_footer", False),
        key="include_footer",
    )
    include_icon = st.checkbox(
        "统一字体样式",
        value=st.session_state.get("include_icon", False),
        key="include_icon",
    )
    
    _save_config({
        "template_name": st.session_state.get("_template_file", {}).get("name", ""),
        "selected_master": selected_master_index,
        "include_header": include_header,
        "include_footer": include_footer,
        "include_icon": include_icon,
    })
    
    # 补充说明
    st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)
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
                if key not in ["current_page"]:
                    del st.session_state[key]
            st.rerun()
    with convert_col:
        template_info = st.session_state.get("_template_file")
        can_convert = uploaded_file is not None and template_info is not None
        convert_button = st.button(
            "开始转换",
            disabled=not can_convert,
            type="primary",
            use_container_width=True,
            key="convert_btn",
        )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # ---------- 转换过程 ----------
    if convert_button and uploaded_file and template_info:
        with st.spinner("正在转换中，请稍候..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, uploaded_file.name)
                output_filename = f"转换后的_{uploaded_file.name}"
                output_path = os.path.join(tmpdir, output_filename)
                template_path = os.path.join(tmpdir, "custom_template.pptx")

                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                if global_template and hasattr(global_template, "getbuffer"):
                    with open(template_path, "wb") as f:
                        f.write(global_template.getbuffer())
                elif os.path.exists(LAST_TEMPLATE_FILE):
                    with open(template_path, "wb") as f:
                        f.write(open(LAST_TEMPLATE_FILE, "rb").read())

                config = UserConfig(
                    input_path=input_path,
                    output_path=output_path,
                    master_style=str(selected_master_index),
                    include_header=include_header,
                    include_footer=include_footer,
                    include_icon=include_icon
                )

                validator = Validator()
                source_report = validator.validate_source(input_path)
                source_fail = [i for i in source_report.issues if i.level == "fail"]
                source_warn = [i for i in source_report.issues if i.level == "warning"]
                
                if source_fail:
                    st.error("❌ 源文件检查失败")
                    for issue in source_fail:
                        st.write(f"  - [{issue.level}] {issue.rule_id}: {issue.message}")
                else:
                    if source_warn:
                        st.warning(f"⚠️ 源文件检查发现 {len(source_warn)} 项警告，将继续转换")
                    
                    try:
                        extractor = PptxExtractor()
                        content_models = extractor.extract(input_path)

                        registry = TemplateRegistry()
                        replayer = ContentReplayer(registry, template_path=template_path)
                        temp_output = replayer.replay(content_models, config)

                        report = validator.validate(temp_output)

                        fail_issues = [i for i in report.issues if i.level == "fail"]
                        warn_issues = [i for i in report.issues if i.level == "warning"]

                        if not fail_issues:
                            st.success(f"🎉 转换成功！共 {len(content_models)} 页")
                            st.info(f"校验结果: {report.summary}")

                            if warn_issues:
                                st.warning(f"⚠️ 发现 {len(warn_issues)} 项警告，建议检查")

                            with open(output_path, "rb") as f:
                                output_bytes = f.read()
                                st.download_button(
                                    label="📥 下载转换后的 PPT",
                                    data=output_bytes,
                                    file_name=output_filename,
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                    use_container_width=True,
                                    type="primary",
                                )

                            # 保存历史记录
                            template_name = template_info.get("name", "未知模板") if template_info else "未知模板"
                            master_name = "未知风格"
                            if template_analysis and selected_master_index < len(template_analysis.get("masters", [])):
                                master_name = template_analysis["masters"][selected_master_index].get("style_name", "未知风格")

                            _add_conversion_record(
                                input_name=uploaded_file.name,
                                output_name=output_filename,
                                template_name=template_name,
                                master_index=selected_master_index,
                                master_name=master_name,
                                slide_count=len(content_models),
                                config={
                                    "include_header": include_header,
                                    "include_footer": include_footer,
                                    "include_icon": include_icon,
                                },
                            )
                        else:
                            st.error("❌ 校验失败，未生成输出文件")
                            if os.path.exists(output_path):
                                os.remove(output_path)

                    except Exception as e:
                        st.error(f"❌ 转换出错: {str(e)}")
                        import traceback
                        with st.expander("查看详细错误"):
                            st.code(traceback.format_exc())
                        if os.path.exists(output_path):
                            os.remove(output_path)


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
        # 统计信息
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

        # 清空历史按钮
        col1, col2 = st.columns([6, 1])
        with col2:
            if st.button("清空记录", type="secondary", key="clear_history"):
                _save_history([])
                st.rerun()

        # 历史记录列表
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
                                <span>🎨 {record['master_name']}</span>
                                <span>📊 {record['slide_count']} 页</span>
                                <span>📁 {record['template_name']}</span>
                            </div>
                            <div style="margin-top: 8px; font-size: 12px; color: #A0AEC0;">
                                输出文件: {record['output_name']}
                            </div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # 操作按钮
                btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 4])
                with btn_col1:
                    if st.button("重新转换", key=f"reconvert_{record['id']}"):
                        st.session_state["current_page"] = "convert"
                        st.session_state["selected_master"] = record["master_index"]
                        st.session_state["include_header"] = record["config"].get("include_header", False)
                        st.session_state["include_footer"] = record["config"].get("include_footer", False)
                        st.session_state["include_icon"] = record["config"].get("include_icon", False)
                        st.rerun()
                with btn_col2:
                    if st.button("删除", key=f"delete_{record['id']}"):
                        history.pop(idx)
                        _save_history(history)
                        st.rerun()


# ============================================
#   页面 3: 模板管理
# ============================================
elif current_page == "templates":
    st.markdown('<h1 class="page-title">模板管理</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">上传和管理公司标准模板</p>', unsafe_allow_html=True)
    
    template_info = st.session_state.get("_template_file")
    
    if template_info:
        st.markdown(f'<div class="config-card">', unsafe_allow_html=True)
        st.markdown('<div class="config-card-title">当前模板</div>', unsafe_allow_html=True)
        
        st.markdown(f"""
        <div style="display: flex; align-items: center; gap: 16px; padding: 12px 0;">
            <div style="width: 48px; height: 48px; border-radius: 8px; background: #E8F8EC; display: flex; align-items: center; justify-content: center; font-size: 20px;">📄</div>
            <div style="flex: 1;">
                <div style="font-size: 15px; font-weight: 600; color: #1A1A1A;">{template_info['name']}</div>
                <div style="font-size: 13px; color: #718096;">{template_info['size'] / 1024:.1f} KB</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        result = st.session_state.get("_template_analysis")
        if result:
            st.markdown(f"""
            <div style="display: flex; gap: 24px; padding-top: 12px; border-top: 1px solid #EDF2F7; margin-top: 12px;">
                <div>
                    <div style="font-size: 12px; color: #718096; margin-bottom: 4px;">Master 数量</div>
                    <div style="font-size: 18px; font-weight: 700; color: #3DCD58;">{len(result['masters'])}</div>
                </div>
                <div>
                    <div style="font-size: 12px; color: #718096; margin-bottom: 4px;">Layout 总数</div>
                    <div style="font-size: 18px; font-weight: 700; color: #1A1A1A;">{result['total_layouts']}</div>
                </div>
                <div>
                    <div style="font-size: 12px; color: #718096; margin-bottom: 4px;">Slide 数量</div>
                    <div style="font-size: 18px; font-weight: 700; color: #1A1A1A;">{result['total_slides']}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.button("更换模板", key="change_template"):
            _clear_last_template()
            st.rerun()
    else:
        st.markdown("""
        <div class="config-card">
            <div style="text-align: center; padding: 40px 20px;">
                <div style="font-size: 48px; margin-bottom: 16px; opacity: 0.4;">📁</div>
                <div style="font-size: 16px; color: #4A5568; font-weight: 500; margin-bottom: 8px;">尚未上传模板</div>
                <div style="font-size: 13px; color: #718096;">上传公司标准模板后可在此处管理</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        new_template = st.file_uploader(
            "上传公司标准模板",
            type=["pptx"],
            key="template_upload_page",
        )
        
        if new_template is not None:
            is_valid, ratio_str, err_msg = _check_template_aspect_ratio(new_template)
            if not is_valid:
                st.error(err_msg)
            else:
                st.success(f"✅ 模板上传成功（{ratio_str}）")
                _save_template(new_template)
                _analyze_template_file(new_template)
                st.rerun()


# ============================================
#   页面 4: 使用说明
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
                <div class="step-title">选择目标模板</div>
                <div class="step-desc">上传公司标准模板 PPT，选择需要应用的模板风格（Master）</div>
            </div>
        </div>
        <div class="step-item">
            <div class="step-number">3</div>
            <div class="step-content">
                <div class="step-title">配置转换选项</div>
                <div class="step-desc">根据需要设置保留原始图片、应用页眉页脚、统一字体样式等选项</div>
            </div>
        </div>
        <div class="step-item">
            <div class="step-number">4</div>
            <div class="step-content">
                <div class="step-title">开始转换</div>
                <div class="step-desc">点击「开始转换」按钮，等待转换完成后下载结果</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="config-card" style="margin-top: 24px;">', unsafe_allow_html=True)
    st.markdown('<div class="config-card-title">关于技术适配模式</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div style="font-size: 14px; color: #4A5568; line-height: 1.8;">
        <p style="margin: 0 0 12px 0;">本工具采用<strong>技术适配模式</strong>，而非传统的模板套用模式：</p>
        <ul style="margin: 0; padding-left: 20px;">
            <li style="margin-bottom: 8px;"><strong>保留原结构</strong>：保留原 PPT 的幻灯片和版式结构，不强制套用模板</li>
            <li style="margin-bottom: 8px;"><strong>样式提取</strong>：自动提取目标模板的字体、颜色、页脚等品牌元素</li>
            <li style="margin-bottom: 8px;"><strong>智能转换</strong>：占位符元素按模板规范转换，非占位符元素保留原内容</li>
            <li><strong>质量校验</strong>：多维度质量校验确保转换结果符合规范</li>
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
