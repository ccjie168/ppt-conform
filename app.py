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
    """获取当前 git commit 哈希（短格式）"""
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
    """获取当前 git commit 日期"""
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


def _check_template_aspect_ratio(template_file) -> tuple[bool, str, str]:
    """检查模板尺寸是否合法。
    
    不再强制要求16:9比例，保留原PPT的任意比例进行技术适配。
    
    返回: (是否通过, 比例描述, 错误消息)
    """
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

st.set_page_config(
    page_title="施耐德 PPT 模板转换工具",
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

st.markdown("""
<div class="schneider-header">
    <div class="schneider-logo">SE<span>｜</span></div>
    <div class="schneider-header-text">
        <h1 style="margin:0;padding:0;font-size:24px;font-weight:700;color:#1A1A1A;">PPT 标准模板转换工具</h1>
        <p style="margin:4px 0 0 0;color:#666666;font-size:13px;">Schneider Electric · 品牌一致性 · 技术适配模式</p>
    </div>
</div>
""", unsafe_allow_html=True)


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
    for key in ["_template_content", "_template_name", "_template_size"]:
        if key in st.session_state:
            del st.session_state[key]


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


# ============ 全局模板上传（两个 tab 共用） ============
st.markdown('<div class="section-title"><h2>上传公司标准模板（全局共享）</h2></div>', unsafe_allow_html=True)
st.markdown("上传一次模板，下方「PPT 转换」和「模板分析」两个页签将共用此模板，无需重复上传。应用会自动记住上次使用的模板，下次打开页面时直接使用。")

col_upload, col_reload = st.columns([3, 1])

with col_upload:
    global_template = st.file_uploader(
        "选择公司标准模板 PPT（技术适配模式，保留原PPT比例）",
        type=["pptx"],
        key="global_template",
        help="此模板将在 PPT 转换和模板分析两个页签中共用。技术适配模式保留原PPT的版式和比例，不会强制套用模板。",
    )

with col_reload:
    st.write("")
    st.write("")
    use_last_template = st.button(
        "📋 使用上次模板",
        disabled=not _has_last_template(),
        key="use_last_template",
        help="使用上次上传的模板",
    )

auto_load_done = False

if global_template is not None:
    is_valid, ratio_str, err_msg = _check_template_aspect_ratio(global_template)
    if not is_valid:
        st.error(err_msg)
        # 清理已保存的4:3模板，避免后续误用
        _clear_last_template()
        global_template = None
    else:
        st.success(f"✅ 已上传模板：{global_template.name}（{global_template.size / 1024:.1f} KB） - 技术适配 ({ratio_str})")
        _save_template(global_template)
        _analyze_template_file(global_template)
        auto_load_done = True
elif use_last_template:
    mock_file, config = _load_last_template()
    if mock_file:
        is_valid, ratio_str, err_msg = _check_template_aspect_ratio(mock_file)
        if not is_valid:
            st.error(err_msg)
            _clear_last_template()
        else:
            st.success(f"✅ 已加载上次模板：{mock_file.name}（{mock_file.size / 1024:.1f} KB） - 技术适配 ({ratio_str})")
            _analyze_template_file(mock_file, is_reload=True)
            global_template = mock_file
            auto_load_done = True
else:
    last_info = _get_last_template_info()
    if last_info and _has_last_template():
        st.info(f"📋 上次使用的模板：{last_info['name']}（{last_info['size'] / 1024:.1f} KB）")
        st.info("点击「使用上次模板」按钮快速加载，或上传新模板替换")
    else:
        st.info("👆 请上传公司标准模板（仅支持 16:9 宽屏），上传后将自动识别其中的风格")

    if st.session_state.get("_template_cache_key") is None and _has_last_template():
        mock_file, config = _load_last_template()
        if mock_file:
            is_valid, ratio_str, err_msg = _check_template_aspect_ratio(mock_file)
            if not is_valid:
                st.error(err_msg)
                _clear_last_template()
            else:
                _analyze_template_file(mock_file, is_reload=True)
                global_template = mock_file
                auto_load_done = True

if auto_load_done and global_template:
    template_info = st.session_state.get("_template_file")
    if template_info:
        st.markdown(f"📎 **当前使用的标准模板**: {template_info['name']}")

st.markdown("---")

tab1, tab2 = st.tabs(["  🔄 PPT 转换  ", "  🔍 模板分析  "])


# ============ Tab 1: PPT 转换 ============
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="section-title"><h3>上传待转换的 PPT</h3></div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("选择要转换的 PPT 文件", type=["pptx"], key="input_ppt")

        template_info = st.session_state.get("_template_file")
        if template_info:
            st.markdown(f"📎 **当前使用的标准模板**: {template_info['name']}")
        else:
            st.warning("⚠️ 请先在页面上方上传公司标准模板")

    with col2:
        st.markdown('<div class="section-title"><h3>转换配置</h3></div>', unsafe_allow_html=True)

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

        if master_options:
            st.markdown('<div class="section-title"><h3>选择模板风格</h3></div>', unsafe_allow_html=True)
            st.markdown("根据上传的公司模板，已识别出以下风格，点击卡片选择：")

            num_cols = min(2, len(master_options))
            cols = st.columns(num_cols)
            selected_master_index = st.session_state.get("selected_master", 0)

            for i, option in enumerate(master_options):
                with cols[i % num_cols]:
                    bg_display = option.get("display_color", "")
                    is_selected = selected_master_index == option["index"]
                    selected_class = "selected" if is_selected else ""

                    if bg_display:
                        bg_hex = f"#{bg_display}" if bg_display and not bg_display.startswith("#") else bg_display
                        text_color = "#1A1A1A" if _is_light_color(bg_display) else "#FFFFFF"
                        
                        st.markdown(f"""
                        <div class="template-card {selected_class}">
                            <div class="template-card-preview" style="background-color: {bg_hex};">
                                {f'<div class="template-card-badge">已选择</div>' if is_selected else ''}
                                <div style="color:{text_color};font-size:13px;font-weight:600;opacity:0.9;">
                                    {option['bg_type'].upper() if option.get('bg_type') else ''}
                                </div>
                            </div>
                            <div class="template-card-content">
                                <div class="template-card-name">{option['style_name']}</div>
                                <div class="template-card-desc">{option['style_desc']}</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div class="template-card {selected_class}">
                            <div class="template-card-preview" style="background-color: #F5F5F5;">
                                {f'<div class="template-card-badge">已选择</div>' if is_selected else ''}
                                <div style="color:#999;font-size:13px;">预览</div>
                            </div>
                            <div class="template-card-content">
                                <div class="template-card-name">{option['style_name']}</div>
                                <div class="template-card-desc">{option['style_desc']}</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                    if st.button(
                        f"选择此风格",
                        key=f"master_btn_{option['index']}",
                        use_container_width=True,
                        type="primary" if is_selected else "secondary",
                    ):
                        _save_config({
                            "template_name": st.session_state.get("_template_file", {}).get("name", ""),
                            "selected_master": option["index"],
                            "include_header": st.session_state.get("include_header", False),
                            "include_footer": st.session_state.get("include_footer", False),
                            "include_icon": st.session_state.get("include_icon", False),
                        })
                        st.session_state["selected_master"] = option["index"]
                        st.rerun()

            st.info(f"**当前选择**: Master #{selected_master_index} - {master_options[selected_master_index]['style_name']}")
        else:
            st.warning("请先在页面上方上传公司标准模板，风格列表将自动加载")
            selected_master_index = 0

        include_header = st.checkbox("包含页眉", value=st.session_state.get("include_header", False), key="include_header")
        include_footer = st.checkbox("包含页脚", value=st.session_state.get("include_footer", False), key="include_footer")
        include_icon = st.checkbox("包含图标", value=st.session_state.get("include_icon", False), key="include_icon")

        _save_config({
            "template_name": st.session_state.get("_template_file", {}).get("name", ""),
            "selected_master": selected_master_index,
            "include_header": include_header,
            "include_footer": include_footer,
            "include_icon": include_icon,
        })

    if uploaded_file is not None:
        st.markdown('<div class="section-title"><h3>文件信息</h3></div>', unsafe_allow_html=True)
        file_details = {
            "待转换文件": uploaded_file.name,
            "大小": f"{uploaded_file.size / 1024:.1f} KB",
            "类型": uploaded_file.type
        }
        template_info = st.session_state.get("_template_file")
        if template_info:
            file_details["标准模板"] = template_info["name"]
        st.write(file_details)

    template_info = st.session_state.get("_template_file")
    convert_button = st.button("⚡ 开始转换", disabled=uploaded_file is None or template_info is None, type="primary")

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

                st.markdown('<div class="step-indicator"><div class="step-number">0</div><div class="step-text">源文件检查...</div></div>', unsafe_allow_html=True)
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
                        for issue in source_warn[:5]:
                            st.write(f"  - [{issue.level}] {issue.rule_id}: {issue.message}")
                    try:
                        st.markdown('<div class="step-indicator"><div class="step-number">1</div><div class="step-text">检测并去除水印...</div></div>', unsafe_allow_html=True)
                        extractor = PptxExtractor()
                        content_models = extractor.extract(input_path)

                        st.markdown('<div class="step-indicator"><div class="step-number">2</div><div class="step-text">加载模板...</div></div>', unsafe_allow_html=True)
                        registry = TemplateRegistry()

                        st.markdown('<div class="step-indicator"><div class="step-number">3</div><div class="step-text">重放内容到新模板...</div></div>', unsafe_allow_html=True)
                        replayer = ContentReplayer(registry, template_path=template_path)
                        temp_output = replayer.replay(content_models, config)

                        st.markdown('<div class="step-indicator"><div class="step-number">4</div><div class="step-text">质量校验...</div></div>', unsafe_allow_html=True)
                        report = validator.validate(temp_output)

                        fail_issues = [i for i in report.issues if i.level == "fail"]
                        warn_issues = [i for i in report.issues if i.level == "warning"]

                        if not fail_issues:
                            st.success(f"🎉 转换成功！共 {len(content_models)} 页")
                            st.info(f"校验结果: {report.summary}")

                            if warn_issues:
                                st.warning(f"⚠️ 发现 {len(warn_issues)} 项警告，建议检查")
                                for issue in warn_issues[:5]:
                                    st.write(f"  - [{issue.level}] {issue.rule_id}: {issue.message}")

                            with open(output_path, "rb") as f:
                                st.download_button(
                                    label="📥 下载转换后的 PPT",
                                    data=f,
                                    file_name=output_filename,
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                )
                        else:
                            st.error("❌ 校验失败，未生成输出文件")
                            st.markdown("### 失败详情")
                            for issue in report.issues:
                                st.write(f"**[{issue.level}] {issue.rule_id}**: {issue.message}")
                            if os.path.exists(output_path):
                                os.remove(output_path)

                    except Exception as e:
                        st.error(f"❌ 转换出错: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())
                        if os.path.exists(output_path):
                            os.remove(output_path)

# ============ Tab 2: 模板分析 ============
with tab2:
    st.markdown('<div class="section-title"><h3>分析公司标准模板</h3></div>', unsafe_allow_html=True)
    st.markdown("使用页面上方上传的公司标准模板，自动识别其中的 master 和 layout 结构。")

    template_info = st.session_state.get("_template_file")
    if template_info is None:
        st.warning("⚠️ 请先在页面上方上传公司标准模板")
    else:
        st.markdown(f"📎 **当前分析模板**: {template_info['name']}")

        result = st.session_state.get("_template_analysis")

        if result is not None:
            st.success(f"✅ 分析完成！")
            st.write(f"- **模板文件**: {template_info['name']}")
            st.write(f"- **Master 数量**: {len(result['masters'])}")
            st.write(f"- **Layout 总数**: {result['total_layouts']}")
            st.write(f"- **现有 Slide 数**: {result['total_slides']}")

            if result.get("theme_colors"):
                st.markdown("### 🎨 主题颜色（从模板提取）")
                color_data = []
                for name, hex_color in result["theme_colors"].items():
                    color_data.append({
                        "颜色名称": name,
                        "HEX 值": f"#{hex_color}",
                    })
                st.table(color_data)

            st.markdown("### 📋 Master 列表")
            master_data = []
            for m in result["masters"]:
                bg = m["background"]
                bg_str = bg["type"]
                if bg.get("color"):
                    bg_str += f" (#{bg['color']})"
                elif bg.get("gradient"):
                    bg_str += f" ({' → '.join(bg['gradient'])})"
                if bg.get("theme_color"):
                    bg_str += f" [主题: {bg['theme_color']}]"
                master_data.append({
                    "Master 序号": m["index"],
                    "Master 名称": m["name"],
                    "风格": m.get("style_name", "未知"),
                    "背景": bg_str,
                    "描述": m.get("style_desc", ""),
                    "Layout 数量": len(m["layouts"]),
                })
            st.table(master_data)

            st.markdown("### 📝 Layout 详细信息")
            for master in result["masters"]:
                with st.expander(
                    f"Master #{master['index']}: {master['name']}（{len(master['layouts'])} 个 layout）"
                ):
                    layout_data = []
                    for layout in master["layouts"]:
                        layout_data.append({
                            "索引": layout["index"],
                            "名称": layout["name"],
                            "推测类型": layout["type_guess"],
                            "占位符数": len(layout["placeholders"]),
                        })
                    st.table(layout_data)
        else:
            st.error("❌ 模板分析失败，请检查模板文件是否有效")

st.markdown("---")
st.markdown('<div class="section-title"><h3>关于应用</h3></div>', unsafe_allow_html=True)

col_about_1, col_about_2, col_about_3 = st.columns(3)

with col_about_1:
    st.markdown("""
    <div class="schneider-card" style="padding: 20px;">
        <div style="font-size: 18px; font-weight: 700; color: #3DCD58; margin-bottom: 8px;">⚙️ 技术适配</div>
        <div style="font-size: 13px; color: #666; line-height: 1.6;">
            基于原PPT创建输出，保留原幻灯片和版式结构，通过技术适配调整样式，不直接套用模板。
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_about_2:
    st.markdown("""
    <div class="schneider-card" style="padding: 20px;">
        <div style="font-size: 18px; font-weight: 700; color: #3DCD58; margin-bottom: 8px;">🎨 品牌一致</div>
        <div style="font-size: 13px; color: #666; line-height: 1.6;">
            自动提取模板的字体、颜色、页脚等品牌元素，确保输出PPT符合施耐德品牌规范。
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_about_3:
    st.markdown("""
    <div class="schneider-card" style="padding: 20px;">
        <div style="font-size: 18px; font-weight: 700; color: #3DCD58; margin-bottom: 8px;">✅ 质量校验</div>
        <div style="font-size: 13px; color: #666; line-height: 1.6;">
            多维度质量校验包括水印检测、字体白名单、布局有效性等，确保产物质量。
        </div>
    </div>
    """, unsafe_allow_html=True)

version_text = f"版本: <code>{APP_VERSION}</code>"
if APP_VERSION_DATE:
    version_text += f" &nbsp;|&nbsp; 更新: {APP_VERSION_DATE}"

st.markdown(f"""
<div class="schneider-footer">
    <div class="schneider-footer-left">
        {version_text}
    </div>
    <div class="schneider-footer-right">
        <span class="schneider-footer-brand">Schneider Electric</span>
        <span>·</span>
        <span>PPT 标准模板转换工具</span>
    </div>
</div>
""", unsafe_allow_html=True)
