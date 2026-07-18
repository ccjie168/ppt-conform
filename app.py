import streamlit as st
import tempfile
import os
import json
from pathlib import Path

from core.extractor.pptx_extractor import PptxExtractor
from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.validator.validator import Validator
from core.models import UserConfig
from core.analyzer import TemplateAnalyzer

st.set_page_config(
    page_title="PPT 标准模板转换智能体",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PPT 标准模板转换智能体")
st.markdown("将 Trae/豆包生成的 PPT 按照公司标准模板进行转换，自动去除水印，确保品牌一致性")


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
    with open(LAST_TEMPLATE_FILE, "wb") as f:
        f.write(template_file.getbuffer())


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
    return os.path.exists(LAST_TEMPLATE_FILE)


def _get_last_template_info():
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
        with open(LAST_TEMPLATE_FILE, "rb") as f:
            file_content = f.read()

        class MockFile:
            name = config.get("template_name", "上次使用的模板") if config else "上次使用的模板"
            size = os.path.getsize(LAST_TEMPLATE_FILE)

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
st.markdown("## 🎨 上传公司标准模板（全局共享）")
st.markdown("上传一次模板，下方「PPT 转换」和「模板分析」两个页签将共用此模板，无需重复上传。")
st.markdown("应用会自动记住上次使用的模板，下次打开页面时直接使用。")

col_upload, col_reload = st.columns([3, 1])

with col_upload:
    global_template = st.file_uploader(
        "选择公司标准模板 PPT",
        type=["pptx"],
        key="global_template",
        help="此模板将在 PPT 转换和模板分析两个页签中共用",
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

if global_template is not None:
    st.success(f"✅ 已上传模板：{global_template.name}（{global_template.size / 1024:.1f} KB）")
    _save_template(global_template)
    _analyze_template_file(global_template)
elif use_last_template:
    mock_file, config = _load_last_template()
    if mock_file:
        st.success(f"✅ 已加载上次模板：{mock_file.name}（{mock_file.size / 1024:.1f} KB）")
        _analyze_template_file(mock_file, is_reload=True)
        global_template = mock_file
else:
    last_info = _get_last_template_info()
    if last_info and _has_last_template():
        st.info(f"📋 上次使用的模板：{last_info['name']}（{last_info['size'] / 1024:.1f} KB）")
        st.info("点击「使用上次模板」按钮快速加载，或上传新模板替换")
    else:
        st.info("👆 请上传公司标准模板，上传后将自动识别其中的风格")

    if st.session_state.get("_template_cache_key") is None and _has_last_template():
        mock_file, config = _load_last_template()
        if mock_file:
            _analyze_template_file(mock_file, is_reload=True)

st.markdown("---")

tab1, tab2 = st.tabs(["🔄 PPT 转换", "🔍 模板分析"])


# ============ Tab 1: PPT 转换 ============
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📁 上传待转换的 PPT")
        uploaded_file = st.file_uploader("选择要转换的 PPT 文件", type=["pptx"], key="input_ppt")

        template_info = st.session_state.get("_template_file")
        if template_info:
            st.markdown(f"📎 **当前使用的标准模板**: {template_info['name']}")
        else:
            st.warning("⚠️ 请先在页面上方上传公司标准模板")

    with col2:
        st.subheader("⚙️ 转换配置")

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
            st.markdown("### 🎨 选择模板风格（点击卡片选择）")
            st.markdown("根据上传的公司模板，已识别出以下风格：")

            num_cols = min(2, len(master_options))
            cols = st.columns(num_cols)
            selected_master_index = st.session_state.get("selected_master", 0)

            for i, option in enumerate(master_options):
                with cols[i % num_cols]:
                    bg_display = option.get("display_color", "")

                    if bg_display:
                        bg_hex = f"#{bg_display}" if bg_display and not bg_display.startswith("#") else bg_display
                        text_color = "#000000" if _is_light_color(bg_display) else "#FFFFFF"
                        is_selected = selected_master_index == option["index"]
                        border_color = "#2196F3" if is_selected else "#e0e0e0"
                        border_width = "3px" if is_selected else "1px"
                        shadow_style = "0 4px 12px rgba(33, 150, 243, 0.3)" if is_selected else "0 2px 8px rgba(0,0,0,0.1)"

                        st.markdown(f"""
                        <div style="
                            border: {border_width} solid {border_color};
                            border-radius: 12px;
                            padding: 16px;
                            background-color: {bg_hex};
                            box-shadow: {shadow_style};
                            text-align: center;
                            margin-bottom: 8px;
                        ">
                            <div style="color: {text_color}; font-weight: bold; font-size: 16px; margin-bottom: 4px;">
                                {option['style_name']}
                            </div>
                            <div style="color: {text_color}; font-size: 12px; opacity: 0.9;">
                                {option['style_desc']}
                            </div>
                            {'' if not is_selected else f'<div style="margin-top: 8px; padding: 4px; background-color: rgba(33,150,243,0.3); border-radius: 4px; color: {text_color}; font-size: 11px; display: inline-block;">✓ 已选择</div>'}
                        </div>
                        """, unsafe_allow_html=True)

                        if st.button(
                            f"选择",
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
                    else:
                        is_selected = selected_master_index == option["index"]
                        card_style = "border: 2px solid #2196F3; background-color: #E3F2FD;" if is_selected else "border: 1px solid #e0e0e0; background-color: #f9f9f9;"

                        st.markdown(f"""
                            <div style="border-radius: 12px; padding: 16px; {card_style}; text-align: center; margin-bottom: 8px;">
                                <div style="font-weight: bold; font-size: 16px; margin-bottom: 4px;">
                                    {option['style_name']}
                                </div>
                                <div style="font-size: 12px; color: #666;">
                                    {option['style_desc']}
                                </div>
                            </div>
                        """, unsafe_allow_html=True)

                        if st.button(
                            f"选择",
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

            st.markdown(f"**当前选择**: Master #{selected_master_index} - {master_options[selected_master_index]['style_name']}")
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
        st.subheader("📋 文件信息")
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
    convert_button = st.button("🚀 开始转换", disabled=uploaded_file is None or template_info is None)

    if convert_button and uploaded_file and template_info:
        with st.spinner("正在转换中..."):
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

                try:
                    st.info("🔍 步骤1: 检测并去除水印...")
                    extractor = PptxExtractor()
                    content_models = extractor.extract(input_path)

                    st.info("📄 步骤2: 加载模板...")
                    registry = TemplateRegistry()

                    st.info("✏️ 步骤3: 重放内容到新模板...")
                    replayer = ContentReplayer(registry, template_path=template_path)
                    temp_output = replayer.replay(content_models, config)

                    st.info("✅ 步骤4: 质量校验...")
                    validator = Validator()
                    report = validator.validate(temp_output)

                    if report.passed:
                        st.success(f"🎉 转换成功！共 {len(content_models)} 页")
                        st.info(f"校验结果: {report.summary}")

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
    st.subheader("🔍 分析公司标准模板")
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
st.markdown("### ℹ️ 关于应用")
st.markdown("""
- **模板持久化**: 上传一次模板后自动保存，下次打开页面可直接使用
- **配置记忆**: 记住上次选择的风格、页眉页脚设置
- **模板共享**: 上传一次模板，PPT 转换和模板分析两个页签共用，无需重复上传
- **水印去除**: 自动检测并去除 "TRAE AI 生成"、"豆包 AI" 等水印
- **动态风格**: 上传模板后自动加载其中的 master 风格供选择
- **颜色预览**: 每个风格选项显示实际背景色预览
- **质量校验**: 水印检测、字体白名单、布局有效性等多维度校验
- **失败阻断**: 校验失败时不输出任何文件，确保产物质量
""")
