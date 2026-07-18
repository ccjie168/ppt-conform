import streamlit as st
import tempfile
import os
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

# 创建标签页
tab1, tab2 = st.tabs(["🔄 PPT 转换", "🔍 模板分析"])

# ============ Tab 1: PPT 转换 ============
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📁 上传待转换的 PPT")
        uploaded_file = st.file_uploader("选择要转换的 PPT 文件", type=["pptx"], key="input_ppt")

        st.subheader("🎨 上传标准模板（可选）")
        template_file = st.file_uploader("选择公司标准模板 PPT（不传则使用内置模板）", type=["pptx"], key="template_ppt")

    with col2:
        st.subheader("⚙️ 转换配置")
        master_style = st.selectbox(
            "选择模板风格",
            [("F1", "白色简约"), ("F2", "浅绿色清新"), ("F3", "深绿色商务"), ("F4", "渐变科技")],
            format_func=lambda x: x[1],
            key="master_style"
        )
        include_header = st.checkbox("包含页眉", value=False, key="include_header")
        include_footer = st.checkbox("包含页脚", value=False, key="include_footer")
        include_icon = st.checkbox("包含图标", value=False, key="include_icon")

    if uploaded_file is not None:
        st.subheader("📋 文件信息")
        file_details = {
            "待转换文件": uploaded_file.name,
            "大小": f"{uploaded_file.size / 1024:.1f} KB",
            "类型": uploaded_file.type
        }
        if template_file is not None:
            file_details["标准模板"] = template_file.name
        st.write(file_details)

    convert_button = st.button("🚀 开始转换", disabled=uploaded_file is None)

    if convert_button and uploaded_file:
        with st.spinner("正在转换中..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, uploaded_file.name)
                output_filename = f"转换后的_{uploaded_file.name}"
                output_path = os.path.join(tmpdir, output_filename)
                template_path = None

                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                if template_file is not None:
                    template_path = os.path.join(tmpdir, "custom_template.pptx")
                    with open(template_path, "wb") as f:
                        f.write(template_file.getbuffer())
                    st.info("📄 使用自定义模板...")
                else:
                    st.info("📄 使用内置模板...")

                config = UserConfig(
                    input_path=input_path,
                    output_path=output_path,
                    master_style=master_style[0],
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
                    if os.path.exists(output_path):
                        os.remove(output_path)

# ============ Tab 2: 模板分析 ============
with tab2:
    st.subheader("🔍 分析公司标准模板")
    st.markdown("上传公司模板，自动识别其中的 master 风格，并与程序中的 F1-F4 四种风格进行匹配。")

    analyze_template = st.file_uploader(
        "上传要分析的标准模板 PPT", type=["pptx"], key="analyze_template"
    )

    if st.button("🔬 开始分析", disabled=analyze_template is None):
        if analyze_template is not None:
            with tempfile.TemporaryDirectory() as tmpdir:
                tpl_path = os.path.join(tmpdir, "analyze_template.pptx")
                with open(tpl_path, "wb") as f:
                    f.write(analyze_template.getbuffer())

                try:
                    analyzer = TemplateAnalyzer()
                    result = analyzer.analyze(tpl_path)

                    st.success(f"✅ 分析完成！")
                    st.write(f"- **模板文件**: {analyze_template.name}")
                    st.write(f"- **Master 数量**: {len(result['masters'])}")
                    st.write(f"- **Layout 总数**: {result['total_layouts']}")
                    st.write(f"- **现有 Slide 数**: {result['total_slides']}")

                    # 风格匹配结果
                    st.markdown("### 🎨 风格匹配结果")
                    match_data = []
                    for m in result["style_matches"]:
                        bg = m["background"]
                        bg_str = f"{bg['type']}"
                        if bg.get("color"):
                            bg_str += f" ({bg['color']})"
                        elif bg.get("gradient"):
                            bg_str += f" ({' → '.join(bg['gradient'])})"
                        match_data.append({
                            "Master 序号": m["master_index"],
                            "Master 名称": m["master_name"],
                            "背景": bg_str,
                            "匹配风格": m["matched_style"],
                            "风格名称": m["matched_name"],
                            "置信度": m["confidence"],
                        })
                    st.table(match_data)

                    # 每个 master 的详细信息
                    st.markdown("### 📋 Master 详细信息")
                    for master in result["masters"]:
                        with st.expander(
                            f"Master #{master['index']}: {master['name']}（{len(master['layouts'])} 个 layout）"
                        ):
                            bg = master["background"]
                            bg_str = f"类型: {bg['type']}"
                            if bg.get("color"):
                                bg_str += f" | 颜色: #{bg['color']}"
                            if bg.get("gradient"):
                                bg_str += f" | 渐变: {' → '.join(bg['gradient'])}"
                            st.write(f"**背景**: {bg_str}")
                            st.write(f"**字体**: 主标题={master['fonts'].get('major')} | 正文={master['fonts'].get('minor')}")

                            st.write("**Layouts:**")
                            layout_data = []
                            for layout in master["layouts"]:
                                layout_data.append({
                                    "索引": layout["index"],
                                    "名称": layout["name"],
                                    "推测类型": layout["type_guess"],
                                    "占位符数": len(layout["placeholders"]),
                                })
                            st.table(layout_data)

                    # 匹配建议
                    st.markdown("### 💡 匹配建议")
                    unmatched = [m for m in result["style_matches"] if m["matched_style"] == "?"]
                    matched = [m for m in result["style_matches"] if m["matched_style"] != "?"]

                    if matched:
                        st.write("✅ 已匹配的 master：")
                        for m in matched:
                            st.write(f"- Master #{m['master_index']} → **{m['matched_style']}** ({m['matched_name']})")

                    if unmatched:
                        st.write("⚠️ 未匹配的 master（颜色不在 F1-F4 预设范围内）：")
                        for m in unmatched:
                            bg = m["background"]
                            st.write(
                                f"- Master #{m['master_index']}（背景: {bg.get('color') or bg.get('type')}）"
                                f"—— 可在 config/master_styles.yaml 中扩展风格定义"
                            )

                    if len(result["masters"]) < 4:
                        st.warning(
                            f"模板中只有 {len(result['masters'])} 个 master，"
                            f"但程序定义了 F1-F4 共 4 种风格。建议公司模板包含 4 个 master。"
                        )
                    elif len(result["masters"]) > 4:
                        st.warning(
                            f"模板中有 {len(result['masters'])} 个 master，"
                            f"超过 F1-F4 的 4 种。多余 master 不会被使用。"
                        )

                except Exception as e:
                    st.error(f"❌ 分析出错: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

st.markdown("---")
st.markdown("### ℹ️ 关于应用")
st.markdown("""
- **水印去除**: 自动检测并去除 "TRAE AI 生成"、"豆包 AI" 等水印
- **模板选择**: 可使用内置模板（F1-F4）或上传自定义公司标准模板
- **模板分析**: 自动识别公司模板的 master 风格，匹配 F1-F4
- **四种风格**: 白色简约 / 浅绿色清新 / 深绿色商务 / 渐变科技
- **质量校验**: 水印检测、字体白名单、布局有效性等多维度校验
- **失败阻断**: 校验失败时不输出任何文件，确保产物质量
""")
