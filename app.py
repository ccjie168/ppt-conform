import streamlit as st
import tempfile
import os
from pathlib import Path

from core.extractor.pptx_extractor import PptxExtractor
from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.validator.validator import Validator
from core.models import UserConfig

st.set_page_config(
    page_title="PPT 标准模板转换智能体",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PPT 标准模板转换智能体")
st.markdown("将 Trae/豆包生成的 PPT 按照公司标准模板进行转换，自动去除水印，确保品牌一致性")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 上传 PPT 文件")
    uploaded_file = st.file_uploader("选择要转换的 PPT 文件", type=["pptx"])

with col2:
    st.subheader("⚙️ 转换配置")
    master_style = st.selectbox(
        "选择模板风格",
        [("F1", "白色简约"), ("F2", "浅绿色清新"), ("F3", "深绿色商务"), ("F4", "渐变科技")],
        format_func=lambda x: x[1]
    )
    include_header = st.checkbox("包含页眉", value=False)
    include_footer = st.checkbox("包含页脚", value=False)
    include_icon = st.checkbox("包含图标", value=False)

if uploaded_file is not None:
    st.subheader("📋 文件信息")
    file_details = {
        "文件名": uploaded_file.name,
        "大小": f"{uploaded_file.size / 1024:.1f} KB",
        "类型": uploaded_file.type
    }
    st.write(file_details)

convert_button = st.button("🚀 开始转换", disabled=uploaded_file is None)

if convert_button and uploaded_file:
    with st.spinner("正在转换中..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, uploaded_file.name)
            output_filename = f"转换后的_{uploaded_file.name}"
            output_path = os.path.join(tmpdir, output_filename)

            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

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

                st.info("📄 步骤2: 加载公司模板...")
                registry = TemplateRegistry()
                registry.load_master_styles()
                registry.load_layout_mappings()

                st.info("✏️ 步骤3: 重放内容到新模板...")
                replayer = ContentReplayer(registry)
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

st.markdown("---")
st.markdown("### ℹ️ 关于应用")
st.markdown("""
- **水印去除**: 自动检测并去除 "TRAE AI 生成"、"豆包 AI" 等水印
- **四种风格**: 白色简约 / 浅绿色清新 / 深绿色商务 / 渐变科技
- **质量校验**: 水印检测、字体白名单、布局有效性等多维度校验
- **失败阻断**: 校验失败时不输出任何文件，确保产物质量
""")