from core.extractor.pptx_extractor import PptxExtractor
from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.validator.validator import Validator
from core.models import UserConfig


def convert_ppt(input_path: str, output_path: str, master_style: str = "F1",
                include_header: bool = False, include_footer: bool = False,
                include_icon: bool = False) -> dict:
    config = UserConfig(
        input_path=input_path,
        output_path=output_path,
        master_style=master_style,
        include_header=include_header,
        include_footer=include_footer,
        include_icon=include_icon
    )

    try:
        extractor = PptxExtractor()
        content_models = extractor.extract(input_path)

        registry = TemplateRegistry()
        
        # 获取模板路径
        template_path = None
        import os
        template_dir = os.path.join(os.path.dirname(__file__), "..", "templates")
        template_path = os.path.join(template_dir, "corporate_template.pptx")
        if not os.path.exists(template_path):
            template_path = None

        replayer = ContentReplayer(registry, template_path)
        temp_output = replayer.replay(content_models, config)

        validator = Validator()
        report = validator.validate(temp_output)

        if report.passed:
            return {
                "success": True,
                "output_path": output_path,
                "pages": len(content_models),
                "validation_summary": report.summary
            }
        else:
            import os
            if os.path.exists(temp_output):
                os.remove(temp_output)
            return {
                "success": False,
                "error": "校验失败",
                "validation_issues": [
                    {"rule_id": i.rule_id, "level": i.level, "message": i.message}
                    for i in report.issues
                ]
            }

    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }


def handle_command(command: str) -> str:
    import re
    input_match = re.search(r'["\']([^"\']+\.pptx)["\']', command)
    output_match = re.search(r'output["\']?\s*[:=]\s*["\']([^"\']+)["\']', command)
    style_match = re.search(r'(格式|风格)\s+(F[1-4])', command)

    input_path = input_match.group(1) if input_match else "/tmp/input.pptx"
    output_path = output_match.group(1) if output_match else "/tmp/output.pptx"
    master_style = style_match.group(2) if style_match else "F1"
    include_header = "页眉" in command
    include_footer = "页脚" in command
    include_icon = "图标" in command

    result = convert_ppt(input_path, output_path, master_style, include_header, include_footer, include_icon)

    if result["success"]:
        return f"✅ PPT 转换成功！\n输出: {result['output_path']}\n页数: {result['pages']}\n校验: {result['validation_summary']}"
    else:
        issues = "\n".join(f"  - [{i['level']}] {i['rule_id']}: {i['message']}" for i in result.get("validation_issues", []))
        return f"❌ 转换失败: {result['error']}\n校验问题:\n{issues}"


if __name__ == "__main__":
    cmd = "把 input.pptx 转成公司标准模板，格式 F2（浅绿色清新），带页眉页脚，图标用 logo-A，输出到 output.pptx"
    print(handle_command(cmd))