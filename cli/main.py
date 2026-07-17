import click
from pathlib import Path
from core.extractor.pptx_extractor import PptxExtractor
from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.validator.validator import Validator
from core.models import UserConfig


@click.command()
@click.argument("input_path")
@click.argument("output_path")
@click.option("--master", "-m", default="F1", help="Master style: F1/F2/F3/F4")
@click.option("--header", "-H", is_flag=True, help="Include header")
@click.option("--footer", "-F", is_flag=True, help="Include footer")
@click.option("--icon", "-I", is_flag=True, help="Include icon")
def main(input_path: str, output_path: str, master: str, header: bool, footer: bool, icon: bool):
    """PPT 标准模板转换工具
    
    INPUT_PATH: 输入 PPT 文件路径
    
    OUTPUT_PATH: 输出 PPT 文件路径
    """
    click.echo(f"转换 PPT: {input_path} -> {output_path}")
    click.echo(f"风格: {master}, 页眉: {header}, 页脚: {footer}, 图标: {icon}")

    # 验证输入文件存在
    if not Path(input_path).exists():
        raise click.ClickException(f"输入文件不存在: {input_path}")
    
    # 创建用户配置
    config = UserConfig(
        input_path=input_path,
        output_path=output_path,
        master_style=master,
        include_header=header,
        include_footer=footer,
        include_icon=icon
    )

    try:
        # 1. 抽取内容
        extractor = PptxExtractor()
        click.echo("1. 抽取内容...")
        content_models = extractor.extract(input_path)
        click.echo(f"   共 {len(content_models)} 页")

        # 2. 加载模板
        registry = TemplateRegistry()
        registry.load_master_styles()
        registry.load_layout_mappings()
        click.echo("2. 加载模板...")

        # 3. 重放内容
        replayer = ContentReplayer(registry)
        click.echo("3. 重放内容...")
        temp_output = replayer.replay(content_models, config)

        # 4. 质量校验
        validator = Validator()
        click.echo("4. 质量校验...")
        report = validator.validate(temp_output)

        if report.passed:
            click.echo(f"✓ 校验通过: {report.summary}")
            click.echo(f"输出: {output_path}")
        else:
            click.echo("✗ 校验失败:")
            for issue in report.issues:
                slide_info = f" (页面 {issue.slide_index})" if issue.slide_index is not None else ""
                click.echo(f"  - [{issue.level}] {issue.rule_id}: {issue.message}{slide_info}")
            
            # 删除校验失败的文件
            import os
            if os.path.exists(temp_output):
                os.remove(temp_output)
            
            raise click.ClickException("校验失败，已阻断输出")

    except Exception as e:
        click.echo(f"错误: {e}", err=True)
        raise


if __name__ == "__main__":
    main()