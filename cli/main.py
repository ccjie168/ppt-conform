import click
from pathlib import Path
from core.adapter.style_adapter import StyleAdapter
from core.validator.validator import Validator

@click.command()
@click.argument("input_path")
@click.argument("output_path")
@click.option("--master", "-m", default="F3", help="Master style: F1/F2/F3/F4")
@click.option("--footer", "-F", is_flag=True, help="Include footer")
@click.option("--icon", "-I", is_flag=True, help="Include icon")
def main(input_path: str, output_path: str, master: str, footer: bool, icon: bool):
    """PPT风格贴合转换器 - 不换模板，直接修改风格
    
    INPUT_PATH: 输入 PPT 文件路径
    
    OUTPUT_PATH: 输出 PPT 文件路径
    
    核心特点：
    1. 技术贴合：不换模板，直接在原PPT上修改风格
    2. 保留版式：完全保留原PPT的排版结构，避免版式错乱
    3. 风格统一：按模板配色调整背景、文字、强调色
    4. 清理水印：自动检测和删除水印
    5. 页脚处理：去除原页脚，添加模板页脚和图标
    """
    click.echo(f"转换 PPT: {input_path} -> {output_path}")
    click.echo(f"风格: {master}, 页脚: {footer}, 图标: {icon}")

    # 验证输入文件存在
    if not Path(input_path).exists():
        raise click.ClickException(f"输入文件不存在: {input_path}")

    try:
        # 1. 创建风格适配器
        adapter = StyleAdapter()
        
        # 2. 配置转换参数
        click.echo("1. 配置风格参数...")
        adapter.configure(master)
        
        # 3. 执行风格贴合转换
        click.echo("2. 执行风格贴合转换...")
        adapter.adapt(input_path, output_path)
        
        # 4. 质量校验
        validator = Validator()
        click.echo("3. 质量校验...")
        report = validator.validate(output_path)

        if report.passed:
            click.echo(f"✓ 校验通过: {report.summary}")
            click.echo(f"输出: {output_path}")
        else:
            click.echo("⚠ 校验存在问题:")
            for issue in report.issues:
                slide_info = f" (页面 {issue.slide_index})" if issue.slide_index is not None else ""
                click.echo(f"  - [{issue.level}] {issue.rule_id}: {issue.message}{slide_info}")
            
            click.echo(f"输出: {output_path}")
            click.echo("提示: 请检查输出文件中的问题，部分校验项可能需要人工确认")

    except Exception as e:
        click.echo(f"错误: {e}", err=True)
        raise


if __name__ == "__main__":
    main()