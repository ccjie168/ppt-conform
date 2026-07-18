from core.models import ValidationReport, ValidationIssue
from core.validator.rules import WatermarkTextRule, FontWhitelistRule, ContentOverflowRule


class Validator:
    def __init__(self):
        self.rules = [
            WatermarkTextRule(),
            FontWhitelistRule(),
            ContentOverflowRule(),
        ]

    def validate(self, pptx_path: str) -> ValidationReport:
        issues = []

        for rule in self.rules:
            rule_issues = rule.check(pptx_path)
            issues.extend(rule_issues)

        has_fail = any(issue.level == "fail" for issue in issues)
        passed = not has_fail

        summary_parts = []
        if has_fail:
            fail_count = sum(1 for i in issues if i.level == "fail")
            summary_parts.append(f"{fail_count} 项失败")
        warn_count = sum(1 for i in issues if i.level == "warning")
        if warn_count > 0:
            summary_parts.append(f"{warn_count} 项警告")
        if passed:
            summary_parts.append("校验通过")

        return ValidationReport(
            passed=passed,
            issues=issues,
            summary="; ".join(summary_parts) if summary_parts else "无问题"
        )