"""Pre-flight environment validation for SAP Automation.

Checks Windows version, Python version, SAP GUI availability,
COM scripting, Office availability, required packages, and
generates a compatibility report before processing.

Usage:
    from sap_automation.core.precheck import PrecheckReport
    report = PrecheckReport.run_all()
    if not report.is_compatible:
        print(report.format())
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any

from .checks import (
    check_click,
    check_log_directory,
    check_office_available,
    check_openpyxl,
    check_output_directory,
    check_platform,
    check_python_version,
    check_pywin32,
    check_pyyaml,
    check_sap_gui,
    check_sap_scripting,
)
from .logger import get_logger

log = get_logger("precheck")


@dataclass
class CheckResult:
    """Result of a single environment check."""

    name: str
    passed: bool
    message: str
    details: str = ""
    suggestion: str = ""


@dataclass
class PrecheckReport:
    """Aggregated result of all environment checks."""

    checks: list[CheckResult] = field(default_factory=list)
    timestamp: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())

    @property
    def is_compatible(self) -> bool:
        return all(
            c.passed for c in self.checks
            if "critical" in c.name.lower()
        )

    @property
    def warnings(self) -> list[CheckResult]:
        return [c for c in self.checks if not c.passed]

    def format(self) -> str:
        lines = [
            "=" * 60,
            "ENVIRONMENT COMPATIBILITY REPORT",
            f"Generated: {self.timestamp}",
            "=" * 60,
        ]

        for check in self.checks:
            status = "PASS" if check.passed else "FAIL"
            lines.append(f"  [{status}] {check.name}")
            lines.append(f"         {check.message}")
            if check.details:
                lines.append(f"         Details: {check.details}")
            if check.suggestion and not check.passed:
                lines.append(f"         Fix: {check.suggestion}")
            lines.append("")

        lines.append("=" * 60)
        if self.is_compatible:
            lines.append("RESULT: Environment is compatible.")
        else:
            failed = [c for c in self.checks if not c.passed]
            lines.append(f"RESULT: {len(failed)} issue(s) found. Review above.")
        lines.append("=" * 60)
        return "\n".join(lines)

    def to_dict(self) -> dict[str, Any]:
        return {
            "timestamp": self.timestamp,
            "is_compatible": self.is_compatible,
            "checks": [
                {
                    "name": c.name,
                    "passed": c.passed,
                    "message": c.message,
                    "details": c.details,
                    "suggestion": c.suggestion,
                }
                for c in self.checks
            ],
        }


def _to_check_result(data: dict[str, Any]) -> CheckResult:
    return CheckResult(
        name=data["name"],
        passed=data["passed"],
        message=data["message"],
        details=data.get("details", ""),
        suggestion=data.get("suggestion", ""),
    )


def run_all(config_path: str | None = None) -> PrecheckReport:
    report = PrecheckReport()

    checks = [
        _to_check_result(check_python_version()),
        _to_check_result(check_platform()),
        _to_check_result(check_pywin32()),
        _to_check_result(check_openpyxl()),
        _to_check_result(check_click()),
        _to_check_result(check_pyyaml()),
        _to_check_result(check_sap_gui()),
        _to_check_result(check_sap_scripting()),
        _to_check_result(check_office_available()),
        _to_check_result(check_output_directory(config_path)),
        _to_check_result(check_log_directory(config_path)),
    ]

    report.checks = checks

    passed = sum(1 for c in checks if c.passed)
    total = len(checks)
    log.info("Precheck complete: %d/%d passed", passed, total)

    return report
