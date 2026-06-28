"""Enterprise diagnostic command for SAP Automation Framework.

Performs comprehensive read-only health checks of the local environment.
Never modifies SAP, never executes transactions, never changes configuration.

Usage:
    sap-auto doctor
    sap-auto doctor --json
    sap-auto doctor --verbose
    sap-auto doctor --save
"""

from __future__ import annotations

import contextlib
import json
import os
import platform
import re
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .checks import (
    check_architecture,
    check_click,
    check_disk_space,
    check_env_vars,
    check_openpyxl,
    check_platform,
    check_python_version,
    check_pywin32,
    check_pyyaml,
    check_sap_gui,
    check_sap_scripting,
    check_secrets,
)
from .config import Config, ConfigError
from .logger import get_logger

log = get_logger("doctor")


@dataclass
class DiagnosticCheck:
    """Result of a single diagnostic check."""

    name: str
    status: str  # "pass", "warning", "fail"
    message: str
    details: str = ""
    suggestion: str = ""
    score_weight: int = 1
    scored: bool = True

    @property
    def icon(self) -> str:
        return {"pass": "\u2713", "warning": "\u26a0", "fail": "\u2717"}.get(self.status, "?")

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "name": self.name,
            "status": self.status,
            "message": self.message,
        }
        if self.details:
            d["details"] = self.details
        if self.suggestion:
            d["suggestion"] = self.suggestion
        return d


@dataclass
class DoctorReport:
    """Full diagnostic report."""

    timestamp: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    framework_version: str = ""
    checks: list[DiagnosticCheck] = field(default_factory=list)
    environment: dict[str, str] = field(default_factory=dict)
    _categories: dict[str, list[str]] = field(default_factory=dict)

    @property
    def health_score(self) -> tuple[int, int]:
        earned = 0
        total = 0
        for c in self.checks:
            if not c.scored or c.score_weight == 0:
                continue
            total += c.score_weight
            if c.status == "pass":
                earned += c.score_weight
            elif c.status == "warning":
                earned += c.score_weight // 2
        return earned, total

    @property
    def health_pct(self) -> int:
        earned, total = self.health_score
        return int((earned / total * 100) if total else 0)

    @property
    def overall_status(self) -> str:
        if any(c.status == "fail" for c in self.checks):
            return "NOT READY"
        if any(c.status == "warning" for c in self.checks):
            return "READY WITH WARNINGS"
        return "READY"

    @property
    def is_ready(self) -> bool:
        return self.overall_status == "READY"

    def add(self, check: DiagnosticCheck, category: str = "General") -> None:
        self.checks.append(check)
        self._categories.setdefault(category, []).append(check.name)

    def get_category(self, name: str) -> list[DiagnosticCheck]:
        names = self._categories.get(name, [])
        lookup = {c.name: c for c in self.checks}
        return [lookup[n] for n in names if n in lookup]

    def failed_checks(self) -> list[DiagnosticCheck]:
        return [c for c in self.checks if c.status == "fail"]

    def warning_checks(self) -> list[DiagnosticCheck]:
        return [c for c in self.checks if c.status == "warning"]

    def to_dict(self) -> dict[str, Any]:
        earned, total = self.health_score
        return {
            "timestamp": self.timestamp,
            "framework_version": self.framework_version,
            "health_score": f"{earned}/{total}",
            "health_pct": self.health_pct,
            "overall_status": self.overall_status,
            "environment": self.environment,
            "checks": [c.to_dict() for c in self.checks],
            "categories": {
                cat: [
                    next(c.to_dict() for c in self.checks if c.name == name)
                    for name in names
                    if any(c.name == name for c in self.checks)
                ]
                for cat, names in self._categories.items()
            },
        }

    def format_console(self, verbose: bool = False) -> str:
        lines = [
            "=" * 60,
            "",
            "  SAP Automation Framework \u2014 Diagnostic Report",
            "",
            "=" * 60,
            "",
            f"  Framework Version ......... {self.framework_version}",
        ]

        env = self.environment
        lines.append(f"  Python .................... {env.get('python_version', '?')}")
        lines.append(f"  Platform .................. {env.get('os', '?')} {env.get('os_release', '')}")
        lines.append(f"  Architecture .............. {env.get('machine', '?')}")
        lines.append(f"  Working Directory ......... {env.get('cwd', '?')}")
        lines.append("")

        for cat, names in self._categories.items():
            lines.append(f"  {cat}")
            lines.append("  " + "-" * 56)
            for name in names:
                c = next((x for x in self.checks if x.name == name), None)
                if c is None:
                    continue
                status_str = f"[{c.icon}]"
                msg = c.message
                if verbose and c.details:
                    msg += f" ({c.details})"
                lines.append(f"  {status_str} {c.name:<28s} {msg}")
                if verbose and c.suggestion:
                    lines.append(f"      Fix: {c.suggestion}")
            lines.append("")

        earned, total = self.health_score
        lines.append("=" * 60)
        lines.append(f"  Health Score: {earned}/{total} ({self.health_pct}%)")
        lines.append("")

        status = self.overall_status
        if status == "READY":
            lines.append("  Status: READY FOR EXECUTION")
        elif status == "READY WITH WARNINGS":
            lines.append("  Status: READY WITH WARNINGS")
            for w in self.warning_checks():
                lines.append(f"    \u26a0 {w.name}: {w.message}")
        else:
            lines.append("  Status: NOT READY")
            for f in self.failed_checks():
                lines.append(f"    \u2717 {f.name}: {f.message}")
                if f.suggestion:
                    lines.append(f"      \u2192 {f.suggestion}")

        lines.append("")
        lines.append("=" * 60)
        return "\n".join(lines)


def _to_diagnostic(data: dict[str, Any], score_weight: int = 1, scored: bool = True) -> DiagnosticCheck:
    passed = data.get("passed", False)
    return DiagnosticCheck(
        name=data["name"],
        status="pass" if passed else "fail",
        message=data["message"],
        details=data.get("details", ""),
        suggestion=data.get("suggestion", ""),
        score_weight=score_weight,
        scored=scored,
    )


def _check_config(config_path: str = "config/default.yaml") -> DiagnosticCheck:
    path = Path(config_path)
    if not path.exists():
        return DiagnosticCheck(
            name="Configuration",
            status="fail",
            message=f"Config not found: {path}",
            suggestion="Run: sap-auto init  or create config/default.yaml",
            score_weight=5,
        )

    try:
        config = Config.from_file(str(path))
    except ConfigError as exc:
        return DiagnosticCheck(
            name="Configuration",
            status="fail",
            message=f"Invalid config: {exc}",
            suggestion="Fix the YAML syntax or missing required keys",
            score_weight=5,
        )

    try:
        config.validate()
    except ConfigError as exc:
        return DiagnosticCheck(
            name="Configuration",
            status="fail",
            message=f"Validation failed: {exc}",
            suggestion="Check config/default.yaml for missing or invalid values",
            score_weight=5,
        )

    raw = path.read_text(encoding="utf-8")
    unresolved = re.findall(r"\$\{(\w+)\}", raw)
    if unresolved:
        missing = [v for v in unresolved if not os.environ.get(v)]
        if missing:
            return DiagnosticCheck(
                name="Configuration",
                status="warning",
                message=f"Unresolved env vars: {', '.join(set(missing))}",
                suggestion="Set environment variables or replace with literal values",
                score_weight=5,
            )

    return DiagnosticCheck(
        name="Configuration",
        status="pass",
        message="OK (valid)",
        score_weight=5,
    )


def _check_transactions(config_path: str = "config/default.yaml") -> DiagnosticCheck:
    from importlib import import_module

    path = Path(config_path)
    if not path.exists():
        return DiagnosticCheck(
            name="Transactions",
            status="fail",
            message="Config not found",
            score_weight=3,
        )

    try:
        config = Config.from_file(str(path))
        config.validate()
    except ConfigError:
        return DiagnosticCheck(
            name="Transactions",
            status="fail",
            message="Cannot validate (config invalid)",
            score_weight=3,
        )

    transactions = config.get("transactions", {})
    if not transactions:
        return DiagnosticCheck(
            name="Transactions",
            status="warning",
            message="No transactions registered",
            suggestion="Add transaction entries in config/default.yaml",
            score_weight=3,
        )

    loaded = 0
    failed_names: list[str] = []
    registered: list[str] = []

    for name, txn in transactions.items():
        class_path = txn.get("class", "")
        if not class_path:
            failed_names.append(f"{name} (no class)")
            continue

        try:
            if not class_path.startswith("sap_automation.transactions."):
                failed_names.append(f"{name} (bad namespace)")
                continue
            module_path, class_name = class_path.rsplit(".", 1)
            mod = import_module(module_path)
            cls = getattr(mod, class_name)
            from .base_transaction import BaseTransaction
            if isinstance(cls, type) and issubclass(cls, BaseTransaction):
                loaded += 1
                registered.append(name.upper())
            else:
                failed_names.append(f"{name} (not BaseTransaction)")
        except Exception as exc:
            failed_names.append(f"{name} ({exc})")

    if failed_names:
        return DiagnosticCheck(
            name="Transactions",
            status="warning",
            message=f"{loaded} OK, {len(failed_names)} failed",
            details=f"Failed: {', '.join(failed_names)}",
            suggestion="Fix class paths in config/default.yaml",
            score_weight=3,
        )

    icons = " ".join(f"\u2713 {n}" for n in registered)
    return DiagnosticCheck(
        name="Transactions",
        status="pass",
        message=f"{loaded} registered ({', '.join(registered)})",
        details=icons,
        score_weight=3,
    )


def _check_output_dir() -> DiagnosticCheck:
    out = Path("./output")
    if not out.exists():
        with contextlib.suppress(OSError):
            out.mkdir(parents=True, exist_ok=True)

    if not out.exists():
        return DiagnosticCheck(
            name="Output Directory",
            status="fail",
            message="NOT FOUND",
            suggestion="Run: mkdir output",
            score_weight=2,
        )

    if not os.access(out, os.W_OK):
        return DiagnosticCheck(
            name="Output Directory",
            status="fail",
            message="NOT WRITABLE",
            suggestion="Check directory permissions",
            score_weight=2,
        )

    file_count = sum(1 for _ in out.iterdir()) if out.is_dir() else 0
    return DiagnosticCheck(
        name="Output Directory",
        status="pass",
        message=f"OK ({out.resolve()}, {file_count} file(s))",
        score_weight=2,
    )


def _check_log_dir() -> DiagnosticCheck:
    logd = Path("./logs")
    if not logd.exists():
        with contextlib.suppress(OSError):
            logd.mkdir(parents=True, exist_ok=True)

    if not logd.exists():
        return DiagnosticCheck(
            name="Log Directory",
            status="fail",
            message="NOT FOUND",
            suggestion="Run: mkdir logs",
            score_weight=2,
        )

    if not os.access(logd, os.W_OK):
        return DiagnosticCheck(
            name="Log Directory",
            status="fail",
            message="NOT WRITABLE",
            suggestion="Check directory permissions",
            score_weight=2,
        )

    return DiagnosticCheck(
        name="Log Directory",
        status="pass",
        message=f"OK ({logd.resolve()})",
        score_weight=2,
    )


def _check_temp_dir() -> DiagnosticCheck:
    tmp = Path(tempfile.gettempdir())
    ok = tmp.exists() and os.access(tmp, os.W_OK)
    return DiagnosticCheck(
        name="Temp Directory",
        status="pass" if ok else "warning",
        message=f"{'OK' if ok else 'NOT WRITABLE'} ({tmp})",
        score_weight=1,
    )


def _check_excel_com() -> DiagnosticCheck:
    if sys.platform != "win32":
        return DiagnosticCheck(
            name="Excel COM",
            status="warning",
            message="Cannot check (not Windows)",
            score_weight=0,
            scored=False,
        )

    excel = None
    try:
        import win32com.client  # type: ignore[import-untyped]
        excel = win32com.client.Dispatch("Excel.Application")
        version = excel.Version
        return DiagnosticCheck(
            name="Excel COM",
            status="pass",
            message=f"OK (Excel {version})",
            score_weight=3,
        )
    except Exception as exc:
        return DiagnosticCheck(
            name="Excel COM",
            status="warning",
            message=f"Not available: {exc}",
            details="PDF export requires Excel via COM",
            suggestion="Install Microsoft Office for PDF export support (optional)",
            score_weight=3,
        )
    finally:
        if excel:
            with contextlib.suppress(Exception):
                excel.Quit()
                del excel


def _check_log_config() -> DiagnosticCheck:
    logd = Path("./logs")
    if not logd.exists():
        return DiagnosticCheck(
            name="Logging",
            status="warning",
            message="Log directory missing",
            suggestion="Run: mkdir logs",
            score_weight=1,
        )

    log_files = list(logd.glob("*.log"))
    if log_files:
        latest = max(log_files, key=lambda p: p.stat().st_mtime)
        size_kb = latest.stat().st_size / 1024
        return DiagnosticCheck(
            name="Logging",
            status="pass",
            message=f"OK ({len(log_files)} log file(s), latest: {latest.name}, {size_kb:.1f} KB)",
            score_weight=1,
        )

    return DiagnosticCheck(
        name="Logging",
        status="pass",
        message="OK (log directory ready, no logs yet)",
        score_weight=1,
    )


def _check_output_paths() -> DiagnosticCheck:
    dangerous = ["/usr", "/etc", "/bin", "/sbin", "C:\\Windows", "C:\\Program Files"]
    cwd = os.getcwd()

    for d in dangerous:
        if cwd.startswith(d):
            return DiagnosticCheck(
                name="Output Paths",
                status="fail",
                message=f"Working directory is a system path: {cwd}",
                suggestion="Run from a user directory, not a system directory",
                score_weight=2,
            )

    return DiagnosticCheck(
        name="Output Paths",
        status="pass",
        message="OK (safe working directory)",
        score_weight=2,
    )


def run_all(config_path: str = "config/default.yaml") -> DoctorReport:
    from sap_automation import __version__

    report = DoctorReport(framework_version=__version__)

    report.environment = {
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        "os": platform.system(),
        "os_release": platform.release(),
        "machine": platform.machine(),
        "cwd": os.getcwd(),
    }

    report.add(_to_diagnostic(check_python_version(), score_weight=10), "System")
    report.add(_to_diagnostic(check_platform(), score_weight=10), "System")
    report.add(_to_diagnostic(check_architecture(), score_weight=2), "System")
    report.add(DiagnosticCheck(
        name="Working Directory",
        status="pass" if os.path.isdir(os.getcwd()) and os.access(os.getcwd(), os.R_OK) else "fail",
        message=f"OK ({os.getcwd()})" if os.path.isdir(os.getcwd()) else "FAIL",
        score_weight=2,
    ), "System")

    report.add(_to_diagnostic(check_pywin32(), score_weight=3), "Python Packages")
    report.add(_to_diagnostic(check_openpyxl(), score_weight=3), "Python Packages")
    report.add(_to_diagnostic(check_click(), score_weight=3), "Python Packages")
    report.add(_to_diagnostic(check_pyyaml(), score_weight=3), "Python Packages")

    sap_gui = check_sap_gui()
    sap_scripting = check_sap_scripting()
    report.add(_to_diagnostic(sap_gui, score_weight=5), "SAP")
    report.add(_to_diagnostic(sap_scripting, score_weight=5), "SAP")

    report.add(_check_config(config_path), "Configuration")
    report.add(_check_transactions(config_path), "Configuration")

    report.add(_check_output_dir(), "File System")
    report.add(_check_log_dir(), "File System")
    report.add(_to_diagnostic(check_disk_space(), score_weight=2), "File System")
    report.add(_check_temp_dir(), "File System")

    report.add(_check_excel_com(), "Office")

    report.add(_check_log_config(), "Logging")

    secrets = check_secrets(config_path)
    env_vars = check_env_vars(config_path)
    report.add(_to_diagnostic(secrets, score_weight=2), "Security")
    report.add(_to_diagnostic(env_vars, score_weight=2), "Security")
    report.add(_check_output_paths(), "Security")

    return report


def save_report(report: DoctorReport, output_dir: str = "diagnostics") -> tuple[Path, Path]:
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    json_path = out / "report.json"
    txt_path = out / "report.txt"

    json_path.write_text(
        json.dumps(report.to_dict(), indent=2, default=str),
        encoding="utf-8",
    )

    txt_path.write_text(
        report.format_console(verbose=True),
        encoding="utf-8",
    )

    return json_path, txt_path
