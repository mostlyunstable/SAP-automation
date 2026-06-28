"""Enterprise compatibility testing for SAP Automation Framework.

Performs deep read-only validation of the machine and SAP environment.
Never modifies SAP data, never executes business transactions.

Usage:
    sap-auto compatibility-test
    sap-auto compatibility-test --json
    sap-auto compatibility-test --verbose
    sap-auto compatibility-test --save
"""

from __future__ import annotations

import contextlib
import json
import os
import platform
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .checks import (
    check_architecture,
    check_disk_space,
    check_office_available,
    check_platform,
    check_python_version,
)
from .config import Config, ConfigError
from .logger import get_logger

log = get_logger("compatibility")


@dataclass
class CompatCheck:
    """Result of a single compatibility check."""

    name: str
    status: str  # "pass", "warning", "fail"
    message: str
    details: str = ""
    suggestion: str = ""
    category: str = "General"
    critical: bool = False

    @property
    def icon(self) -> str:
        return {"pass": "\u2713", "warning": "\u26a0", "fail": "\u2717"}.get(self.status, "?")

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "name": self.name,
            "status": self.status,
            "message": self.message,
            "category": self.category,
            "critical": self.critical,
        }
        if self.details:
            d["details"] = self.details
        if self.suggestion:
            d["suggestion"] = self.suggestion
        return d


@dataclass
class TimingResult:
    """Performance timing measurement."""

    name: str
    value_ms: float
    threshold_ms: float
    status: str = "pass"

    @property
    def message(self) -> str:
        if self.value_ms > self.threshold_ms:
            return f"SLOW ({self.value_ms:.0f}ms, threshold: {self.threshold_ms:.0f}ms)"
        return f"OK ({self.value_ms:.0f}ms)"

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "value_ms": round(self.value_ms, 1),
            "threshold_ms": self.threshold_ms,
            "status": self.status,
        }


@dataclass
class CompatReport:
    """Full compatibility report."""

    timestamp: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    framework_version: str = ""
    checks: list[CompatCheck] = field(default_factory=list)
    timings: list[TimingResult] = field(default_factory=list)
    environment: dict[str, str] = field(default_factory=dict)
    sap_info: dict[str, str] = field(default_factory=dict)
    recommendations: list[str] = field(default_factory=list)

    @property
    def score(self) -> tuple[int, int]:
        weights = {
            "Environment": 15, "SAP": 20, "Session": 15,
            "Screen": 10, "Fields": 15, "Export": 10,
            "Permissions": 5, "Timing": 5, "Theme": 5,
        }
        cat_status: dict[str, list[str]] = {}
        for c in self.checks:
            cat_status.setdefault(c.category, []).append(c.status)

        earned = 0
        total = 0
        for cat, weight in weights.items():
            statuses = cat_status.get(cat, [])
            if not statuses:
                continue
            total += weight
            if all(s == "pass" for s in statuses):
                earned += weight
            elif any(s == "fail" for s in statuses):
                pass_count = sum(1 for s in statuses if s == "pass")
                earned += max(0, int(weight * pass_count / len(statuses)))
            else:
                earned += weight // 2

        return earned, total

    @property
    def score_pct(self) -> int:
        earned, total = self.score
        return int((earned / total * 100) if total else 0)

    @property
    def risk_level(self) -> str:
        critical_fails = sum(1 for c in self.checks if c.status == "fail" and c.critical)
        all_fails = sum(1 for c in self.checks if c.status == "fail")
        all_warnings = sum(1 for c in self.checks if c.status == "warning")

        if critical_fails > 0:
            return "CRITICAL"
        if all_fails > 2:
            return "HIGH"
        if all_fails > 0 or all_warnings > 3:
            return "MEDIUM"
        return "LOW"

    @property
    def compatibility_status(self) -> str:
        if any(c.status == "fail" and c.critical for c in self.checks):
            return "NOT COMPATIBLE"
        if any(c.status == "fail" for c in self.checks):
            return "NOT COMPATIBLE"
        if any(c.status == "warning" for c in self.checks):
            return "COMPATIBLE WITH WARNINGS"
        return "COMPATIBLE"

    @property
    def exit_code(self) -> int:
        if any(c.status == "fail" for c in self.checks):
            return 2
        if any(c.status == "warning" for c in self.checks):
            return 1
        return 0

    def add(self, check: CompatCheck) -> None:
        self.checks.append(check)

    def add_timing(self, timing: TimingResult) -> None:
        self.timings.append(timing)

    def add_recommendation(self, text: str) -> None:
        self.recommendations.append(text)

    def failed_checks(self) -> list[CompatCheck]:
        return [c for c in self.checks if c.status == "fail"]

    def warning_checks(self) -> list[CompatCheck]:
        return [c for c in self.checks if c.status == "warning"]

    def to_dict(self) -> dict[str, Any]:
        earned, total = self.score
        return {
            "timestamp": self.timestamp,
            "framework_version": self.framework_version,
            "compatibility_status": self.compatibility_status,
            "risk_level": self.risk_level,
            "score": f"{earned}/{total}",
            "score_pct": self.score_pct,
            "exit_code": self.exit_code,
            "environment": self.environment,
            "sap_info": self.sap_info,
            "checks": [c.to_dict() for c in self.checks],
            "timings": [t.to_dict() for t in self.timings],
            "recommendations": self.recommendations,
        }

    def format_console(self, verbose: bool = False) -> str:
        lines = [
            "=" * 60,
            "",
            "  SAP Automation \u2014 Compatibility Test Report",
            "",
            "=" * 60,
            "",
            f"  Framework Version ......... {self.framework_version}",
        ]

        env = self.environment
        lines.append(f"  Python .................... {env.get('python_version', '?')}")
        lines.append(f"  Platform .................. {env.get('os', '?')} {env.get('os_release', '')}")
        lines.append(f"  Architecture .............. {env.get('machine', '?')}")
        lines.append("")

        categories: dict[str, list[CompatCheck]] = {}
        for c in self.checks:
            categories.setdefault(c.category, []).append(c)

        for cat_name, checks in categories.items():
            lines.append(f"  {cat_name}")
            lines.append("  " + "-" * 56)
            for c in checks:
                status_str = f"[{c.icon}]"
                msg = c.message
                if verbose and c.details:
                    msg += f" ({c.details})"
                lines.append(f"  {status_str} {c.name:<28s} {msg}")
                if verbose and c.suggestion:
                    lines.append(f"      Fix: {c.suggestion}")
            lines.append("")

        if self.sap_info:
            lines.append("  SAP Session Details")
            lines.append("  " + "-" * 56)
            for k, v in self.sap_info.items():
                lines.append(f"    {k:<24s} {v}")
            lines.append("")

        if self.timings:
            lines.append("  Performance")
            lines.append("  " + "-" * 56)
            for t in self.timings:
                icon = "\u2713" if t.status == "pass" else "\u26a0"
                lines.append(f"    [{icon}] {t.name:<24s} {t.message}")
            lines.append("")

        earned, total = self.score
        lines.append("=" * 60)
        lines.append(f"  Compatibility Score: {earned}/{total} ({self.score_pct}%)")
        lines.append(f"  Risk Level: {self.risk_level}")
        lines.append("")

        status = self.compatibility_status
        if status == "COMPATIBLE":
            lines.append("  Status: \u2713 COMPATIBLE")
            lines.append("  No blocking issues detected.")
        elif status == "COMPATIBLE WITH WARNINGS":
            lines.append("  Status: \u26a0 COMPATIBLE WITH WARNINGS")
            for w in self.warning_checks():
                lines.append(f"    \u26a0 {w.name}: {w.message}")
        else:
            lines.append("  Status: \u2717 NOT COMPATIBLE")
            for f in self.failed_checks():
                lines.append(f"    \u2717 {f.name}: {f.message}")
                if f.suggestion:
                    lines.append(f"      \u2192 {f.suggestion}")

        if self.recommendations:
            lines.append("")
            lines.append("  Recommendations:")
            for r in self.recommendations:
                lines.append(f"    \u2022 {r}")

        lines.append("")
        lines.append("=" * 60)
        return "\n".join(lines)


def _to_compat(data: dict[str, Any], category: str = "General", critical: bool = False) -> CompatCheck:
    passed = data.get("passed", False)
    return CompatCheck(
        name=data["name"],
        status="pass" if passed else "fail",
        message=data["message"],
        details=data.get("details", ""),
        suggestion=data.get("suggestion", ""),
        category=category,
        critical=critical,
    )


def _check_environment() -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    plat = check_platform()
    checks.append(CompatCheck(
        name=plat["name"],
        status="pass" if plat["passed"] else "fail",
        message=f"{'OK' if plat['passed'] else 'FAIL'} \u2014 {plat['message']}",
        category="Environment",
        critical=not plat["passed"],
        suggestion=plat.get("suggestion", ""),
    ))

    pyver = check_python_version()
    checks.append(CompatCheck(
        name=pyver["name"],
        status="pass" if pyver["passed"] else "fail",
        message=f"{'OK' if pyver['passed'] else 'FAIL'} \u2014 {pyver['message']}",
        details=pyver.get("details", ""),
        category="Environment",
        critical=not pyver["passed"],
        suggestion=pyver.get("suggestion", ""),
    ))

    arch = check_architecture()
    checks.append(CompatCheck(
        name=arch["name"],
        status="pass" if arch["passed"] else "warning",
        message=f"{'OK' if arch['passed'] else 'WARNING'} \u2014 {arch['message']}",
        category="Environment",
    ))

    writable = 0
    for p in [Path("."), Path.home()]:
        with contextlib.suppress(OSError):
            if os.access(p, os.W_OK):
                writable += 1
    checks.append(CompatCheck(
        name="User Permissions",
        status="pass" if writable > 0 else "warning",
        message=f"{'OK' if writable > 0 else 'WARNING'} \u2014 {writable} writable location(s)",
        category="Environment",
    ))

    disk = check_disk_space()
    disk_status = "pass" if disk["passed"] else ("warning" if "LOW" in disk.get("message", "") else "fail")
    checks.append(CompatCheck(
        name=disk["name"],
        status=disk_status,
        message=f"{'OK' if disk_status == 'pass' else disk_status.upper()} \u2014 {disk['message']}",
        category="Environment",
        critical=disk_status == "fail",
        suggestion=disk.get("suggestion", ""),
    ))

    return checks


def _check_sap_environment() -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    if sys.platform != "win32":
        checks.append(CompatCheck(
            name="SAP GUI",
            status="warning",
            message="Cannot check (not Windows)",
            category="SAP",
        ))
        checks.append(CompatCheck(
            name="SAP Scripting",
            status="warning",
            message="Cannot check (not Windows)",
            category="SAP",
        ))
        return checks

    try:
        import win32com.client  # type: ignore[import-untyped]
        gui = win32com.client.GetObject("SAPGUI")
        version = str(getattr(gui, "Version", "unknown"))
        checks.append(CompatCheck(
            name="SAP GUI",
            status="pass",
            message=f"OK \u2014 version {version}",
            category="SAP",
            critical=True,
        ))
    except Exception as exc:
        checks.append(CompatCheck(
            name="SAP GUI",
            status="fail",
            message="NOT DETECTED",
            details=str(exc),
            category="SAP",
            critical=True,
            suggestion="Ensure SAP Logon is running",
        ))
        return checks

    try:
        app = gui.GetScriptingEngine
        count = app.Children.Count
        checks.append(CompatCheck(
            name="SAP Scripting",
            status="pass",
            message=f"ENABLED \u2014 {count} connection(s)",
            category="SAP",
            critical=True,
        ))
    except Exception as exc:
        checks.append(CompatCheck(
            name="SAP Scripting",
            status="fail",
            message="DISABLED or NOT ACCESSIBLE",
            details=str(exc),
            category="SAP",
            critical=True,
            suggestion="Enable: SAP Logon > Options > Accessibility > GUI Scripting",
        ))

    return checks


def _check_session(config: Config) -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    if sys.platform != "win32":
        checks.append(CompatCheck(
            name="Session Connectivity",
            status="warning",
            message="Cannot check (not Windows)",
            category="Session",
        ))
        return checks

    try:
        from .connection import SAPConnection

        system_name = config.get("sap.system_name")
        client = config.get("sap.client")

        start = time.monotonic()
        with SAPConnection(
            system_name=system_name,
            client=client,
            session_timeout=10.0,
        ) as sap:
            elapsed_ms = (time.monotonic() - start) * 1000

            checks.append(CompatCheck(
                name="Session Alive",
                status="pass",
                message="OK",
                category="Session",
                critical=True,
            ))

            info = sap.get_session_info()
            checks.append(CompatCheck(
                name="Connection Match",
                status="pass",
                message=f"OK \u2014 system={info.get('system', '?')}, client={info.get('client', '?')}",
                category="Session",
            ))

            is_valid = sap.validate_session()
            checks.append(CompatCheck(
                name="Session Responsive",
                status="pass" if is_valid else "fail",
                message="OK" if is_valid else "NOT RESPONSIVE",
                category="Session",
                critical=not is_valid,
            ))

            busy_str = info.get("busy", "unknown")
            is_busy = busy_str == "True"
            checks.append(CompatCheck(
                name="Session Busy State",
                status="warning" if is_busy else "pass",
                message="BUSY \u2014 may be processing" if is_busy else "IDLE",
                category="Session",
            ))

            try:
                conn_count = sap.connection.Sessions.Count
                checks.append(CompatCheck(
                    name="Multiple Sessions",
                    status="pass",
                    message=f"{conn_count} session(s) on this connection",
                    category="Session",
                ))
            except Exception:  # noqa: S110 — best-effort COM check
                pass

            checks.append(CompatCheck(
                name="Connection Time",
                status="pass" if elapsed_ms < 5000 else "warning",
                message=f"{'OK' if elapsed_ms < 5000 else 'SLOW'} \u2014 {elapsed_ms:.0f}ms",
                category="Timing",
            ))

    except Exception as exc:
        checks.append(CompatCheck(
            name="Session Connectivity",
            status="fail",
            message=f"FAILED \u2014 {exc}",
            category="Session",
            critical=True,
            suggestion="Ensure SAP Logon is open and you are logged in",
        ))

    return checks


def _check_screen(config: Config) -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    if sys.platform != "win32":
        checks.append(CompatCheck(
            name="Screen Validation",
            status="warning",
            message="Cannot check (not Windows)",
            category="Screen",
        ))
        return checks

    transactions = config.get("transactions", {})
    if not transactions:
        checks.append(CompatCheck(
            name="Screen Validation",
            status="warning",
            message="No transactions registered to test",
            category="Screen",
        ))
        return checks

    test_tcode = next(iter(transactions.keys())).upper()

    try:
        from .connection import SAPConnection

        system_name = config.get("sap.system_name")
        client = config.get("sap.client")

        with SAPConnection(
            system_name=system_name,
            client=client,
            session_timeout=10.0,
        ) as sap:
            start = time.monotonic()
            sap.open_transaction(test_tcode)
            elapsed_ms = (time.monotonic() - start) * 1000

            checks.append(CompatCheck(
                name="Transaction Opens",
                status="pass",
                message=f"OK \u2014 /n{test_tcode} opened",
                category="Screen",
                critical=True,
            ))

            checks.append(CompatCheck(
                name="Screen Load Time",
                status="pass" if elapsed_ms < 10000 else "warning",
                message=f"{'OK' if elapsed_ms < 10000 else 'SLOW'} \u2014 {elapsed_ms:.0f}ms",
                category="Timing",
            ))

            try:
                window = sap.session.findById("wnd[0]")
                title = str(window.Text)
                checks.append(CompatCheck(
                    name="Window Title",
                    status="pass",
                    message=f"OK \u2014 '{title}'",
                    category="Screen",
                ))
            except Exception:
                checks.append(CompatCheck(
                    name="Window Title",
                    status="warning",
                    message="Cannot read window title",
                    category="Screen",
                ))

            try:
                sbar = sap.session.findById("wnd[0]/sbar")
                msg_type = str(sbar.MessageType)
                checks.append(CompatCheck(
                    name="Status Bar",
                    status="pass",
                    message=f"OK \u2014 accessible (type: {msg_type})",
                    category="Screen",
                ))
            except Exception:
                checks.append(CompatCheck(
                    name="Status Bar",
                    status="warning",
                    message="Cannot access status bar",
                    category="Screen",
                ))

            try:
                sap.session.findById("wnd[0]")
                checks.append(CompatCheck(
                    name="Main Window",
                    status="pass",
                    message="OK \u2014 active",
                    category="Screen",
                ))
            except Exception:
                checks.append(CompatCheck(
                    name="Main Window",
                    status="warning",
                    message="Cannot verify main window",
                    category="Screen",
                ))

            with contextlib.suppress(Exception):
                sap.session.findById("wnd[0]").sendVKey(3)

    except Exception as exc:
        checks.append(CompatCheck(
            name="Transaction Opens",
            status="fail",
            message=f"FAILED \u2014 {exc}",
            category="Screen",
            critical=True,
            suggestion=f"Check transaction /n{test_tcode} authorization and screen layout",
        ))

    return checks


def _check_fields(config: Config) -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    if sys.platform != "win32":
        checks.append(CompatCheck(
            name="Field Validation",
            status="warning",
            message="Cannot check (not Windows)",
            category="Fields",
        ))
        return checks

    field_ids = config.get("transactions.va23.field_ids", {})
    if not field_ids:
        field_ids = {
            "input_vbeln": "wnd[0]/usr/ctxtRV45A-VBELN",
            "sold_to": "wnd[0]/usr/subHEADER/SUB1/RBHP-VERTR",
            "document_date": "wnd[0]/usr/subHEADER/SUB1/RBHP-AUDAT",
            "net_value": "wnd[0]/usr/subHEADER/SUB1/RBHP-NETWR",
            "currency": "wnd[0]/usr/subHEADER/SUB1/RBHP-WAERK",
        }

    try:
        from .connection import SAPConnection

        system_name = config.get("sap.system_name")
        client = config.get("sap.client")

        with SAPConnection(
            system_name=system_name,
            client=client,
            session_timeout=10.0,
        ) as sap:
            sap.open_transaction("VA23")

            for field_name, field_id in field_ids.items():
                start = time.monotonic()
                try:
                    control = sap.session.findById(field_id)
                    elapsed_ms = (time.monotonic() - start) * 1000
                    ctrl_type = type(control).__name__
                    checks.append(CompatCheck(
                        name=f"Field: {field_name}",
                        status="pass",
                        message=f"OK \u2014 {field_id}",
                        details=f"type={ctrl_type}, lookup={elapsed_ms:.0f}ms",
                        category="Fields",
                    ))
                except Exception as exc:
                    elapsed_ms = (time.monotonic() - start) * 1000
                    checks.append(CompatCheck(
                        name=f"Field: {field_name}",
                        status="fail",
                        message=f"MISSING \u2014 {field_id}",
                        details=str(exc),
                        category="Fields",
                        suggestion=(
                            f"Field '{field_id}' not found. "
                            "Check SAP GUI theme and transaction variant. "
                            "Update field mapping in config/default.yaml."
                        ),
                    ))

            with contextlib.suppress(Exception):
                sap.session.findById("wnd[0]").sendVKey(3)

    except Exception as exc:
        checks.append(CompatCheck(
            name="Field Validation",
            status="fail",
            message=f"Cannot validate fields \u2014 {exc}",
            category="Fields",
        ))

    return checks


def _check_export(config: Config) -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    office = check_office_available()
    checks.append(CompatCheck(
        name=office["name"],
        status="pass" if office["passed"] else "warning",
        message=office["message"],
        category="Export",
        suggestion=office.get("suggestion", ""),
    ))

    out = Path("./output")
    if not out.exists():
        with contextlib.suppress(OSError):
            out.mkdir(parents=True, exist_ok=True)

    if out.exists() and os.access(out, os.W_OK):
        checks.append(CompatCheck(
            name="Output Directory",
            status="pass",
            message=f"OK \u2014 {out.resolve()}",
            category="Export",
        ))
    else:
        checks.append(CompatCheck(
            name="Output Directory",
            status="fail",
            message="NOT WRITABLE",
            category="Export",
            suggestion="Check directory permissions",
        ))

    logd = Path("./logs")
    if not logd.exists():
        with contextlib.suppress(OSError):
            logd.mkdir(parents=True, exist_ok=True)

    if logd.exists() and os.access(logd, os.W_OK):
        checks.append(CompatCheck(
            name="Log Directory",
            status="pass",
            message=f"OK \u2014 {logd.resolve()}",
            category="Export",
        ))
    else:
        checks.append(CompatCheck(
            name="Log Directory",
            status="fail",
            message="NOT WRITABLE",
            category="Export",
            suggestion="Check directory permissions",
        ))

    import tempfile
    tmp = Path(tempfile.gettempdir())
    ok = tmp.exists() and os.access(tmp, os.W_OK)
    checks.append(CompatCheck(
        name="Temp Directory",
        status="pass" if ok else "warning",
        message=f"{'OK' if ok else 'NOT WRITABLE'} \u2014 {tmp}",
        category="Export",
    ))

    return checks


def _check_permissions(config: Config) -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    if sys.platform != "win32":
        checks.append(CompatCheck(
            name="Transaction Auth",
            status="warning",
            message="Cannot check (not Windows)",
            category="Permissions",
        ))
        return checks

    transactions = config.get("transactions", {})
    if transactions:
        test_tcode = next(iter(transactions.keys())).upper()
        try:
            from .connection import SAPConnection

            system_name = config.get("sap.system_name")
            client = config.get("sap.client")

            with SAPConnection(
                system_name=system_name,
                client=client,
                session_timeout=10.0,
            ) as sap:
                sap.open_transaction(test_tcode)
                checks.append(CompatCheck(
                    name="Transaction Auth",
                    status="pass",
                    message=f"OK \u2014 /n{test_tcode} authorized",
                    category="Permissions",
                ))
                with contextlib.suppress(Exception):
                    sap.session.findById("wnd[0]").sendVKey(3)
        except Exception as exc:
            msg = str(exc)
            is_auth = any(kw in msg.lower() for kw in [
                "authorization", "not authorized", "no authorization"
            ])
            checks.append(CompatCheck(
                name="Transaction Auth",
                status="fail" if is_auth else "warning",
                message=f"{'DENIED' if is_auth else 'FAILED'} \u2014 {msg[:80]}",
                category="Permissions",
                suggestion="Contact SAP security team" if is_auth else "",
            ))

    for name, path in [("Output Write", "./output"), ("Log Write", "./logs")]:
        p = Path(path)
        if p.exists():
            writable = os.access(p, os.W_OK)
            checks.append(CompatCheck(
                name=name,
                status="pass" if writable else "fail",
                message="OK" if writable else "NOT WRITABLE",
                category="Permissions",
            ))

    return checks


def _check_theme(config: Config) -> list[CompatCheck]:
    checks: list[CompatCheck] = []

    if sys.platform != "win32":
        checks.append(CompatCheck(
            name="SAP Theme",
            status="warning",
            message="Cannot detect (not Windows)",
            category="Theme",
        ))
        return checks

    try:
        from .connection import SAPConnection

        system_name = config.get("sap.system_name")
        client = config.get("sap.client")

        with SAPConnection(
            system_name=system_name,
            client=client,
            session_timeout=10.0,
        ) as sap:
            try:
                info = sap.session.Info
                theme = str(getattr(info, "Theme", "unknown"))
                checks.append(CompatCheck(
                    name="SAP Theme",
                    status="pass",
                    message=f"Detected \u2014 {theme}",
                    category="Theme",
                ))
            except Exception:
                checks.append(CompatCheck(
                    name="SAP Theme",
                    status="warning",
                    message="Cannot detect theme",
                    category="Theme",
                    suggestion=(
                        "If fields are missing, check your SAP GUI theme. "
                        "Quartz/Belize themes may need different field IDs."
                    ),
                ))

    except Exception:
        checks.append(CompatCheck(
            name="SAP Theme",
            status="warning",
            message="Cannot detect (connection failed)",
            category="Theme",
        ))

    return checks


def _generate_recommendations(report: CompatReport) -> None:
    for c in report.checks:
        if (c.status == "fail" and c.suggestion) or (c.status == "warning" and c.suggestion):
            report.add_recommendation(f"{c.name}: {c.suggestion}")

    for t in report.timings:
        if t.status == "warning":
            report.add_recommendation(
                f"Consider increasing timeout due to slow {t.name} ({t.value_ms:.0f}ms)"
            )

    theme_checks = [c for c in report.checks if c.category == "Theme"]
    if any(c.status == "warning" for c in theme_checks):
        report.add_recommendation(
            "SAP theme could not be detected. If fields are missing, "
            "check your SAP GUI theme and update field IDs in config."
        )


def run_all(config_path: str = "config/default.yaml") -> CompatReport:
    from sap_automation import __version__

    report = CompatReport(framework_version=__version__)

    try:
        config = Config.from_file(config_path)
        config.validate()
    except ConfigError as exc:
        report.add(CompatCheck(
            name="Configuration",
            status="fail",
            message=f"Config error: {exc}",
            category="Environment",
            critical=True,
        ))
        config = None

    for c in _check_environment():
        report.add(c)

    for c in _check_sap_environment():
        report.add(c)

    if config:
        for c in _check_session(config):
            report.add(c)

        for c in _check_screen(config):
            report.add(c)

        for c in _check_fields(config):
            report.add(c)

        for c in _check_theme(config):
            report.add(c)

    for c in _check_export(config or Config.from_dict({"sap": {}, "paths": {}, "export": {}, "retry": {}, "logging": {}, "transactions": {}})):
        report.add(c)

    if config:
        for c in _check_permissions(config):
            report.add(c)

    if sys.platform == "win32" and config:
        try:
            from .connection import SAPConnection

            system_name = config.get("sap.system_name")
            client = config.get("sap.client")

            with SAPConnection(
                system_name=system_name,
                client=client,
                session_timeout=5.0,
            ) as sap:
                info = sap.get_session_info()
                report.sap_info = {
                    "System": info.get("system", "unknown"),
                    "Client": info.get("client", "unknown"),
                    "User": info.get("user", "unknown"),
                    "Busy": info.get("busy", "unknown"),
                    "Alive": info.get("alive", "unknown"),
                }
        except Exception:  # noqa: S110 — best-effort COM check
            pass

    report.environment = {
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        "os": platform.system(),
        "os_release": platform.release(),
        "machine": platform.machine(),
    }

    _generate_recommendations(report)

    return report


def save_report(report: CompatReport, output_dir: str = "compatibility") -> tuple[Path, Path]:
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
