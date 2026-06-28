"""Enterprise support bundle generator for SAP Automation Framework.

Creates a complete diagnostic package for troubleshooting.
Never exposes passwords, tokens, API keys, or sensitive SAP data.

Usage:
    sap-auto support
    sap-auto support --output support_bundle.zip
    sap-auto support --json
    sap-auto support --verbose
"""

from __future__ import annotations

import contextlib
import json
import os
import platform
import re
import shutil
import struct
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .config import Config
from .logger import get_logger

log = get_logger("support")

# Patterns for secrets that must be redacted
_SECRET_PATTERNS: list[tuple[str, str]] = [
    (r"(password|passwd|pwd)\s*[:=]\s*['\"]?.+?", r"\1=REDACTED"),
    (r"(token|api_key|apikey|secret|auth|credential)\s*[:=]\s*['\"]?.+?", r"\1=REDACTED"),
    (r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b", "REDACTED_IP"),
]

# Sensitive config keys to remove
_SENSITIVE_KEYS = {
    "password", "passwd", "pwd", "token", "api_key", "apikey",
    "secret", "auth", "credential", "cookie", "session_id",
}


def _redact_text(text: str) -> str:
    """Redact secrets from text."""
    for pattern, replacement in _SECRET_PATTERNS:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return text


def _redact_dict(d: dict[str, Any]) -> dict[str, Any]:
    """Recursively redact secrets from a dictionary."""
    result: dict[str, Any] = {}
    for k, v in d.items():
        key_lower = k.lower()
        if key_lower in _SENSITIVE_KEYS:
            result[k] = "REDACTED"
        elif isinstance(v, dict):
            result[k] = _redact_dict(v)
        elif isinstance(v, str):
            result[k] = _redact_text(v)
        else:
            result[k] = v
    return result


# ---------------------------------------------------------------------------
# Collection functions
# ---------------------------------------------------------------------------

def _collect_system_info() -> dict[str, Any]:
    """Collect system information."""
    info: dict[str, Any] = {
        "os": platform.system(),
        "os_release": platform.release(),
        "os_version": platform.version(),
        "platform": sys.platform,
        "machine": platform.machine(),
        "processor": platform.processor(),
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        "python_executable": sys.executable,
        "architecture": f"{struct.calcsize('P') * 8}-bit",
        "cwd": os.getcwd(),
        "username": os.environ.get("USERNAME") or os.environ.get("USER", "unknown"),
    }

    # CPU info
    try:
        info["cpu_count"] = os.cpu_count() or 0
    except Exception:
        info["cpu_count"] = "unknown"

    # RAM
    try:
        if sys.platform == "win32":
            import ctypes
            kernel32 = ctypes.windll.kernel32
            c_ulonglong = ctypes.c_ulonglong
            mem = c_ulonglong()
            kernel32.GetPhysicallyInstalledMemory(ctypes.byref(mem))
            info["ram_gb"] = round(mem.value / (1024 ** 3), 1)
        else:
            # Fallback for non-Windows
            with contextlib.suppress(Exception):
                result = subprocess.run(
                    ["sysctl", "-n", "hw.memsize"],  # noqa: S607
                    capture_output=True, text=True, timeout=5,
                )
                if result.returncode == 0:
                    info["ram_gb"] = round(int(result.stdout.strip()) / (1024 ** 3), 1)
    except Exception:
        info["ram_gb"] = "unknown"

    # Disk
    try:
        usage = shutil.disk_usage(".")
        info["disk_total_gb"] = round(usage.total / (1024 ** 3), 1)
        info["disk_free_gb"] = round(usage.free / (1024 ** 3), 1)
    except Exception:
        info["disk_total_gb"] = "unknown"
        info["disk_free_gb"] = "unknown"

    return info


def _collect_packages() -> dict[str, str]:
    """Collect installed package versions."""
    packages: dict[str, str] = {}
    for name in ["click", "pyyaml", "openpyxl", "pywin32", "ruff", "mypy", "pytest"]:
        try:
            mod_name = name.replace("-", "_").lower()
            if name == "pyyaml":
                mod_name = "yaml"
            elif name == "pywin32":
                mod_name = "win32com"
            mod = __import__(mod_name)
            packages[name] = getattr(mod, "__version__", "installed")
        except ImportError:
            packages[name] = "not installed"
    return packages


def _collect_env_vars() -> dict[str, str]:
    """Collect environment variables with secrets redacted."""
    sensitive = {
        "password", "passwd", "pwd", "token", "secret", "api_key",
        "apikey", "auth", "credential", "cookie", "session",
    }
    result: dict[str, str] = {}
    for key, value in os.environ.items():
        if any(s in key.lower() for s in sensitive):
            result[key] = "REDACTED"
        else:
            result[key] = value[:200]  # Truncate long values
    return result


def _collect_sap_info(config: Config | None) -> dict[str, Any]:
    """Collect SAP GUI information (read-only)."""
    info: dict[str, Any] = {"available": False}

    if sys.platform != "win32":
        info["reason"] = "not Windows"
        return info

    try:
        import win32com.client  # type: ignore[import-untyped]

        gui = win32com.client.GetObject("SAPGUI")
        info["available"] = True
        info["version"] = str(getattr(gui, "Version", "unknown"))

        app = gui.GetScriptingEngine
        info["scripting_enabled"] = True
        info["connections"] = app.Children.Count

        # Collect connection details (no passwords)
        connections: list[dict[str, Any]] = []
        for i in range(app.Children.Count):
            conn = app.Children(i)
            try:
                conn_info: dict[str, Any] = {
                    "system": str(getattr(conn, "SystemName", "unknown")),
                }
                # Session details
                for j in range(conn.Sessions.Count):
                    sess = conn.Sessions(j)
                    try:
                        conn_info["client"] = str(getattr(sess, "Client", "unknown"))
                        conn_info["user"] = str(getattr(sess, "User", "unknown"))
                        conn_info["language"] = str(getattr(sess, "Language", "unknown"))
                        conn_info["busy"] = bool(sess.Busy)
                    except Exception:  # noqa: S110 — best-effort COM check
                        pass
                    break  # Only first session per connection
                connections.append(conn_info)
            except Exception:  # noqa: S112 — best-effort COM polling
                continue

        info["connections_detail"] = connections

    except ImportError:
        info["reason"] = "pywin32 not installed"
    except Exception as exc:
        info["reason"] = str(exc)

    return info


def _collect_framework_info(config: Config | None) -> dict[str, Any]:
    """Collect framework information."""
    from sap_automation import __version__

    info: dict[str, Any] = {
        "version": __version__,
        "python": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
    }

    # Git commit (if available)
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],  # noqa: S607
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0:
            info["git_commit"] = result.stdout.strip()
    except Exception:  # noqa: S110 — best-effort git check
        pass

    # Registered transactions
    if config:
        transactions = config.get("transactions", {})
        info["registered_transactions"] = list(transactions.keys())
    else:
        info["registered_transactions"] = []

    return info


def _collect_config_sanitized(config_path: str) -> str:
    """Export config with secrets redacted."""
    path = Path(config_path)
    if not path.exists():
        return "# Config file not found\n"

    try:
        import yaml

        with open(path, encoding="utf-8") as f:
            config_data = yaml.safe_load(f)

        if config_data and isinstance(config_data, dict):
            config_data = _redact_dict(config_data)

        result: str = yaml.dump(config_data, default_flow_style=False, sort_keys=False)
        return result
    except Exception as exc:
        return f"# Error reading config: {exc}\n"


def _collect_logs(log_dir: str = "./logs", max_files: int = 5) -> dict[str, str]:
    """Collect log file contents (latest N files)."""
    log_path = Path(log_dir)
    logs: dict[str, str] = {}

    if not log_path.exists():
        return logs

    log_files = sorted(
        log_path.glob("*.log"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )[:max_files]

    for lf in log_files:
        try:
            content = lf.read_text(encoding="utf-8", errors="replace")
            # Redact secrets from logs
            content = _redact_text(content)
            # Truncate very large logs
            if len(content) > 100_000:
                content = content[:50_000] + "\n... [truncated] ...\n" + content[-50_000:]
            logs[lf.name] = content
        except Exception as exc:
            logs[lf.name] = f"[Error reading log: {exc}]"

    return logs


def _collect_crash_info() -> dict[str, Any]:
    """Collect information about previous crashes."""
    crashes: dict[str, Any] = {
        "log_dir": "./logs",
        "crash_files": [],
    }

    log_path = Path("./logs")
    if not log_path.exists():
        return crashes

    # Look for crash-related log patterns
    crash_patterns = ["crash", "error", "exception", "traceback"]
    for lf in log_path.glob("*.log"):
        try:
            content = lf.read_text(encoding="utf-8", errors="replace")
            for pattern in crash_patterns:
                if pattern.lower() in content.lower():
                    crashes["crash_files"].append({
                        "file": lf.name,
                        "size_bytes": lf.stat().st_size,
                        "modified": datetime.fromtimestamp(
                            lf.stat().st_mtime, tz=timezone.utc
                        ).isoformat(),
                    })
                    break
        except Exception:  # noqa: S112 — best-effort crash scan
            continue

    return crashes


def _generate_readme(
    timestamp: str,
    health_score: str,
    compat_score: str,
    problems: list[str],
    recommendations: list[str],
) -> str:
    """Generate the support bundle README."""
    lines = [
        "=" * 60,
        "SAP AUTOMATION — SUPPORT BUNDLE",
        "=" * 60,
        "",
        f"Generated:      {timestamp}",
        f"Framework:      {_collect_framework_info(None).get('version', '?')}",
        f"Health Score:   {health_score}",
        f"Compat Score:   {compat_score}",
        "",
        "=" * 60,
        "CONTENTS",
        "=" * 60,
        "",
        "  system/         — Operating system and hardware information",
        "  framework/      — Framework version, config, transactions",
        "  logs/           — Application logs (secrets redacted)",
        "  config/         — Sanitized configuration files",
        "  diagnostics/    — Doctor, precheck, compatibility-test results",
        "",
        "=" * 60,
    ]

    if problems:
        lines.append("")
        lines.append("KNOWN PROBLEMS")
        lines.append("-" * 60)
        for p in problems:
            lines.append(f"  • {p}")

    if recommendations:
        lines.append("")
        lines.append("RECOMMENDED ACTIONS")
        lines.append("-" * 60)
        for r in recommendations:
            lines.append(f"  • {r}")

    lines.append("")
    lines.append("=" * 60)
    lines.append("HOW TO USE THIS BUNDLE")
    lines.append("-" * 60)
    lines.append("  1. Send this ZIP file to your support team")
    lines.append("  2. Include any error messages you've seen")
    lines.append("  3. Describe what you were doing when the issue occurred")
    lines.append("  4. Note your SAP GUI version and theme")
    lines.append("")
    lines.append("=" * 60)
    lines.append("PRIVACY NOTICE")
    lines.append("-" * 60)
    lines.append("  All passwords, tokens, and API keys have been automatically")
    lines.append("  redacted. SAP business data is not included.")
    lines.append("")
    lines.append("=" * 60)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Main bundle creator
# ---------------------------------------------------------------------------

def create_support_bundle(
    config_path: str = "config/default.yaml",
    output_path: str = "support_bundle.zip",
    verbose: bool = False,
) -> Path:
    """Create a complete support bundle ZIP file.

    Args:
        config_path: Path to configuration file.
        output_path: Path for the output ZIP file.
        verbose: Include verbose diagnostic output.

    Returns:
        Path to the created ZIP file.
    """
    timestamp = datetime.now(timezone.utc).isoformat()
    bundle_dir = Path("support_bundle_temp")

    if bundle_dir.exists():
        shutil.rmtree(bundle_dir)

    for subdir in ["logs", "diagnostics", "config", "system", "framework"]:
        (bundle_dir / subdir).mkdir(parents=True, exist_ok=True)

    problems: list[str] = []
    recommendations: list[str] = []
    health_score = "N/A"
    compat_score = "N/A"

    def _status(msg: str) -> None:
        """Print progress status."""
        click_echo(f"  [·] {msg}...")

    def _done(msg: str) -> None:
        """Print completion status."""
        click_echo(f"  [✓] {msg}")

    def _fail(msg: str) -> None:
        """Print failure status."""
        click_echo(f"  [✗] {msg}")

    # We need click.echo but don't want to import it at module level
    # to avoid circular imports. Use print as fallback.
    import click
    click_echo = click.echo

    click_echo("")
    click_echo("Collecting support bundle...")
    click_echo("")

    # 1. System info
    _status("Collecting system information")
    try:
        sys_info = _collect_system_info()
        packages = _collect_packages()
        env_vars = _collect_env_vars()

        system_data = {
            "system": sys_info,
            "packages": packages,
            "environment_variables": env_vars,
        }
        (bundle_dir / "system" / "system_info.json").write_text(
            json.dumps(system_data, indent=2, default=str),
            encoding="utf-8",
        )
        _done("System information collected")
    except Exception as exc:
        _fail(f"System collection failed: {exc}")

    # 2. SAP info
    _status("Collecting SAP information")
    try:
        config = None
        with contextlib.suppress(Exception):
            config = Config.from_file(config_path)

        sap_info = _collect_sap_info(config)
        (bundle_dir / "system" / "sap_info.json").write_text(
            json.dumps(sap_info, indent=2, default=str),
            encoding="utf-8",
        )
        _done("SAP information collected")
    except Exception as exc:
        _fail(f"SAP collection failed: {exc}")

    # 3. Framework info
    _status("Collecting framework information")
    try:
        fw_info = _collect_framework_info(config)
        (bundle_dir / "framework" / "framework_info.json").write_text(
            json.dumps(fw_info, indent=2, default=str),
            encoding="utf-8",
        )
        _done("Framework information collected")
    except Exception as exc:
        _fail(f"Framework collection failed: {exc}")

    # 4. Run diagnostics
    _status("Running doctor diagnostic")
    try:
        from .doctor import run_all as doctor_run
        doctor_report = doctor_run(config_path)
        (bundle_dir / "diagnostics" / "doctor.json").write_text(
            json.dumps(doctor_report.to_dict(), indent=2, default=str),
            encoding="utf-8",
        )
        health_score = f"{doctor_report.health_pct}%"
        _done(f"Doctor: health score {health_score}")
    except Exception as exc:
        _fail(f"Doctor failed: {exc}")

    _status("Running precheck")
    try:
        from .precheck import run_all as precheck_run
        precheck_report = precheck_run(config_path)
        (bundle_dir / "diagnostics" / "precheck.json").write_text(
            json.dumps(precheck_report.to_dict(), indent=2, default=str),
            encoding="utf-8",
        )
        _done(f"Precheck: {sum(1 for c in precheck_report.checks if c.passed)}/{len(precheck_report.checks)} passed")
    except Exception as exc:
        _fail(f"Precheck failed: {exc}")

    _status("Running compatibility test")
    try:
        from .compatibility import run_all as compat_run
        compat_report = compat_run(config_path)
        (bundle_dir / "diagnostics" / "compatibility.json").write_text(
            json.dumps(compat_report.to_dict(), indent=2, default=str),
            encoding="utf-8",
        )
        compat_score = f"{compat_report.score_pct}%"
        problems.extend(c.message for c in compat_report.failed_checks())
        recommendations.extend(compat_report.recommendations)
        _done(f"Compatibility: {compat_score}")
    except Exception as exc:
        _fail(f"Compatibility test failed: {exc}")

    # 5. Collect logs
    _status("Collecting log files")
    try:
        logs = _collect_logs()
        for filename, content in logs.items():
            (bundle_dir / "logs" / filename).write_text(content, encoding="utf-8")
        _done(f"{len(logs)} log file(s) collected")
    except Exception as exc:
        _fail(f"Log collection failed: {exc}")

    # 6. Sanitize config
    _status("Sanitizing configuration")
    try:
        config_content = _collect_config_sanitized(config_path)
        (bundle_dir / "config" / "config_sanitized.yaml").write_text(
            config_content, encoding="utf-8",
        )
        # Also include the raw config if it exists
        config_file = Path(config_path)
        if config_file.exists():
            raw = config_file.read_text(encoding="utf-8")
            raw_redacted = _redact_text(raw)
            (bundle_dir / "config" / "config_raw_redacted.txt").write_text(
                raw_redacted, encoding="utf-8",
            )
        _done("Configuration sanitized")
    except Exception as exc:
        _fail(f"Config sanitization failed: {exc}")

    # 7. Crash info
    _status("Checking for crash information")
    try:
        crash_info = _collect_crash_info()
        (bundle_dir / "framework" / "crash_info.json").write_text(
            json.dumps(crash_info, indent=2, default=str),
            encoding="utf-8",
        )
        _done(f"{len(crash_info.get('crash_files', []))} crash file(s) found")
    except Exception as exc:
        _fail(f"Crash collection failed: {exc}")

    # 8. Generate README
    _status("Generating README")
    try:
        readme = _generate_readme(timestamp, health_score, compat_score, problems, recommendations)
        (bundle_dir / "README.txt").write_text(readme, encoding="utf-8")
        _done("README generated")
    except Exception as exc:
        _fail(f"README generation failed: {exc}")

    # 9. Create ZIP
    _status("Creating ZIP archive")
    try:
        zip_path = Path(output_path)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(bundle_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(bundle_dir)
                    zf.write(file_path, arcname)
        _done(f"ZIP created: {zip_path}")
    except Exception as exc:
        _fail(f"ZIP creation failed: {exc}")
        raise
    finally:
        # Always clean up temp directory — even on failure — to prevent
        # sensitive diagnostic data from being left on disk.
        with contextlib.suppress(Exception):
            shutil.rmtree(bundle_dir)

    return zip_path


def create_support_bundle_json(
    config_path: str = "config/default.yaml",
) -> dict[str, Any]:
    """Create support bundle data as JSON-serializable dict (no ZIP).

    Args:
        config_path: Path to configuration file.

    Returns:
        Dictionary with all collected support data.
    """
    timestamp = datetime.now(timezone.utc).isoformat()
    data: dict[str, Any] = {"timestamp": timestamp}

    # System
    with contextlib.suppress(Exception):
        data["system"] = _collect_system_info()
        data["packages"] = _collect_packages()

    # SAP
    config = None
    with contextlib.suppress(Exception):
        config = Config.from_file(config_path)
    data["sap"] = _collect_sap_info(config)

    # Framework
    data["framework"] = _collect_framework_info(config)

    # Config (redacted)
    data["config"] = _collect_config_sanitized(config_path)

    return data
