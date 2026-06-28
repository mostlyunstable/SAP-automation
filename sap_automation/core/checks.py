"""Shared diagnostic check functions used by precheck, doctor, and compatibility.

Each function returns a lightweight result dict so consumers can adapt
it to their own report types without importing a specific dataclass.
"""

from __future__ import annotations

import contextlib
import os
import platform
import shutil
import struct
import sys
from pathlib import Path
from typing import Any


def check_python_version() -> dict[str, Any]:
    v = sys.version_info
    ok = v >= (3, 10)
    return {
        "name": "Python Version",
        "passed": ok,
        "message": f"Python {v.major}.{v.minor}.{v.micro}",
        "details": "Required: >= 3.10",
        "suggestion": "Upgrade Python to 3.10 or later" if not ok else "",
    }


def check_platform() -> dict[str, Any]:
    is_win = sys.platform == "win32"
    return {
        "name": "Operating System",
        "passed": is_win,
        "message": f"{platform.system()} {platform.release()}",
        "details": f"Version: {platform.version()}, Platform: {sys.platform}",
        "suggestion": "SAP GUI automation requires Windows with COM support" if not is_win else "",
    }


def check_architecture() -> dict[str, Any]:
    bits = struct.calcsize("P") * 8
    return {
        "name": "Architecture",
        "passed": bits == 64,
        "message": f"{bits}-bit",
        "details": "Required: 64-bit",
        "suggestion": "Use 64-bit Python for best compatibility" if bits != 64 else "",
    }


def check_pywin32() -> dict[str, Any]:
    if sys.platform != "win32":
        return {
            "name": "pywin32 Package",
            "passed": False,
            "message": "Cannot check (not Windows)",
            "suggestion": "Run on Windows to check pywin32 availability",
        }
    try:
        import win32api  # noqa: F401  # type: ignore[import-untyped]
        import win32com.client  # noqa: F401  # type: ignore[import-untyped]
        return {
            "name": "pywin32 Package",
            "passed": True,
            "message": "pywin32 is installed",
        }
    except ImportError:
        return {
            "name": "pywin32 Package",
            "passed": False,
            "message": "pywin32 is not installed",
            "suggestion": "Install: pip install pywin32",
        }


def check_openpyxl() -> dict[str, Any]:
    try:
        import openpyxl
        return {
            "name": "openpyxl Package",
            "passed": True,
            "message": f"openpyxl {openpyxl.__version__}",
        }
    except ImportError:
        return {
            "name": "openpyxl Package",
            "passed": False,
            "message": "openpyxl is not installed",
            "suggestion": "Install: pip install openpyxl",
        }


def check_click() -> dict[str, Any]:
    try:
        import click
        return {
            "name": "click Package",
            "passed": True,
            "message": f"click {click.__version__}",
        }
    except ImportError:
        return {
            "name": "click Package",
            "passed": False,
            "message": "click is not installed",
            "suggestion": "Install: pip install click",
        }


def check_pyyaml() -> dict[str, Any]:
    try:
        import yaml  # noqa: F401
        return {
            "name": "PyYAML Package",
            "passed": True,
            "message": "PyYAML is installed",
        }
    except ImportError:
        return {
            "name": "PyYAML Package",
            "passed": False,
            "message": "PyYAML is not installed",
            "suggestion": "Install: pip install pyyaml",
        }


def check_sap_gui() -> dict[str, Any]:
    if sys.platform != "win32":
        return {
            "name": "SAP GUI",
            "passed": False,
            "message": "Cannot check (not Windows)",
            "suggestion": "Run on Windows to check SAP GUI availability",
        }
    try:
        import win32com.client  # type: ignore[import-untyped]
        gui = win32com.client.GetObject("SAPGUI")
        version = getattr(gui, "Version", "unknown")
        return {
            "name": "SAP GUI",
            "passed": True,
            "message": f"SAP GUI version: {version}",
            "details": "COM connection successful",
        }
    except Exception as exc:
        return {
            "name": "SAP GUI",
            "passed": False,
            "message": f"Not found: {exc}",
            "suggestion": "Install SAP GUI for Windows and ensure SAP Logon is running",
        }


def check_sap_scripting() -> dict[str, Any]:
    if sys.platform != "win32":
        return {
            "name": "SAP Scripting",
            "passed": False,
            "message": "Cannot check (not Windows)",
        }
    try:
        import win32com.client  # type: ignore[import-untyped]
        gui = win32com.client.GetObject("SAPGUI")
        app = gui.GetScriptingEngine
        count = app.Children.Count
        return {
            "name": "SAP Scripting",
            "passed": True,
            "message": f"Scripting enabled, {count} connection(s)",
        }
    except Exception as exc:
        return {
            "name": "SAP Scripting",
            "passed": False,
            "message": f"Scripting not accessible: {exc}",
            "details": "Ensure GUI Scripting is enabled in SAP Options > Accessibility",
            "suggestion": "Enable scripting: SAP Logon > Options > Accessibility > GUI Scripting > Enable",
        }


def check_office_available() -> dict[str, Any]:
    if sys.platform != "win32":
        return {
            "name": "Microsoft Office",
            "passed": False,
            "message": "Cannot check (not Windows)",
        }
    excel = None
    try:
        import win32com.client  # type: ignore[import-untyped]
        excel = win32com.client.Dispatch("Excel.Application")
        version = excel.Version
        return {
            "name": "Microsoft Office",
            "passed": True,
            "message": f"Excel {version} available",
            "details": "PDF export will work",
        }
    except Exception:
        return {
            "name": "Microsoft Office",
            "passed": False,
            "message": "Microsoft Excel not found via COM",
            "details": "PDF export requires Excel via COM",
            "suggestion": "Install Microsoft Office for PDF export support (optional)",
        }
    finally:
        if excel:
            with contextlib.suppress(Exception):
                excel.Quit()
                del excel


def check_disk_space() -> dict[str, Any]:
    try:
        usage = shutil.disk_usage(".")
        free_gb = usage.free / (1024 ** 3)
        if free_gb < 1:
            return {
                "name": "Disk Space",
                "passed": False,
                "message": f"LOW ({free_gb:.1f} GB free)",
                "suggestion": "Free up disk space before running automation",
            }
        if free_gb < 5:
            return {
                "name": "Disk Space",
                "passed": True,
                "message": f"LOW ({free_gb:.1f} GB free)",
                "suggestion": "Consider freeing up disk space",
            }
        return {
            "name": "Disk Space",
            "passed": True,
            "message": f"OK ({free_gb:.1f} GB free)",
        }
    except Exception as exc:
        return {
            "name": "Disk Space",
            "passed": True,
            "message": f"Cannot determine: {exc}",
        }


def check_output_directory(config_path: str | None = None) -> dict[str, Any]:
    output_dir = Path("./output")
    if config_path:
        try:
            from .config import Config
            cfg = Config.from_file(config_path)
            output_dir = Path(cfg.get("paths.output_dir", "./output"))
        except Exception:  # noqa: S110 — best-effort config load
            pass

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        test_file = output_dir / ".write_test"
        test_file.write_text("test", encoding="utf-8")
        test_file.unlink()
        return {
            "name": "Output Directory",
            "passed": True,
            "message": f"Writable: {output_dir.resolve()}",
        }
    except Exception as exc:
        return {
            "name": "Output Directory",
            "passed": False,
            "message": f"Cannot write to {output_dir}: {exc}",
            "suggestion": "Check directory permissions or choose a different output path",
        }


def check_log_directory(config_path: str | None = None) -> dict[str, Any]:
    log_dir = Path("./logs")
    if config_path:
        try:
            from .config import Config
            cfg = Config.from_file(config_path)
            log_dir = Path(cfg.get("paths.log_dir", "./logs"))
        except Exception:  # noqa: S110 — best-effort config load
            pass

    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        test_file = log_dir / ".write_test"
        test_file.write_text("test", encoding="utf-8")
        test_file.unlink()
        return {
            "name": "Log Directory",
            "passed": True,
            "message": f"Writable: {log_dir.resolve()}",
        }
    except Exception as exc:
        return {
            "name": "Log Directory",
            "passed": False,
            "message": f"Cannot write to {log_dir}: {exc}",
            "suggestion": "Check directory permissions or choose a different log path",
        }


def check_secrets(config_path: str = "config/default.yaml") -> dict[str, Any]:
    import re
    path = Path(config_path)
    if not path.exists():
        return {
            "name": "Secrets Scan",
            "passed": True,
            "message": "Cannot scan (config not found)",
        }

    raw = path.read_text(encoding="utf-8")
    secret_patterns = [
        (r"password\s*:\s*['\"]?(?!.*\$\{)(?!.*null)[^\s'\"]+", "password"),
        (r"api_key\s*:\s*['\"]?(?!.*\$\{)(?!.*null)[^\s'\"]+", "api_key"),
        (r"secret\s*:\s*['\"]?(?!.*\$\{)(?!.*null)[^\s'\"]+", "secret"),
        (r"token\s*:\s*['\"]?(?!.*\$\{)(?!.*null)[^\s'\"]+", "token"),
    ]

    found: list[str] = []
    for pattern, label in secret_patterns:
        matches = re.findall(pattern, raw, re.IGNORECASE)
        if matches:
            found.append(label)

    if found:
        return {
            "name": "Secrets Scan",
            "passed": False,
            "message": f"Potential secrets found: {', '.join(found)}",
            "suggestion": "Use environment variables (${VAR}) instead of hardcoded values",
        }

    return {
        "name": "Secrets Scan",
        "passed": True,
        "message": "No hardcoded secrets detected",
    }


def check_env_vars(config_path: str = "config/default.yaml") -> dict[str, Any]:
    import re
    path = Path(config_path)
    if not path.exists():
        return {
            "name": "Environment Vars",
            "passed": True,
            "message": "Cannot check (config not found)",
        }

    raw = path.read_text(encoding="utf-8")
    all_vars = re.findall(r"\$\{(\w+)\}", raw)
    if not all_vars:
        return {
            "name": "Environment Vars",
            "passed": True,
            "message": "No env vars referenced",
        }

    unresolved = [v for v in all_vars if not os.environ.get(v)]
    if unresolved:
        return {
            "name": "Environment Vars",
            "passed": False,
            "message": f"{len(unresolved)} unresolved: {', '.join(set(unresolved))}",
            "suggestion": "Set missing environment variables before running",
        }

    return {
        "name": "Environment Vars",
        "passed": True,
        "message": f"All {len(set(all_vars))} env var(s) resolved",
    }
