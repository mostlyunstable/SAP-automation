"""Report exporter supporting Excel, CSV, and PDF formats with atomic writes."""

from __future__ import annotations

import contextlib
import csv
import os
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from ..core.exceptions import ExportError
from ..core.logger import get_logger

log = get_logger("exporter")

# Characters unsafe in filenames (Windows + Unix)
_UNSAFE_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_filename(name: str, max_length: int = 200) -> str:
    """Remove unsafe characters and truncate to max_length.

    Args:
        name: Proposed filename.
        max_length: Maximum character length.

    Returns:
        Sanitized filename string.
    """
    if not name:
        return "report"

    sanitized = _UNSAFE_FILENAME_CHARS.sub("", name)
    if not sanitized:
        return "report"

    if len(sanitized) > max_length:
        sanitized = sanitized[:max_length]

    return sanitized


def _sanitize_csv_value(value: Any) -> str:
    """Prevent CSV injection by prefixing formula characters.

    Characters =, +, -, @, and tab at the start of a value are prefixed
    with a single quote to neutralize formula execution in Excel.
    """
    str_value = str(value)
    if str_value and str_value[0] in ("=", "+", "-", "@", "\t"):
        return f"'{str_value}"
    return str_value


class ReportExporter:
    """Export data to Excel, CSV, or PDF files with atomic writes.

    Args:
        output_dir: Directory for output files.
        timestamp_in_name: Whether to append timestamp to filenames.
    """

    def __init__(
        self,
        output_dir: str | Path = "./output",
        timestamp_in_name: bool = True,
    ) -> None:
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.timestamp_in_name = timestamp_in_name

    def _make_path(self, filename: str, ext: str) -> Path:
        """Build full output path with optional timestamp."""
        name = sanitize_filename(filename)
        if self.timestamp_in_name:
            ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
            name = f"{name}_{ts}"
        return self.output_dir / f"{name}{ext}"

    def export_excel(
        self,
        data: list[dict[str, Any]],
        filename: str = "report",
        sheet_name: str | None = None,
    ) -> Path:
        """Export data to an Excel file.

        Args:
            data: List of dictionaries to write.
            filename: Base filename (without extension).
            sheet_name: Custom sheet name. Defaults to 'Data'.

        Returns:
            Path to the created file.

        Raises:
            ExportError: If data is empty or write fails.
        """
        if not data:
            raise ExportError("No data to export")

        path = self._make_path(filename, ".xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name or "Data"

        # Headers
        headers = list(data[0].keys())
        ws.append(headers)

        # Rows
        for row in data:
            ws.append([row.get(h, "") for h in headers])

        # Atomic write
        tmp_fd, tmp_path = tempfile.mkstemp(
            dir=self.output_dir, suffix=".tmp"
        )
        try:
            os.close(tmp_fd)
            wb.save(tmp_path)
            os.replace(tmp_path, path)
        except Exception:
            with contextlib.suppress(OSError):
                os.unlink(tmp_path)
            raise
        finally:
            wb.close()

        log.info("Exported Excel: %s (%d rows)", path, len(data))
        self._verify_output(path, "Excel")
        return path

    def export_csv(
        self,
        data: list[dict[str, Any]],
        filename: str = "report",
    ) -> Path:
        """Export data to a CSV file with injection prevention.

        Uses UTF-8-BOM encoding for Excel compatibility.

        Args:
            data: List of dictionaries to write.
            filename: Base filename (without extension).

        Returns:
            Path to the created file.

        Raises:
            ExportError: If data is empty or write fails.
        """
        if not data:
            raise ExportError("No data to export")

        path = self._make_path(filename, ".csv")
        headers = list(data[0].keys())

        # Atomic write
        tmp_fd, tmp_path = tempfile.mkstemp(
            dir=self.output_dir, suffix=".tmp"
        )
        try:
            with os.fdopen(tmp_fd, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for row in data:
                    writer.writerow([
                        _sanitize_csv_value(row.get(h, "")) for h in headers
                    ])
            os.replace(tmp_path, path)
        except Exception:
            with contextlib.suppress(OSError):
                os.unlink(tmp_path)
            raise

        log.info("Exported CSV: %s (%d rows)", path, len(data))
        self._verify_output(path, "CSV")
        return path

    def export_pdf(
        self,
        data: list[dict[str, Any]],
        filename: str = "report",
    ) -> Path:
        """Export data to a PDF file using Excel COM automation.

        Requires pywin32 (Windows only).

        Args:
            data: List of dictionaries to write.
            filename: Base filename (without extension).

        Returns:
            Path to the created file.

        Raises:
            ExportError: If pywin32 is unavailable, data is empty, or export fails.
        """
        if not data:
            raise ExportError("No data to export")

        try:
            import win32com.client  # type: ignore[import-untyped]
        except ImportError:
            raise ExportError(
                "PDF export requires pywin32. "
                "Install it: pip install pywin32"
            ) from None

        # Create temp Excel first — use a fixed temp name for reliable cleanup
        tmp_xlsx_name = f"_pdf_tmp_{filename}"
        xlsx_path = self._make_path(tmp_xlsx_name, ".xlsx")
        self.export_excel(data, filename=tmp_xlsx_name)

        path = self._make_path(filename, ".pdf")

        excel = None
        wb_com = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb_com = excel.Workbooks.Open(str(xlsx_path.resolve()))
            wb_com.ExportAsFixedFormat(0, str(path.resolve()))  # 0 = PDF
        except Exception as exc:
            raise ExportError(f"PDF export failed: {exc}") from exc
        finally:
            if wb_com:
                with contextlib.suppress(Exception):
                    wb_com.Close(False)
                    del wb_com
            if excel:
                with contextlib.suppress(Exception):
                    excel.Quit()
                    del excel
            # Clean up temp xlsx
            with contextlib.suppress(OSError):
                os.unlink(xlsx_path)

        log.info("Exported PDF: %s (%d rows)", path, len(data))
        self._verify_output(path, "PDF")
        return path

    def _verify_output(self, path: Path, format_name: str) -> None:
        """Verify that an exported file exists and is valid.

        Args:
            path: Path to the exported file.
            format_name: Format name for error messages.

        Raises:
            ExportError: If verification fails.
        """
        if not path.exists():
            raise ExportError(
                f"{format_name} export file not created: {path}",
                details={
                    "path": str(path),
                    "suggestion": "Check output directory permissions and disk space.",
                },
            )

        if not path.is_file():
            raise ExportError(
                f"{format_name} export path is not a file: {path}",
                details={"path": str(path)},
            )

        file_size = path.stat().st_size
        if file_size == 0:
            raise ExportError(
                f"{format_name} export file is empty: {path}",
                details={
                    "path": str(path),
                    "suggestion": "The export may have failed silently. Check for write errors.",
                },
            )

        log.debug(
            "Verified %s export: %s (%d bytes)",
            format_name,
            path,
            file_size,
        )

    def export(
        self,
        data: list[dict[str, Any]],
        filename: str = "report",
        formats: list[str] | None = None,
    ) -> list[Path]:
        """Export data to one or more formats.

        Args:
            data: List of dictionaries to write.
            filename: Base filename (without extension).
            formats: List of format strings ('xlsx', 'csv', 'pdf', 'all').

        Returns:
            List of paths to created files.
        """
        if not data:
            raise ExportError("No data to export")

        if formats is None:
            formats = ["xlsx"]

        paths: list[Path] = []
        exporters = {
            "xlsx": self.export_excel,
            "csv": self.export_csv,
            "pdf": self.export_pdf,
        }

        for fmt in formats:
            fmt = fmt.lower().strip()
            if fmt == "all":
                for key, func in exporters.items():
                    if key == "pdf":
                        try:
                            paths.append(func(data, filename))
                        except ExportError:
                            log.warning("Skipping PDF export (unavailable)")
                    else:
                        paths.append(func(data, filename))
            elif fmt in exporters:
                paths.append(exporters[fmt](data, filename))
            else:
                log.warning("Unknown format: %s (skipping)", fmt)

        return paths
