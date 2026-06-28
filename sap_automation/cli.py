"""Click CLI for SAP Automation Framework.

Commands:
    run                 Run a transaction for all records in an Excel file
    validate            Validate a configuration file
    precheck            Run environment compatibility checks
    doctor              Run comprehensive environment diagnostics
    compatibility-test  Run deep SAP compatibility validation
    support             Create support bundle for troubleshooting
    list-transactions   List configured transaction handlers
    init                Initialize output directory structure
"""

from __future__ import annotations

import json
from pathlib import Path

import click

from . import __version__
from .core.config import Config, ConfigError
from .core.logger import setup_logging
from .utils.excel_reader import ExcelReader


@click.group()
@click.version_option(version=__version__, prog_name="SAP Automation")
def main() -> None:
    """SAP Automation — Batch transaction processing framework."""


@main.command()
@click.option("-c", "--config", "config_path", default="config/default.yaml", help="Path to config YAML file.")
@click.option("-t", "--transaction", "txn_name", required=True, help="Transaction name from config.")
@click.option("-i", "--input", "input_path", required=True, help="Path to Excel input file.")
@click.option("-o", "--output", "output_dir", default=None, help="Output directory override.")
@click.option("--export", "export_format", default=None, help="Export format (xlsx, csv, pdf, all).")
@click.option("--dry-run", is_flag=True, default=False, help="Validate without modifying SAP.")
def run(
    config_path: str,
    txn_name: str,
    input_path: str,
    output_dir: str | None,
    export_format: str | None,
    dry_run: bool,
) -> None:
    """Run a transaction for all records in an Excel file."""
    # Load config
    try:
        config = Config.from_file(config_path)
        config.validate()
    except ConfigError as exc:
        click.echo(f"Config error: {exc}", err=True)
        raise SystemExit(1) from exc

    if output_dir:
        config.set("paths.output_dir", output_dir)

    if export_format:
        config.set("export.format", export_format)

    try:
        doc_col = config.get("excel.document_number_column", "Document Number")
        reader = ExcelReader(input_path, document_number_column=doc_col)
        records = reader.read()
    except Exception as exc:
        click.echo(f"Input error: {exc}", err=True)
        raise SystemExit(1) from exc

    if not records:
        click.echo("No records found in input file.", err=True)
        raise SystemExit(1)

    click.echo(f"Loaded {len(records)} records from {input_path}")

    log_dir = config.get("paths.log_dir", "./logs")
    log_level = config.get("logging.level", "INFO")
    setup_logging(log_dir=log_dir, level=log_level)

    from .orchestrator import Orchestrator

    orchestrator = Orchestrator(config)

    if dry_run:
        click.echo("Running dry run (no SAP modifications)...")
        summary = orchestrator.dry_run(txn_name, records)
    else:
        summary = orchestrator.run(txn_name, records)

    click.echo(summary.report())

    output_path = Path(config.get("paths.output_dir", "./output"))
    output_path.mkdir(parents=True, exist_ok=True)
    summary_file = output_path / f"summary_{summary.run_id}.json"
    with open(summary_file, "w", encoding="utf-8") as f:
        json.dump(summary.to_dict(), f, indent=2, default=str)
    click.echo(f"Summary saved: {summary_file}")

    if summary.failed > 0:
        raise SystemExit(1)


@main.command()
@click.option("-c", "--config", "config_path", default="config/default.yaml", help="Path to config YAML file.")
def validate(config_path: str) -> None:
    """Validate a configuration file."""
    try:
        config = Config.from_file(config_path)
        config.validate()
        click.echo(f"Config is valid: {config_path}")
    except ConfigError as exc:
        click.echo(f"Validation failed: {exc}", err=True)
        raise SystemExit(1) from exc


@main.command()
@click.option("-c", "--config", "config_path", default=None, help="Path to config YAML file (optional).")
@click.option("--json-output", is_flag=True, default=False, help="Output as JSON.")
def precheck(config_path: str | None, json_output: bool) -> None:
    """Run environment compatibility checks.

    Checks Python version, OS, SAP GUI, COM scripting,
    required packages, and output directories.
    """
    from .core import precheck as precheck_mod

    report = precheck_mod.run_all(config_path)

    if json_output:
        click.echo(json.dumps(report.to_dict(), indent=2))
    else:
        click.echo(report.format())

    if not report.is_compatible:
        raise SystemExit(1)


@main.command("list-transactions")
@click.option("-c", "--config", "config_path", default="config/default.yaml", help="Path to config YAML file.")
def list_transactions(config_path: str) -> None:
    """List configured transaction handlers."""
    try:
        config = Config.from_file(config_path)
        config.validate()
    except ConfigError as exc:
        click.echo(f"Config error: {exc}", err=True)
        raise SystemExit(1) from exc

    transactions = config.get("transactions", {})
    if not transactions:
        click.echo("No transactions configured.")
        return

    click.echo("Configured transactions:")
    click.echo("-" * 60)
    for name, txn in transactions.items():
        desc = txn.get("description", "")
        cls = txn.get("class", "")
        click.echo(f"  {name:15s} {desc}")
        click.echo(f"  {'':15s} Class: {cls}")
        click.echo()


@main.command()
@click.argument("dirname", default=".")
def init(dirname: str) -> None:
    """Initialize output directory structure.

    Creates output/, logs/, and config/ directories in the specified path.
    """
    path = Path(dirname)
    (path / "output").mkdir(parents=True, exist_ok=True)
    (path / "logs").mkdir(parents=True, exist_ok=True)
    (path / "config").mkdir(parents=True, exist_ok=True)
    click.echo(f"Initialized directory structure in: {path.resolve()}")
    click.echo("  created: output/")
    click.echo("  created: logs/")
    click.echo("  created: config/")


@main.command()
@click.option("-c", "--config", "config_path", default="config/default.yaml", help="Path to config YAML file.")
@click.option("--json", "json_output", is_flag=True, default=False, help="Output as machine-readable JSON.")
@click.option("--verbose", is_flag=True, default=False, help="Show detailed diagnostics.")
@click.option("--save", "save_report", is_flag=True, default=False, help="Save report to diagnostics/ directory.")
def doctor(config_path: str, json_output: bool, verbose: bool, save_report: bool) -> None:
    """Run comprehensive environment diagnostics.

    Performs read-only health checks of system, Python, SAP, configuration,
    file system, Office, logging, and security. Never modifies SAP or config.

    Exit codes:
        0  READY or READY WITH WARNINGS
        1  NOT READY (critical failures detected)
    """
    from .core.doctor import run_all
    from .core.doctor import save_report as do_save

    report = run_all(config_path)

    if json_output:
        click.echo(json.dumps(report.to_dict(), indent=2, default=str))
    else:
        click.echo(report.format_console(verbose=verbose))

    if save_report:
        json_path, txt_path = do_save(report)
        click.echo("\nReport saved:")
        click.echo(f"  {json_path}")
        click.echo(f"  {txt_path}")

    if report.overall_status == "NOT READY":
        raise SystemExit(1)


@main.command("compatibility-test")
@click.option("-c", "--config", "config_path", default="config/default.yaml", help="Path to config YAML file.")
@click.option("--json", "json_output", is_flag=True, default=False, help="Output as machine-readable JSON.")
@click.option("--verbose", is_flag=True, default=False, help="Show detailed diagnostics.")
@click.option("--save", "save_report", is_flag=True, default=False, help="Save report to compatibility/ directory.")
def compatibility_test(config_path: str, json_output: bool, verbose: bool, save_report: bool) -> None:
    """Run deep compatibility validation against the SAP environment.

    Tests system, SAP GUI, session, screen, fields, export, permissions,
    timing, and theme. Never modifies SAP or executes business logic.

    Exit codes:
        0  Compatible
        1  Compatible with warnings
        2  Not compatible
    """
    from .core.compatibility import run_all
    from .core.compatibility import save_report as compat_save

    report = run_all(config_path)

    if json_output:
        click.echo(json.dumps(report.to_dict(), indent=2, default=str))
    else:
        click.echo(report.format_console(verbose=verbose))

    if save_report:
        json_path, txt_path = compat_save(report)
        click.echo("\nReport saved:")
        click.echo(f"  {json_path}")
        click.echo(f"  {txt_path}")

    raise SystemExit(report.exit_code)


@main.command()
@click.option("-c", "--config", "config_path", default="config/default.yaml", help="Path to config YAML file.")
@click.option("-o", "--output", "output_path", default="support_bundle.zip", help="Output ZIP file path.")
@click.option("--json", "json_output", is_flag=True, default=False, help="Output as machine-readable JSON.")
@click.option("--verbose", is_flag=True, default=False, help="Include verbose diagnostics.")
def support(config_path: str, output_path: str, json_output: bool, verbose: bool) -> None:
    """Create a complete support bundle for troubleshooting.

    Collects system info, SAP diagnostics, logs, config, and runs
    doctor/precheck/compatibility-test. All secrets are automatically
    redacted. Never modifies SAP.
    """
    from .core.support import create_support_bundle, create_support_bundle_json

    if json_output:
        data = create_support_bundle_json(config_path)
        click.echo(json.dumps(data, indent=2, default=str))
        return

    zip_path = create_support_bundle(
        config_path=config_path,
        output_path=output_path,
        verbose=verbose,
    )

    click.echo("")
    click.echo("=" * 60)
    click.echo("  ✓ Support Bundle Created")
    click.echo("")
    click.echo(f"  {zip_path.resolve()}")
    click.echo("")
    click.echo("  Ready to send to support.")
    click.echo("=" * 60)
