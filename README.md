# SAP Automation Framework

Enterprise-grade SAP GUI automation for Windows. Attaches to an existing SAP session, processes batch records from Excel, executes transactions, exports reports, and produces execution summaries.

## Overview

This framework automates repetitive SAP report export workflows. It connects to a running SAP GUI instance via COM, reads input data from Excel files, executes configured transactions against each record, and exports the results.

**Key design decisions:**

- Attaches to an already logged-in SAP session. No credential storage. Works with SSO and MFA.
- Uses SAP GUI Scripting APIs only. No image recognition or coordinate-based interaction.
- Field IDs are configurable. Different SAP themes, versions, and transaction variants can be accommodated without code changes.
- Atomic file writes prevent corrupted output. Export files are verified immediately after creation.

**Who should use this:**

- Teams that manually export SAP reports to Excel on a recurring basis.
- Environments where SAP does not offer a direct API or where GUI scripting is the only available automation path.

**Platform:** Windows only (requires COM automation via `pywin32`).

---

## Features

- SAP GUI COM session attachment with connection filtering by system name and client
- Batch processing from Excel input (single-column document list or full data rows)
- Transaction plugin architecture with abstract base class
- Automatic retry with exponential backoff for transient SAP errors
- Structured logging with secret filtering and execution context
- YAML configuration with environment variable substitution
- Export to Excel, CSV, and PDF formats with atomic writes
- Dry run mode for validation without modifying SAP
- Pre-flight environment checks (11 diagnostics)
- Comprehensive health diagnostics (`doctor` command)
- Deep SAP compatibility testing (`compatibility-test` command)
- Support bundle generation for troubleshooting
- Runtime popup handling with categorization (authorization, warning, information, session timeout)
- Error recovery with typed exception hierarchy (12 exception classes)
- Transaction code injection prevention via regex validation
- CSV formula injection prevention
- Secret redaction in logs and support bundles

---

## Architecture

```
sap_automation/
    __init__.py              Package version
    __main__.py              python -m entry point
    main.py                  Script entry point
    cli.py                   Click CLI (8 commands)
    orchestrator.py          Batch processing, retry, connection lifecycle
    core/
        __init__.py          Re-exports
        base_transaction.py  Abstract base class + TransactionResult
        checks.py            17 shared diagnostic check functions
        compatibility.py     Deep SAP compatibility validation
        config.py            YAML config, env var substitution, validation
        connection.py        SAP GUI COM session management
        doctor.py            Comprehensive environment diagnostics
        exceptions.py        12 typed exception classes
        logger.py            Secret filtering, execution context, rotation
        precheck.py          Pre-flight environment validation
        retry.py             Exponential backoff decorator
        support.py           Support bundle generator with redaction
    transactions/
        va23_display_quotation.py  VA23 handler (reference implementation)
    utils/
        excel_reader.py      Excel input with validation
        exporter.py          Excel/CSV/PDF export with atomic writes
config/
    default.yaml             Default configuration
tests/
    test_*.py                143 tests
```

### Runtime Flow

```
CLI (cli.py)
    |
    v
Configuration (config.py)
    |  Load YAML, resolve env vars, validate
    v
Precheck (precheck.py)
    |  Verify Python, OS, SAP GUI, scripting, packages
    v
Excel Reader (excel_reader.py)
    |  Parse input file, validate records
    v
Orchestrator (orchestrator.py)
    |  Connect once to SAP, iterate records
    v
Transaction Handler (base_transaction.py -> va23_display_quotation.py)
    |  Validate input, open transaction, execute, extract data, post-validate
    v
Exporter (exporter.py)
    |  Write Excel/CSV/PDF, atomic writes, verify output
    v
Summary (RunSummary)
    |  JSON report, console output, exit code
    v
Done
```

---

## Requirements

| Requirement | Version | Notes |
|-------------|---------|-------|
| Python | >= 3.10 | 64-bit recommended |
| Windows | 10/11 or Server 2016+ | COM automation required |
| SAP GUI for Windows | 7.60+ | With scripting enabled |
| Microsoft Excel | Any version | Required for PDF export only |

**SAP GUI Scripting must be enabled:**

1. Open SAP Logon.
2. Go to **Options > Accessibility > GUI Scripting**.
3. Enable **Enable GUI Scripting**.
4. Disable **Notify / Do not notify** to prevent popup interruptions.

**Required Python packages (installed automatically):**

| Package | Purpose |
|---------|---------|
| click >= 8.1 | CLI framework |
| pyyaml >= 6.0 | Configuration parsing |
| openpyxl >= 3.1 | Excel read/write |
| pywin32 >= 306 | Windows COM automation |

---

## Installation

### 1. Clone the repository

```bash
git clone <repository-url>
cd sap-automation
```

### 2. Create a virtual environment

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -e .
```

For development (includes pytest, ruff, mypy):

```bash
pip install -e ".[dev]"
```

### 4. Verify installation

```bash
sap-auto --version
```

Expected output:

```
SAP Automation, version 1.0.0
```

### 5. Initialize directory structure

```bash
sap-auto init
```

This creates `output/`, `logs/`, and `config/` directories.

---

## Configuration

Configuration is stored in `config/default.yaml`. All settings have sensible defaults.

### Configuration Reference

```yaml
sap:
  system_name: "PRD"           # SAP system name (matches SAP Logon)
  client: "100"                # SAP client number
  language: "EN"               # SAP logon language

paths:
  output_dir: "./output"       # Export file destination
  log_dir: "./logs"            # Log file destination
  template_dir: "./config/templates"

export:
  format: "xlsx"               # xlsx | csv | pdf | all
  filename_pattern: "{transaction}_{document_number}_{timestamp}"
  include_timestamp: true

retry:
  max_attempts: 3              # Retry count per record
  delay_seconds: 2.0           # Initial delay between retries
  backoff_multiplier: 2.0      # Delay multiplier after each retry
  max_delay_seconds: 60.0      # Maximum delay cap
  retry_on_exceptions:
    - "SAPBusyError"
    - "SessionNotFoundError"

logging:
  level: "INFO"                # DEBUG | INFO | WARNING | ERROR
  file_rotation: "daily"
  max_log_files: 30
  format: "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"

transactions:
  va23:
    class: "sap_automation.transactions.va23_display_quotation.VA23DisplayQuotation"
    description: "Display and export SAP quotations"
    export_format: "xlsx"

excel:
  document_number_column: "Document Number"
  sheet_name: null             # null = first sheet
```

### Environment Variables

Reference environment variables in config using `${VAR_NAME}` syntax:

```yaml
sap:
  system_name: "${SAP_SYSTEM}"
  client: "${SAP_CLIENT}"
```

If a referenced variable is not set, the framework reports which config field uses it and exits.

---

## Excel Input Format

The framework accepts two input patterns:

### Pattern 1: Document Number List

A single column containing document numbers.

| Document Number |
|-----------------|
| 1000012345 |
| 1000012346 |
| 1000012347 |

### Pattern 2: Full Data Rows

Multiple columns. The column specified by `excel.document_number_column` in config is mapped to the document number.

| Document Number | Material | Plant | Date |
|-----------------|----------|-------|------|
| 1000012345 | MAT-001 | 1000 | 2026-01-15 |
| 1000012346 | MAT-002 | 2000 | 2026-01-16 |

### Validation Rules

| Rule | Detail |
|------|--------|
| File format | `.xlsx` or `.xlsm` only |
| Required column | `Document Number` (configurable) |
| Empty document numbers | Rows with empty document numbers are processed with an empty string |
| Duplicate detection | Warning logged for duplicate document numbers |
| Blank rows | Automatically skipped |
| File locks | Fails fast if the file is open in another application |

---

## CLI Commands

### `sap-auto run`

Execute a transaction for all records in an Excel file.

```bash
sap-auto run -t va23 -i input.xlsx
```

| Flag | Default | Description |
|------|---------|-------------|
| `-c, --config` | `config/default.yaml` | Configuration file path |
| `-t, --transaction` | (required) | Transaction name from config |
| `-i, --input` | (required) | Excel input file path |
| `-o, --output` | config value | Output directory override |
| `--export` | config value | Export format: `xlsx`, `csv`, `pdf`, `all` |
| `--dry-run` | off | Validate without modifying SAP |

Exit codes:
- `0` -- All records processed successfully.
- `1` -- One or more records failed.

### `sap-auto validate`

Validate a configuration file.

```bash
sap-auto validate -c config/default.yaml
```

### `sap-auto precheck`

Run environment compatibility checks. Checks Python version, OS, SAP GUI availability, COM scripting, required packages, and output directories.

```bash
sap-auto precheck
sap-auto precheck --json-output
```

| Flag | Description |
|------|-------------|
| `-c, --config` | Configuration file path (optional) |
| `--json-output` | Output as JSON |

Exit codes:
- `0` -- Environment is compatible.
- `1` -- Critical incompatibilities detected.

### `sap-auto doctor`

Run comprehensive environment diagnostics. Performs read-only health checks of system, Python, SAP, configuration, file system, Office, logging, and security. Never modifies SAP or configuration.

```bash
sap-auto doctor
sap-auto doctor --json
sap-auto doctor --verbose
sap-auto doctor --save
```

| Flag | Description |
|------|-------------|
| `-c, --config` | Configuration file path |
| `--json` | Machine-readable JSON output |
| `--verbose` | Show detailed diagnostics |
| `--save` | Save report to `diagnostics/` directory |

Exit codes:
- `0` -- READY or READY WITH WARNINGS.
- `1` -- NOT READY (critical failures).

### `sap-auto compatibility-test`

Run deep compatibility validation against the SAP environment. Tests system, SAP GUI, session, screen, fields, export, permissions, timing, and theme. Never modifies SAP or executes business logic.

```bash
sap-auto compatibility-test
sap-auto compatibility-test --json
sap-auto compatibility-test --verbose
sap-auto compatibility-test --save
```

| Flag | Description |
|------|-------------|
| `-c, --config` | Configuration file path |
| `--json` | Machine-readable JSON output |
| `--verbose` | Show detailed diagnostics |
| `--save` | Save report to `compatibility/` directory |

Exit codes:
- `0` -- Compatible.
- `1` -- Compatible with warnings.
- `2` -- Not compatible.

### `sap-auto support`

Create a complete support bundle for troubleshooting. Collects system info, SAP diagnostics, logs, config, and runs doctor/precheck/compatibility-test. All secrets are automatically redacted. Never modifies SAP.

```bash
sap-auto support
sap-auto support -o support_bundle.zip
sap-auto support --json
sap-auto support --verbose
```

| Flag | Default | Description |
|------|---------|-------------|
| `-c, --config` | `config/default.yaml` | Configuration file path |
| `-o, --output` | `support_bundle.zip` | Output ZIP file path |
| `--json` | off | Output as JSON (no ZIP) |
| `--verbose` | off | Include verbose diagnostics |

### `sap-auto list-transactions`

List configured transaction handlers.

```bash
sap-auto list-transactions
sap-auto list-transactions -c config/default.yaml
```

### `sap-auto init`

Initialize output directory structure. Creates `output/`, `logs/`, and `config/` directories.

```bash
sap-auto init
sap-auto init /path/to/project
```

---

## Workflow

The automation processes each record through the following steps:

1. **Read Excel** -- Parse the input file, validate document numbers, detect duplicates and empty rows.

2. **Connect to SAP** -- Attach to the running SAP GUI instance. The connection is established once per batch, not per record. Filters connections by `system_name` and `client` from config.

3. **Open Transaction** -- Navigate to the configured transaction code (e.g., `/nVA23`). Waits for SAP to be ready. Handles popups that appear after navigation.

4. **Pre-validate** -- Verify the session is healthy, the screen loaded correctly, and required fields are present.

5. **Execute** -- Enter the document number into the input field. Press Enter. Handle any popups.

6. **Extract Data** -- Read field values from the SAP screen. Each field extraction is independent -- if one field is missing, others are still attempted.

7. **Post-validate** -- Check the SAP status bar for errors. Verify the session is still alive.

8. **Export** -- Write extracted data to the configured format (Excel, CSV, PDF). Files are written atomically using a temp file and `os.replace()` to prevent corruption.

9. **Verify Export** -- Confirm the output file exists, is a file (not directory), and is not empty.

10. **Retry on Failure** -- If the transaction fails with a transient error (`SAPBusyError`, `SessionNotFoundError`), retry with exponential backoff. Validation errors are never retried.

11. **Record Result** -- Log the outcome and add to the batch summary.

12. **Next Record** -- Repeat from step 3 until all records are processed.

---

## Output

### Export Files

Located in `./output/` (configurable). Naming pattern:

```
{transaction}_{document_number}_{timestamp}.{ext}
```

Example: `VA23_1000012345_20260628_143022.xlsx`

### Summary JSON

Saved after every run in the output directory:

```
output/summary_{run_id}.json
```

Contains:

- Run ID, transaction name, dry run flag
- Start/end timestamps
- Total/successful/failed/retried/skipped counts
- Success rate percentage
- Per-record results with exported file paths, error types, and timing
- Environment information (Python version, OS, SAP GUI version)

### Logs

Located in `./logs/`:

- `sap_automation.log` -- Rotated daily, 30-day retention.
- Log level configurable via `logging.level` in config.
- Secret values automatically redacted.
- Each log entry includes execution context (`run_id`) for correlation.

### Diagnostics

Generated by `doctor`, `compatibility-test`, and `support` commands:

- `diagnostics/report.json` and `diagnostics/report.txt` -- Doctor output.
- `compatibility/report.json` and `compatibility/report.txt` -- Compatibility output.
- `support_bundle.zip` -- Complete support package with all diagnostics.

---

## Error Handling

### Retryable Errors

| Exception | Cause | Action |
|-----------|-------|--------|
| `SAPBusyError` | SAP is processing a long-running operation | Wait with exponential backoff, retry |
| `SessionNotFoundError` | SAP session was closed or lost | Wait with exponential backoff, retry |

Retries use exponential backoff: `delay * backoff_multiplier`, capped at `max_delay_seconds`.

### Non-Retryable Errors

| Exception | Cause | Action |
|-----------|-------|--------|
| `ValidationError` | Input record is invalid (missing field, wrong format) | Skip record, log error |
| `TransactionError` | Transaction failed for a specific reason | Skip record, log error with details |
| `FieldError` | A required SAP field could not be found | Skip record, log field ID and suggestion |
| `ExportError` | File write or export failed | Skip record, log file path and suggestion |
| `SAPCancelledError` | User cancelled an SAP operation | Skip record, log cancellation |

### Connection Errors

| Exception | Cause | Action |
|-----------|-------|--------|
| `SAPConnectionError` | Cannot find SAP GUI, cannot access scripting engine | Exit with actionable diagnostics |
| `SessionNotFoundError` | No active SAP session within timeout | Exit with connection guidance |

### Configuration Errors

| Exception | Cause | Action |
|-----------|-------|--------|
| `ConfigError` | Invalid YAML, missing required fields, invalid class paths | Exit with specific field reference |

### Error Reporting

Every exception includes:

- Human-readable message.
- `details` dictionary with additional context.
- `suggestion` field with actionable remediation steps.

---

## Security

| Control | Implementation |
|---------|----------------|
| No credential storage | Framework attaches to existing SAP login. No passwords stored in config or code. |
| Environment variable substitution | Secrets referenced via `${VAR_NAME}` in config. Never hardcoded. |
| Secret redaction | Passwords, tokens, API keys, and IPs automatically redacted from logs and support bundles. |
| CSV injection prevention | Formula characters (`=`, `+`, `-`, `@`, `\t`) at the start of CSV values are prefixed with a single quote. |
| Transaction code validation | Regex validation ensures transaction codes are exactly 4 alphanumeric characters. Prevents OKCode injection. |
| Path validation | Output paths checked against dangerous system directories (`C:\Windows`, `/usr`, etc.). |
| Safe file handling | Atomic writes using temp file + `os.replace()`. No partial files left on disk. |
| Read-only SAP connection | Framework never modifies SAP master data. Only reads transaction screens and exports data. |
| Config secrets scan | `doctor` command scans config for hardcoded passwords, API keys, and tokens. |

---

## Troubleshooting

| Problem | Possible Cause | Solution |
|---------|---------------|----------|
| `Cannot find running SAP GUI` | SAP Logon is not running | Start SAP Logon and log in before running the framework. |
| `Cannot access SAP scripting engine` | GUI Scripting is disabled | Enable: SAP Logon > Options > Accessibility > GUI Scripting > Enable GUI Scripting. |
| `No SAP connections found` | Not logged into any system | Open SAP Logon, select a system, and log in. |
| `No active SAP session found` | Session timed out or was closed | Ensure the SAP session remains open during automation. |
| `Connection matching system=X not logged in` | Wrong system name in config | Check `sap.system_name` matches your SAP Logon connection name. |
| `Connection matching system=X client=Y not logged in` | Wrong client number | Check `sap.client` matches the client shown in SAP Logon. |
| `Export file not created` | Output directory not writable | Check directory permissions. Run `sap-auto init` to create directories. |
| `File not readable (permission denied)` | Excel file is open in another application | Close the file in Excel and retry. |
| `Invalid file type` | Input is not `.xlsx` or `.xlsm` | Convert the input file to Excel format. |
| `pywin32 is required` | `pywin32` not installed | Run: `pip install pywin32`. |
| `VA23 input field not found` | SAP theme or transaction variant changed the screen layout | Run `sap-auto compatibility-test --verbose` to see which fields are missing. Update `transactions.va23.field_ids` in config. |
| `PDF export requires pywin32` | Microsoft Excel not available via COM | Install Microsoft Office, or use `xlsx`/`csv` export format instead. |
| `Environment variable 'X' is not set` | Referenced env var is missing | Set the environment variable before running. |

---

## Client Deployment Guide

### Step 1: Verify Prerequisites

```bash
sap-auto precheck
```

All critical checks must pass. Review any warnings.

### Step 2: Run Full Diagnostics

```bash
sap-auto doctor
sap-auto doctor --verbose --save
```

Review `diagnostics/report.txt` for any issues.

### Step 3: Test SAP Compatibility

```bash
sap-auto compatibility-test
sap-auto compatibility-test --verbose --save
```

This validates session connectivity, field availability, transaction authorization, and screen layout. Review `compatibility/report.txt`.

If fields are reported as missing:

1. Open SAP GUI manually and navigate to the transaction.
2. Enable GUI Scripting trace (SAP Logon > Options > Scripting > Trace).
3. Manually perform the transaction and note the field IDs from the trace.
4. Update `transactions.<name>.field_ids` in `config/default.yaml`.

### Step 4: Dry Run

```bash
sap-auto run -t va23 -i input.xlsx --dry-run
```

Dry run validates configuration, input records, and optionally tests SAP connectivity. No SAP data is modified.

### Step 5: Small Production Batch

```bash
sap-auto run -t va23 -i small_input.xlsx
```

Process a small batch (5-10 records) first. Verify:

- Export files are created in `./output/`.
- Exported data matches what you see in SAP.
- Summary JSON is generated.
- No errors in `./logs/sap_automation.log`.

### Step 6: Full Production Rollout

```bash
sap-auto run -t va23 -i full_input.xlsx
```

Monitor the console output and log files. If errors occur, check the summary JSON for failed documents and reprocess them.

---

## UAT Checklist

Before deploying to production, confirm each item:

- [ ] SAP GUI for Windows is installed (version 7.60+)
- [ ] GUI Scripting is enabled in SAP options
- [ ] User is logged into the correct SAP system and client
- [ ] `config/default.yaml` has correct `system_name` and `client`
- [ ] `sap-auto precheck` passes all critical checks
- [ ] `sap-auto doctor` shows READY or READY WITH WARNINGS
- [ ] `sap-auto compatibility-test` shows COMPATIBLE or COMPATIBLE WITH WARNINGS
- [ ] `sap-auto run --dry-run` validates input without errors
- [ ] `sap-auto compatibility-test --verbose` confirms all field IDs are found
- [ ] Small batch (5-10 records) processes successfully
- [ ] Exported files contain correct data
- [ ] Summary JSON is generated with expected counts
- [ ] No errors in log files
- [ ] Full production batch processes successfully
- [ ] Failed documents are reviewed and reprocessed if needed

---

## Repository Structure

```
sap-automation/
    .github/
        workflows/
            ci.yml               CI pipeline (lint, test, typecheck)
            release.yml          Release packaging
    config/
        default.yaml             Default configuration
    sap_automation/
        __init__.py              Package version (1.0.0)
        __main__.py              python -m entry point
        main.py                  Script entry point
        cli.py                   Click CLI with 8 commands
        orchestrator.py          Batch processing, retry, connection lifecycle
        core/
            base_transaction.py  Abstract base class + TransactionResult
            checks.py            17 shared diagnostic check functions
            compatibility.py     Deep SAP compatibility validation
            config.py            YAML config, env var substitution
            connection.py        SAP GUI COM session management
            doctor.py            Comprehensive health diagnostics
            exceptions.py        12 typed exception classes
            logger.py            Secret filtering, execution context
            precheck.py          Pre-flight environment validation
            retry.py             Exponential backoff decorator
            support.py           Support bundle generator
        transactions/
            va23_display_quotation.py  VA23 reference implementation
        utils/
            excel_reader.py      Excel input with validation
            exporter.py          Excel/CSV/PDF export with atomic writes
    tests/
        test_*.py                143 tests
    scripts/
        create_example_input.py  Sample input generator
    .gitignore
    .pre-commit-config.yaml
    CHANGELOG.md
    LICENSE
    pyproject.toml
    README.md
```

---

## Development

### Adding a New Transaction

1. Create a new file in `sap_automation/transactions/`:

```python
from ..core.base_transaction import BaseTransaction

class MyTransaction(BaseTransaction):
    @property
    def tcode(self) -> str:
        return "MYTC"

    @property
    def description(self) -> str:
        return "My Transaction Description"

    def validate_record(self, record):
        # Validate input fields
        ...

    def execute(self, session, record):
        # Interact with SAP, return extracted data
        ...
```

2. Register in `config/default.yaml`:

```yaml
transactions:
  mytc:
    class: "sap_automation.transactions.my_transaction.MyTransaction"
    description: "My transaction description"
    export_format: "xlsx"
```

3. Run validation:

```bash
sap-auto validate
sap-auto list-transactions
```

### Coding Standards

- Python 3.10+ syntax (`str | None` unions, `match` statements).
- Type hints on all public functions.
- Line length: 100 characters (configured in `pyproject.toml`).
- Lint: `ruff check .`
- Type check: `mypy .`
- No comments unless requested.
- All COM imports are lazy-loaded (Windows-only modules).

### Running Tests

```bash
pytest                    # Run all 143 tests
pytest -x                 # Stop on first failure
pytest tests/test_config.py  # Run specific test file
```

### Linting and Type Checking

```bash
ruff check .              # Lint
ruff format .             # Format
mypy .                    # Type check
```

---

## Testing

The test suite includes 143 tests across 11 test files:

| Test File | Coverage |
|-----------|----------|
| `test_config.py` | Config loading, validation, env vars, dot notation, merge |
| `test_connection.py` | Session attachment, popup handling, tcode validation |
| `test_base_transaction.py` | Transaction lifecycle, export, verification |
| `test_excel_reader.py` | Input parsing, validation, edge cases |
| `test_exporter.py` | Excel/CSV export, atomic writes, injection prevention |
| `test_logger.py` | Secret filtering, execution context, setup |
| `test_retry.py` | Exponential backoff, exception filtering |
| `test_orchestrator.py` | Batch processing, dry run, retry integration |
| `test_exceptions.py` | Exception hierarchy, details, suggestions |
| `test_cli.py` | CLI command invocation and argument handling |
| `test_va23.py` | VA23 handler validation and execution |

All tests pass, ruff reports no issues, and mypy reports no type errors.

---

## Known Limitations

- **Windows only.** The framework requires Windows COM automation (`pywin32`). It does not run on Linux or macOS.
- **SAP GUI Scripting must be enabled.** Some SAP environments disable scripting for security. The framework cannot function without it.
- **Requires an existing SAP login.** The user must be logged into SAP manually before running the framework. No credential-based login is supported.
- **Field IDs may require configuration.** Different SAP GUI themes (Quartz, Belize, etc.), transaction variants, and screen customizations can change field IDs. The `compatibility-test` command identifies missing fields.
- **PDF export requires Microsoft Excel.** PDF export uses Excel COM automation. If Excel is not installed, only xlsx and csv formats are available.
- **No parallel record processing.** Records are processed sequentially to avoid SAP session conflicts.
- **No scheduling.** The framework is a CLI tool. For scheduled execution, use Windows Task Scheduler or a CI/CD pipeline.

---

## License

MIT License. See [LICENSE](LICENSE) for details.

---

## Contributing

1. Fork the repository.
2. Create a feature branch.
3. Make changes with tests.
4. Run `ruff check .`, `mypy .`, `pytest` before committing.
5. Submit a pull request.

All contributions must include tests for new functionality and maintain existing test coverage.

---

## Support

Before contacting support, generate a support bundle:

```bash
sap-auto support -o support_bundle.zip --verbose
```

This ZIP file contains:

- System information (OS, Python, hardware)
- SAP GUI version and connection details
- Framework version and configuration (secrets redacted)
- Latest log files (secrets redacted)
- Doctor, precheck, and compatibility-test results
- Crash information if available

Send the ZIP file along with:

1. A description of the problem.
2. The exact command you ran.
3. Any error messages displayed.
4. Your SAP GUI version and theme.
