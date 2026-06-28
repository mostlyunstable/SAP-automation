# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-06-28

### Added
- Full CLI interface (`sap-auto run`, `validate`, `precheck`, `list-transactions`, `init`)
- YAML-based configuration with environment variable resolution
- COM-based SAP GUI session attachment (no image recognition)
- Configurable field IDs for different SAP themes/versions
- Batch Excel input processing (single-column and full-data-row patterns)
- Multi-format report export (Excel, CSV, PDF) with atomic writes
- Automatic retry with exponential backoff
- Structured logging with secret redaction and execution context
- 12 typed exception classes for granular error handling
- Pre-flight environment validation (11 checks)
- Dry-run mode for safe validation without SAP data modification
- Popup categorization and handling (authorization, warnings, session timeout)
- VA23 (Display Quotation) transaction handler as reference implementation
- 143 unit tests with full type checking (mypy) and lint (ruff)
- GitHub Actions CI/CD pipeline (lint, test, typecheck, release)

### Security
- CSV injection prevention (formula prefix neutralization)
- Secret values redacted from all log output
- Environment variable substitution for credentials (no hardcoded secrets)
- Input validation (transaction code sanitization, SQL injection prevention)
