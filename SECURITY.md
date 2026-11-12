# Security Policy

## Supported Versions

Only the latest release of ExcelClone is currently supported with security updates.

| Version | Supported          |
| ------- | ------------------ |
| 1.3.x   | :white_check_mark: |
| < 1.3   | :x:                |

## Reporting a Vulnerability

If you discover a security vulnerability within ExcelClone, please follow these guidelines:

1. **Do not open a public GitHub issue.** Disclosing vulnerabilities publicly can put users at risk before a fix is released.
2. Email your findings directly to the repository maintainer or project security contact: `security@excelclone-app.com` (or submit a private vulnerability report via GitHub Security Advisories if enabled).
3. Include detailed information to help reproduce the issue:
   - Type of vulnerability.
   - Operating system and environment details.
   - Exact steps or proof-of-concept files.
   - Any proposed remediation or mitigation.

## Security Architecture

- **Standalone Execution:** ExcelClone operates entirely on your local machine without mandatory network communication or accounts.
- **VBA Macro Isolation:** ExcelClone preserves `.xlsm` macro storage for roundtrip file compatibility, but does not execute untrusted VBA macro scripts automatically without OS-level sandboxing.
- **Safe Parsing:** OpenXML (.xlsx) files are parsed using verified document format boundaries.
