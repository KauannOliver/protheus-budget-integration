# Protheus Budget Integration

A desktop workflow for validating budget workbooks and sending organized records to Protheus APIs.

**Python, Requests, openpyxl, Tkinter, Windows Excel COM, TOTVS Protheus**

## What it does

- Reads and validates workbook structures and required headers.
- Separates budget and supplementary flows with backups and execution logs.
- Sends records to configured APIs and retains response information.

## Workflow

```mermaid
flowchart LR
  N0["Workbook"] --> N1
  N1["Structure validation"] --> N2
  N2["Record organization"] --> N3
  N3["Protheus API"] --> N4
  N4["Response and traceability"]
```

## Local setup and use

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```
Python 3.11+ is recommended. Tkinter requires a desktop Python installation.
Windows and Microsoft Excel are required for COM-based workbook operations. Set the workbook paths, output directory, endpoint URLs and credentials using the variables in `.env.example`. This script reads process environment variables; it does not automatically load `.env`.

## Configuration

Environment names declared by the included examples (values must be supplied locally):

`PROTHEUS_BACKUP_DIR`, `PROTHEUS_BUDGET_URL`, `PROTHEUS_EXTRA_URL`, `PROTHEUS_EXTRA_WORKBOOK`, `PROTHEUS_PASSWORD`, `PROTHEUS_USERNAME`, `PROTHEUS_WORKBOOK`.

## Boundaries and limitations

Endpoints, schemas and permissions are installation-specific. No live request was made during publication review. Never test against a production budget without an authorized validation process.

This is a source snapshot, not a claim of production readiness. Business results, dates and unmeasured performance are intentionally not asserted. Application UI language is preserved.

## License

No new distribution license is granted for this corporate source edition. Public visibility does not grant reuse rights; obtain permission from the relevant rights holder. Existing third-party notices remain applicable.

## Verification record

See [snapshot checks and remaining limitations](docs/VERIFICATION.md).
