# Healthcare Enrollment Reconciliation Tool

A lightweight Streamlit application for comparing a Vendor/Client enrollment file with a Benefit Administration (BenAdmin) extract. It creates an Excel report with `Match`, `Changes`, and `Summary` worksheets.

**Demo safety:** This project is intended for synthetic data. Do not upload real SSNs, dates of birth, or employee benefit data to a public deployment. A production deployment requires organization-approved hosting and security review.

## Features

- Upload `.xlsx`, `.xls` (requires `xlrd`), and `.csv` files.
- Select worksheets for Excel files.
- Heuristically suggest matching keys, with explicit user confirmation.
- Map equivalent columns with different headers.
- Normalize whitespace, common dates, and numeric values through user settings.
- Detect duplicate keys, missing keys, and records missing from either source.
- Compare selected fields and generate actionable comments.
- Run starter data-quality checks, including Active status with termination date, effective date before eligibility date, and an HSA outlier warning.
- Download a formatted Excel workbook.
- No intentional long-term storage of uploads or reports.

## Local setup

Python 3.10+ is recommended.

```bash
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL printed by Streamlit, usually `http://localhost:8501`.

## Run tests

```bash
pip install -r requirements-dev.txt
pytest -q
```

## Use the sample data

The `sample_data/` folder contains synthetic vendor and BenAdmin sample files plus a scenario catalog, if included. Upload both files, select the suggested Employer ID + Employee ID key if offered, review the mappings, and compare.

The data is fabricated for QA. It must not be interpreted as real employee data or as authoritative benefits/tax guidance.

## Deploy a public synthetic-data demo

One common route is Streamlit Community Cloud:

1. Create a GitHub repository and push this project.
2. Ensure `app.py`, `reconciliation_engine.py`, and `requirements.txt` are at the repository root.
3. Open Streamlit Community Cloud and connect your GitHub account.
4. Create a new app, select the repository, branch, and `app.py`.
5. Deploy and wait for the platform to build the dependencies.
6. Test the resulting URL with synthetic files only.

Deployment platform names, limits, availability, and free-tier terms may change. This repository does not contain deployment credentials and does not claim that a live URL has been published.

## Security and data retention

- This application does not intentionally write uploaded files or generated reports to a persistent directory.
- Streamlit session memory and hosting infrastructure can retain transient data while a session is active; this is not a guarantee of secure deletion from all infrastructure.
- Use only synthetic data on a public demo.
- Do not claim HIPAA compliance based solely on this code.
- Before using real employee data, obtain approval for the hosting environment, authentication, access controls, encryption, logging, retention/deletion, incident response, and applicable contractual requirements.
- Avoid logging source row values or SSNs. Do not share a public app with real employee data.
- For production, consider private hosting behind organization identity/access controls and a formal security review.

## Important comparison behavior

- Suggested keys are heuristic recommendations, not proof of identity. The user must verify the key selection.
- Duplicate keys are flagged and excluded from automatic one-to-one matching.
- The Changes sheet is long-form: it creates one row per field/rule exception and may include unmapped-column warnings.
- `Match` means no differences were found in mapped, comparable fields and enabled starter rules. It does not certify that enrollment is substantively correct.
- Empty plan fields are compared as values; they are not automatically interpreted as a waiver.
- Starter rules are examples only. Waiting periods, plan-year limits, eligibility, retroactive coverage, and termination conventions vary by client and plan.
- For production-grade reconciliation, adapt matching rules to each client’s data contract and validate with approved expected-results test cases.

## Project files

- `app.py`: Streamlit user interface.
- `reconciliation_engine.py`: File loading, key suggestions, comparison, validation, and workbook generation.
- `tests/`: Automated tests.
- `requirements.txt`: Runtime dependencies.
- `requirements-dev.txt`: Test dependencies.
