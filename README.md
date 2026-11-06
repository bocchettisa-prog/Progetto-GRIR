# GR/IR Control Center — Web Edition

Web adaptation of the personal GR/IR reconciliation tool by Salvatore Bocchetti.

## Public prototype / data warnings

- The Streamlit application **is public** if you deploy it with public visibility.
- Uploads are transmitted to and processed by the **third-party hosting server**. **DO NOT upload confidential SAP extracts, real company financial records, personal data, or other material without authorization.** Use synthetic or anonymized files for the public demo.
- The server has **no shared SQLite database**. Each browser session has independent workbooks and edits in memory; nothing persists across restarts and users. Download the reconciliation before leaving. This is not intended as a production multiuser solution.
- Do not upload .xlsx files to GitHub. `.gitignore` protects local Git commits but **GitHub website uploads may still include selected files**; inspect files before committing.
- Static/public GitHub source code must never include real company reports. The app can be deployed from a private repository while configured public in Streamlit's app sharing settings, subject to host policy.

## Deploy for free

1. Create a GitHub repository and upload only the source files, including `app.py`, `requirements.txt`, the `core` folder and `.streamlit/config.toml`.
2. Go to https://share.streamlit.io/ and sign in with GitHub.
3. Select **Create app** / **I have an app**.
4. Choose the repository, branch `main`, and entrypoint `app.py`.
5. Press **Deploy**. In app settings > Sharing choose public access if needed.
6. Share the resulting `*.streamlit.app` URL.

## Local test

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Input naming

`GB12 GRIR 09.26.xlsx` (previous workbook: SAP + REC)
`GB12 GRIR 10.26.xlsx` (current workbook: SAP only)

Importer retains original signed SAP subtotals, validates against the SAP grand total, and uses PO number + item line. Export preserves the official REC layout and stores Resolved / zero items in Reconciliation Status only.

## Differences from Windows V12

The interface is rebuilt with Streamlit; it does not embed the Windows PySide6 widgets.
The public prototype's history is session-only. It does **not** sync with your Windows SQLite database and does not preserve data between visits. Multi-entity analytics displays snapshots currently loaded in the same session only. If you need a persistent multi-user public site later, design authentication, per-user databases, retention, and authorization first.
