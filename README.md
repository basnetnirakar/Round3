Antibody Data Manager (Streamlit)

Files included

- app.py — Streamlit application (single-file app)
- Round3_2.xlsx — Workbook required by the app (must be present next to app.py)
- requirements.txt — Python dependencies to install

Quick start (macOS / Linux)

1. Create a virtual environment and activate it:

   python -m venv .venv
   source .venv/bin/activate

2. Install dependencies:

   pip install -r requirements.txt

3. Run the app:

   streamlit run app.py

4. Open the URL printed by Streamlit (usually http://localhost:8501)

Notes

- This app expects the workbook file `Round3_2.xlsx` to be located in the same folder as `app.py`.
- If you don't have `streamlit-aggrid` or `pillow` installed, some interactive features or image resizing may fall back to simpler behavior.
- If the workbook contains sensitive data, share the zip only with intended recipients.

Troubleshooting

- If Streamlit errors on start, confirm `Round3_2.xlsx` exists and that you installed the packages in the active environment.
- If images are missing, ensure the workbook contains in-cell (rich-value) images; the app extracts these from the XLSX package internals.
