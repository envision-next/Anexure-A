# Annexure A Generator — Envision Next

A local web app that generates **Annexure A – Disclosure of Interest in Other Real Estate Organizations**
Word documents from a MahaRERA project spreadsheet.

## How it works

1. Search an **organisation / promoter name**.
2. The app finds all of that organisation's **registered projects** (RERA numbers, address, completion date, complaint number).
3. Enter the **DIN/DPIN** and untick any project row you don't want.
4. Watch the **live preview**, then click **Generate Word** to download a filled `.docx`
   (date auto-set to today, "Place" line removed).

## Run

```bash
pip install flask openpyxl python-docx
python app.py
# open http://127.0.0.1:5000   (or double-click run_annexure.bat on Windows)
```

## Data

- Drop an `.xlsx` in this folder; the app auto-detects columns by header keyword.
  Expected columns: `MahaRERA Number`, `Name of Organisation`,
  `Promoter Official Communication Address`, `Complaint Number`, `Completion Date`.
- With no Excel present, built-in sample data is used.
- `Warrants` and `Revoked` are not in the sheet and default to **No**.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask server + JSON API |
| `data_source.py` | Reads the Excel, auto-detects columns, groups projects by organisation |
| `docx_fill.py` | Fills the Word template in memory |
| `templates/index.html` | Web UI (dark intro splash → light app, live preview) |
| `static/` | Envision Next logos + favicon |
| `run_annexure.bat` | Windows launcher |
