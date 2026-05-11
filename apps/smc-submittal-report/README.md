# NSW Ports Weekly Hold Point Report

Generates a formatted Word document from a Procore submittal log CSV export.

## Usage

### Streamlit (hosted)
Upload the CSV at the Streamlit Community Cloud URL — report downloads automatically.

### CLI
```bash
pip install -r requirements.txt
python src/generate_weekly_report.py SubmittalLog.csv
# Output: output/NSW_Ports_Weekly_Report_YYYYMMDD.docx
```

## Folder layout
```
app.py               Streamlit entry point
src/
  generate_weekly_report.py   data analysis + Word generation
experiments/         throwaway scripts (committed)
output/              generated .docx files (gitignored)
```

## Procore export
1. Submittals → export all to CSV
2. Save as `SubmittalLog.csv`
3. Upload to Streamlit or place alongside the script for CLI

## Deployment (Streamlit Community Cloud)
- Connect repo at share.streamlit.io
- Main file path: `apps/smc-submittal-report/app.py`
- No secrets required
