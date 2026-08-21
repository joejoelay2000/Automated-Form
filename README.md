# Automated-Form

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app uses `sample.docx` as its built-in template. Company profiles are saved as JSON files in `companies/` and are available the next time the app starts.

DOCX and PDF downloads work with the Python dependencies in `requirements.txt`. If LibreOffice is available, it is used for higher-fidelity PDF conversion. Otherwise, the app creates a readable PDF directly with PyMuPDF.