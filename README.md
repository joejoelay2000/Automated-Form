# Automated-Form

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app uses `sample.docx` as its built-in template. Company profiles are saved as JSON files in `companies/` and are available the next time the app starts.

DOCX downloads work with the Python dependencies in `requirements.txt`. PDF downloads convert the generated DOCX with LibreOffice so the PDF keeps the same format as the Word document.

For Streamlit deployments, `packages.txt` installs LibreOffice automatically.
For a local Ubuntu installation, run:

```bash
sudo apt install libreoffice
```