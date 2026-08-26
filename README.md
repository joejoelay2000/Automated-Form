# Automated-Form

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app uses `sample.docx` as its built-in template. For local development, company profiles are saved as JSON files in `companies/`.

## Persistent company storage on Streamlit Community Cloud

The filesystem of a deployed Streamlit app is temporary. To retain company profiles across app restarts, create a Supabase project and run this SQL in its SQL editor:

```sql
create table companies (
	name text primary key,
	data jsonb not null
);
```

In the Streamlit app settings, add these secrets:

```toml
SUPABASE_URL = "https://your-project.supabase.co"
SUPABASE_KEY = "your-server-side-key"
```

Use a server-side Supabase key because this app does not have user authentication. Keep the key in Streamlit Secrets and never commit it to the repository. Once both secrets are present, the app reads and writes company profiles in Supabase; without them it continues to use local JSON files.

DOCX downloads work with the Python dependencies in `requirements.txt`. PDF downloads convert the generated DOCX with LibreOffice so the PDF keeps the same format as the Word document.

For Streamlit deployments, `packages.txt` installs LibreOffice automatically.
For a local Ubuntu installation, run:

```bash
sudo apt install libreoffice
```