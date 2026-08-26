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

If you use the `anon` key, enable access with these policies (the `service_role` key bypasses RLS):

```sql
alter table companies enable row level security;
create policy "allow company reads" on companies for select to anon using (true);
create policy "allow company writes" on companies for insert to anon with check (true);
create policy "allow company updates" on companies for update to anon using (true) with check (true);
create policy "allow company deletes" on companies for delete to anon using (true);
```

Also confirm the table is named exactly `companies` and has exactly these columns: `name` (text primary key) and `data` (jsonb).

In the Streamlit app settings, add these secrets:

```toml
SUPABASE_URL = "https://your-project.supabase.co"
SUPABASE_KEY = "your-server-side-key"
```

`SUPABASE_URL` must be the Supabase Project URL from **Project Settings -> API**. Do not use the Supabase dashboard URL. `SUPABASE_KEY` should be the server-side `service_role` key from that same page. After changing secrets, reboot the Streamlit app from **Manage app**.

Use a server-side Supabase key because this app does not have user authentication. Keep the key in Streamlit Secrets and never commit it to the repository. Once both secrets are present, the app reads and writes company profiles in Supabase; without them it continues to use local JSON files.

DOCX downloads work with the Python dependencies in `requirements.txt`. PDF downloads convert the generated DOCX with LibreOffice so the PDF keeps the same format as the Word document.

Existing profiles can be edited from the company list. The completed certificate can be downloaded as Word, PDF, or a text summary, and the summary can be shared through WhatsApp, Telegram, or email. Attach the downloaded Word or PDF file separately when sending it through a messaging service.

Each company profile includes a `Kepada` choice for Melaka or Selangor. Existing profiles without this field use Melaka by default. Each choice inserts its corresponding recipient address into the certificate.

For Streamlit deployments, `packages.txt` installs LibreOffice automatically.
For a local Ubuntu installation, run:

```bash
sudo apt install libreoffice
```