# Reservist Family Wellbeing Dashboard

Streamlit-based RTL Hebrew dashboard for analyzing wellbeing questionnaires
filled out by families of reservists.

## Run locally

```
pip install -r requirements.txt
streamlit run app.py
```

On Windows, double-click `run.bat` — it kills any old Streamlit on port 8501,
installs deps if needed, and opens the browser automatically.

## Data

Default data file: `data.xlsx` in the repo root.

A new file can be uploaded at runtime via the sidebar uploader; the
upload lives only for the current browser session. To make a new file
the permanent default, replace `data.xlsx` and push.

## Deploy

Hosted on Streamlit Community Cloud (private — viewer allowlist).
A push to `main` triggers a redeploy within ~30 seconds.
