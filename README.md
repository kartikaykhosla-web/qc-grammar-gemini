# QC Grammar Gemini

Streamlit app for English editorial QC using Gemini on Vertex AI.

## Main app

- `qc_code_ssg.py`

## Supporting files

- `requirements.txt`
- `packages.txt`
- `runtime.txt`

## Local run

```bash
streamlit run qc_code_ssg.py
```

## Required secret

Add `GCP_SERVICE_ACCOUNT_JSON_B64` in Streamlit secrets.

## Keep-awake workflow

This repo includes a GitHub Actions workflow at `.github/workflows/keep_streamlit_awake.yml`
that pings the deployed app every 6 hours.

Set a repository variable named `STREAMLIT_APP_URL` to your deployed app URL, for example:

```text
https://your-app-name.streamlit.app
```
