# FA Portal — Factory Audit Report Generator

Streamlit app that matches factory QMS documents to FA checklist sections and fills the official TÜV Rheinland FA Word template.

## Quick start (local)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run fa_portal/app.py
```

Open http://localhost:8501

## Backend template (shared for all users)

The FA Word template lives in the repo — **not** on each user's Desktop:

```
FA/templates/FA_template.docx
```

Every deploy uses this file automatically. To update the template, replace that file and redeploy.

Optional env overrides:

| Variable | Purpose |
|----------|---------|
| `FA_TEMPLATE_PATH` | Absolute/relative path to template `.docx` |
| `FA_DATA_DIR` | Writable folder for settings + feedback |
| `FA_RUNS_DIR` | Writable folder for generated reports |

## Deploy

**Do not use Vercel** — this is a long-running Python/Streamlit app with large file uploads.

### Option A — Streamlit Community Cloud (simplest)

1. Push this repo to GitHub (already: `Shreeya1-pixel/FA`).
2. Go to https://share.streamlit.io → **New app**.
3. Set:
   - **Repository:** `Shreeya1-pixel/FA`
   - **Branch:** `main`
   - **Main file path:** `fa_portal/app.py`
4. Deploy. Share the public URL with your team.

### Option B — Railway

1. New project → Deploy from GitHub repo.
2. Railway reads `railway.toml` / `Procfile`.
3. Set start command if needed:
   ```bash
   streamlit run fa_portal/app.py --server.port=$PORT --server.address=0.0.0.0
   ```

### Option C — Render

1. New **Web Service** from this repo.
2. Render uses `render.yaml`, or set:
   - Build: `pip install -r requirements.txt`
   - Start: `streamlit run fa_portal/app.py --server.port=$PORT --server.address=0.0.0.0`

## How users use it

1. Open the portal URL.
2. Upload factory documents (ZIP and/or files).
3. Optional: audit plan / overrides.
4. Click **Generate FA Report**.
5. Download the filled `.docx`.

No one needs the template on their laptop.

## Project layout

```
FA/
  fa_automation.py          # matcher + template filler
  templates/FA_template.docx # shared backend template
  signatures/               # optional auditor signature images
fa_portal/
  app.py                    # Streamlit UI
requirements.txt
Procfile / railway.toml / render.yaml
.streamlit/config.toml
```

## Notes

- Upload limit is set to **500 MB** in `.streamlit/config.toml`.
- Generated runs are stored under `fa_portal/data/runs` locally, or `/tmp/fa_portal_data` on read-only hosts.
- Local Ollama / Qwen diagnosis is optional and only works where Ollama is installed (not on Streamlit Cloud by default).
- Do **not** add an empty/`packages.txt` with only comments — Streamlit Cloud apt install fails with `Unsupported file / given on commandline`.
