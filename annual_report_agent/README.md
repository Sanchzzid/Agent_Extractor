# Annual Report Extraction Agent

This agent replaces the manual extraction process. You upload one company's annual report PDFs, one per year, and get back a single Excel workbook covering all the years. It is built with LangGraph and uses the OpenAI-compatible LLM configured in `.env`.

## Setup (first time only)

You need Python 3.12 and Node 22. Run these commands in PowerShell from the `annual_report_agent` folder:

```powershell
python -m venv .venv
.venv\Scripts\pip.exe install -r backend\requirements.txt
cd frontend; npm install; npm run build; cd ..
```

Copy `.env.example` to `.env` in this folder and fill in three lines:

```
LLM_BASE_URL=https://your-endpoint/v1
LLM_MODEL=your-model-name
LLM_API_KEY=...
```

One model does all the work, including the verification pass. Any OpenAI-compatible API works (OpenAI, Gemini, OpenRouter, vLLM, Ollama and others), so to switch provider you only change these three lines. Everything else in `.env.example` is optional. The older `PUBLICA_LLM_*` names are still accepted, and a `.env` in the folder above is used if this folder has none.

## Run the app locally (no Docker)

```powershell
cd backend
..\.venv\Scripts\python.exe run.py
```

Open **http://localhost:8031**. Enter the company name, upload the PDFs, click **Run extraction** and download the Excel file.

`run.py` takes these options:

| Option | Effect |
|---|---|
| `--port 9000` | Use a different port (the default is 8031) |
| `--reload` | Restart automatically when the code changes |

If you change the frontend code, run `npm run build` in `frontend`, then restart `run.py`.

### Frontend dev server (only when editing the UI)

With `run.py` running, start the dev server in a second terminal. It sends API calls to the backend set by `VITE_DEV_API` in `frontend\.env` (for example `VITE_DEV_API=http://localhost:8034`; copy `.env.example`):

```powershell
cd frontend
npm run dev
```

Then open http://localhost:5173.

## Run the backend with Docker

```powershell
docker compose up -d --build
```

The API is then at http://localhost:8024. You can change the port with `BACKEND_PORT` in `.env`.

## Deploy the frontend (Firebase Hosting)

```powershell
cd frontend
# set VITE_API_URL in frontend\.env to the public backend URL
npm run build
firebase deploy
```

## Tests

| Test | Command |
|---|---|
| Offline tests (no LLM calls) | `.venv\Scripts\pytest.exe backend\tests -q` |
| Recreate the sample PDFs | `.venv\Scripts\python.exe backend\tests\make_sample_reports.py` |
| End-to-end check against a running app (real LLM) | `.venv\Scripts\python.exe backend\tests\e2e_check.py http://localhost:8031` |

The end-to-end check uploads the three sample PDFs in `backend\tests\fixtures\` and compares 16 known values. It should finish with `0 failure(s)`.

## Output

The workbook has one sheet for each sheet in the spec, with all years merged. It also has three audit sheets:

| Sheet | Contents |
|---|---|
| `_RUN_SUMMARY` | What was processed |
| `_VERIFICATION_LOG` | What the verifier LLM checked and changed |
| `_WARNINGS` | Missing years, duplicate years and failed reports |

Each row's `verification_status` is `VERIFIED`, `CORRECTED` or `NEEDS_REVIEW`.

All prompts are in `backend/app/prompts.py`.
