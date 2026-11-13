# Annual Report Vision Agent - VS Code Setup Guide

An AI-assisted Streamlit application that reads 1 to 6 annual reports for **one company**, extracts only the **current financial year** from each report according to an uploaded sheet specification, and produces one consolidated Excel workbook plus a separate audit workbook.

This guide is written for **Windows / VS Code / PowerShell**. macOS and Linux alternatives appear below.

## 1. Prerequisites

Install the following on your own computer:

- **Python 3.11 or 3.12**, with the Windows Python Launcher (`py`) available.
- **Visual Studio Code** with Microsoft's **Python** and **Python Debugger** extensions.
- Internet access for installing dependencies and making your chosen model API calls.
- An API key from **OpenAI, Gemini, or Anthropic**; billing and model availability depend on your provider.

In VS Code, open **Terminal > New Terminal** and check:

```powershell
py --version
```

If Python is not found, install it and restart VS Code. This code requires Python 3.10+, but 3.11 or 3.12 is the recommended setup for current Streamlit releases.

## 2. Open the project

1. Unzip the project bundle into a folder such as `Documents\Annual_Report_Vision_Agent`.
2. In VS Code, select **File > Open Folder** and open that folder. Open the folder, **not** just the `.py` file.
3. Verify that `Annual_Report_Vision_Agent.py`, `requirements.txt` and `SCHEMA_SPEC.txt` are at the root of the folder.

If you already have a Git repository, clone and open it, then copy the project files into that repository root instead:

```powershell
git clone YOUR_REPOSITORY_URL
cd YOUR_REPOSITORY_FOLDER
code .
```

## 3. Create the Python virtual environment

From VS Code's PowerShell terminal in the project folder:

```powershell
py -3.11 -m venv .venv
```

If you installed Python 3.12 rather than 3.11, substitute `py -3.12`. Then press **Ctrl+Shift+P** in VS Code, run **Python: Select Interpreter**, and choose the interpreter inside `.venv`.

The next commands use the virtual environment's Python *directly*, so they work even if PowerShell blocks `Activate.ps1`:

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Check that Streamlit installed:

```powershell
.\.venv\Scripts\python.exe -m streamlit --version
```

## 4. Configure your AI provider securely

Create your private `.env` from the included template:

```powershell
Copy-Item .env.example .env
```

Edit `.env` in VS Code. Example for **OpenAI** (substitute your actual available model and secret key):

```dotenv
VISION_PROVIDER=openai
VISION_MODEL=YOUR_COMPATIBLE_VISION_MODEL_ID
OPENAI_API_KEY=YOUR_PRIVATE_API_KEY
```

For **Gemini**, use `VISION_PROVIDER=gemini` and `GEMINI_API_KEY=...`. For **Anthropic**, use `VISION_PROVIDER=anthropic` and `ANTHROPIC_API_KEY=...`. `VISION_MODEL` must name an accessible model that can accept images, return JSON, and handle the current SDK settings. **Do not paste your secret key into source code or commit `.env` to Git.**

Optional: set `VISION_INPUT_USD_PER_M` and `VISION_OUTPUT_USD_PER_M` after verifying the provider's current USD prices per million tokens. Without these, the app cannot compute a USD estimate; it can still show usage tokens.

## 5. Verify source files and Python syntax before API use

In the VS Code terminal:

```powershell
.\.venv\Scripts\python.exe -m py_compile Annual_Report_Vision_Agent.py
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

These checks do **not** call the AI API. The included tests verify the original `SCHEMA_SPEC.txt`: **41 sheets, 937 columns, 5 intentionally header-only sheets**, and preserved employee turnover column headings.

## 6. Run the web app in VS Code

```powershell
.\.venv\Scripts\python.exe -m streamlit run Annual_Report_Vision_Agent.py
```

Streamlit will normally open a browser at `http://localhost:8501`. If that port is already in use, use the Local URL displayed in the terminal. To stop the app, press **Ctrl+C** in the terminal.

Alternatively, after selecting the `.venv` interpreter, open VS Code's **Run and Debug** view and choose **Annual Report Vision Agent (Streamlit)**. The included `.vscode/launch.json` supplies the debug configuration.

**Do not run this as `python Annual_Report_Vision_Agent.py` or with VS Code's ordinary Run Python File button**; the app's UI needs Streamlit's runtime.

## 7. Complete the Streamlit workflow

### Step 1 - Company and inputs

In the sidebar, enter the company name and static company ID. Rating cycle year and source code are optional. Upload the included `SCHEMA_SPEC.txt`, or paste another numbered sheet specification. Upload **1 to 6 PDF/XLSX reports for the same company**, or enter direct PDF URLs. The application does **not** support `.docx`.

### Step 2 - Confirm sheet spec

Inspect the parsed sheet and header counts and tick **I have checked and confirm the parsed sheet names and exact columns**. The five sheets explicitly marked `(EMPTY)` are created with headers only and are not sent to the AI.

### Step 3 - Profile and confirm reports

Click **Profile Reports**. This may make AI API calls and incur provider costs. Review the editable report table:

- Confirm the detected company, report title and the report's **current** financial year.
- Check the fiscal year end and statement-specific currency/scale.
- Replace every `UNKNOWN` financial scale based on the source report; do not guess it.
- Keep only **one report per year**; untick any misidentified company or duplicate year.
- A missing year generates an audit warning, whereas duplicate years block extraction.

### Step 4 - Extract and review

Review the rough token estimate and click **Extract All Reports**. The app extracts each non-empty sheet, consolidates by financial year, then displays QA results and per-sheet previews. Download **both** workbooks:

1. **MAIN workbook** - exactly the sheets and columns in your spec.
2. **AUDIT workbook** - evidence, conversion, QA checks and summary.

Do not import the results into a production database until the audit workbook has been reviewed and material values and page references have been confirmed manually.

### Append mode

Use **Optional existing output .xlsx (APPEND MODE)** only with a new spec containing **new sheet names**. The agent will not overwrite a sheet that already exists in the uploaded workbook.

## 8. First practical pilot

Begin with **one annual report** and a small sheet specification, such as `COMPANY_PROFILES` and `FINANCIALS`. Confirm that the current-year numbers, report scaling, and page citations are correct. Then move to the full 41-sheet spec and up to 6 reports. More pages and sheets mean more paid model calls.

A short, valid pilot spec might be:

```text
1. COMPANY_PROFILES: static_company_id, financial_year, revenue, profit_after_tax, total_assets, source_location
2. FINANCIALS: static_company_id, financial_year, currency, revenue, profit_before_tax, profit_after_tax, total_assets, total_liabilities, total_equity, source_location
```

## 9. Troubleshooting

| Symptom | Action |
| --- | --- |
| `py` not recognized | Install Python and the Windows launcher or use `python` in place of `py`. Restart VS Code. |
| `No module named streamlit` | Install requirements into `.venv` and run the app using `.venv\Scripts\python.exe`. |
| Provider/API key warning | Confirm `.env` is in the same directory as the Python script; verify `VISION_PROVIDER`, `VISION_MODEL` and its matching API key; restart the app. |
| Model rejects `temperature`, JSON mode or image input | Select a compatible model or revise the provider adapter after checking the current SDK/API requirements. |
| Cannot extract: scale UNKNOWN | Open the report and set the explicit statement scale in the confirmation table. |
| Duplicate financial year | Deselect extra reports for that year before extracting. |
| API rate limit or quota error | Confirm quota and billing with your provider; reduce initial workload, or lower `MAX_WORKERS` in the script. |
| PDF link rejected | Download the report manually and upload the `.pdf`; only direct links returning PDFs are accepted. |
| Poor or missing extraction | Inspect the audit workbook, citation pages and source quality; do not silently substitute inferred figures. |
| File already exists in append workbook | Use a spec with **only new** sheet names, or omit the existing workbook for a fresh output. |

## 10. Publish or update the repo

If you use Git from VS Code, make sure `.env`, `.venv`, cached model responses, and report PDFs/XLSX files remain ignored. Then use the **Source Control** panel to stage and commit, or run:

```powershell
git status
git add .
git commit -m "Add VS Code Streamlit annual report extraction agent"
git push
```

If the repository has no remote or no upstream branch, configure that first. Never push real API keys or confidential annual reports into a public repository.

## macOS / Linux equivalents

```bash
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
cp .env.example .env
.venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m streamlit run Annual_Report_Vision_Agent.py
```

## Project tree

```text
Annual_Report_Vision_Agent/
|-- Annual_Report_Vision_Agent.py
|-- requirements.txt
|-- SCHEMA_SPEC.txt
|-- .env.example
|-- .gitignore
|-- README.md
|-- .vscode/
|   `-- launch.json
|-- tests/
|   `-- test_schema.py
`-- docs/
    `-- AGENT_PROMPT.txt
```

At runtime, the program may create `.annual_report_vision_cache/` and `temp_images/`; these are not files to commit. Source data remains private on the local machine unless the configured API provider is called for extraction.
