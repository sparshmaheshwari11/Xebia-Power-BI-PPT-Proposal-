# Xebia Power BI Proposal Studio — Groq Editorial Edition

A web application that turns Power BI project requirements into a detailed, editable PowerPoint proposal. Users can provide requirements as pasted text **or** upload a BRD in PDF, DOCX, or TXT format. The application uses the Groq API for proposal content and `python-pptx` to generate the PPTX.

> **AI provider:** Groq only. This project does not use Gemini or OpenRouter.

## Features

- Plain-text requirements **or** one uploaded BRD; the two input methods cannot be used together.
- Supported uploads: `.pdf`, `.docx`, `.txt` (maximum size configurable; default 25 MB).
- Optional client name, industry, and proposal tone fields.
- Fixed Xebia capability and accelerator section followed by a client requirements slide and client-specific proposal content.
- Proposal coverage can include architecture, source ingestion, semantic model design, KPI definitions and DAX examples, report-page specifications, security, governance, data quality, performance, CI/CD, delivery phases, risks, acceptance criteria, and next steps.
- Editorial PowerPoint design with narrative slides, native PowerPoint tables, process flows, architecture diagrams, and roadmap timelines.
- PPTX elements are editable in Microsoft PowerPoint.
- Health endpoint: `/api/health`.

## Technology stack

- **Frontend:** HTML, CSS, JavaScript
- **Backend:** Python, FastAPI, Uvicorn
- **AI:** Groq Chat Completions API
- **PowerPoint generation:** `python-pptx`
- **BRD parsing:** `pypdf` and `python-docx`

## Requirements

- Python 3.10+ recommended
- A Groq API key
- Git, if deploying from GitHub
- A Render account, if deploying to Render

## Run locally on Windows / VS Code

### 1. Extract and open the project

Extract the ZIP and open the project root in VS Code. The root should contain `requirements.txt`, `.env.example`, `backend/`, and `frontend/`.

### 2. Create and activate a virtual environment

Run these commands in the VS Code PowerShell terminal from the project root:

```powershell
py -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell activation is blocked, you can run the environment's Python directly with `\.venv\Scripts\python.exe` instead of activating it.

### 3. Create the `.env` file

```powershell
Copy-Item .env.example .env
notepad .env
```

Set the values in `.env`:

```dotenv
GROQ_API_KEY=your_actual_groq_api_key
GROQ_MODEL=openai/gpt-oss-120b
MAX_UPLOAD_MB=25
```

Replace the example key with your own key. Use a model that is currently available to your Groq account. **Never commit `.env` or expose the API key in frontend code.** The repository `.gitignore` excludes `.env`.

### 4. Start the server

```powershell
python -m uvicorn backend.main:app --reload
```

Open [http://127.0.0.1:8000](http://127.0.0.1:8000).

### 5. Verify the health endpoint

Open [http://127.0.0.1:8000/api/health](http://127.0.0.1:8000/api/health). The response should report `"status": "ok"`, `"provider": "Groq"`, the configured model, and whether the key is configured.

## Deploy to Render.com

This application can be deployed as a **Python Web Service** from a GitHub repository.

### 1. Push the project to GitHub

Create a repository, then run these commands from the project root. Replace the repository URL with your own:

```powershell
git init
git add .
git status
git commit -m "Initial deployment of Power BI Proposal Studio"
git branch -M main
git remote add origin https://github.com/YOUR_GITHUB_USERNAME/xebia-powerbi-proposal-studio.git
git push -u origin main
```

Check `git status` before committing and make sure `.env` is not staged. If Git is already initialized or a remote already exists, don't repeat those setup commands; update the existing repository instead.

### 2. Create the Render service

1. Sign in at [Render Dashboard](https://dashboard.render.com/).
2. Select **New + → Web Service**.
3. Connect GitHub if prompted and select your repository.
4. Use these settings:

| Setting | Value |
|---|---|
| Name | `xebia-powerbi-proposal-studio` (or another available name) |
| Runtime | Python 3 |
| Branch | `main` |
| Root Directory | Leave blank; `requirements.txt` is at repository root |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn backend.main:app --host 0.0.0.0 --port $PORT` |
| Health Check Path | `/api/health` |

Choose an instance type suitable for your expected usage and budget. Free instance availability and behavior can change; check the current Render plan details.

### 3. Add environment variables in Render

In the service's **Environment** settings, add:

| Key | Value |
|---|---|
| `GROQ_API_KEY` | Your actual Groq API key (secret) |
| `GROQ_MODEL` | `openai/gpt-oss-120b` or another model available to your account |
| `MAX_UPLOAD_MB` | `25` |

Do not add the API key to GitHub or to frontend JavaScript. Save the settings and deploy.

### 4. Verify the deployed service

When Render reports a successful deployment:

1. Open the public service URL shown in the Render dashboard.
2. Open `<your-render-url>/api/health`; verify `status` is `ok` and `configured` is `true`.
3. Generate a proposal using a test requirement of at least 80 characters, or upload a supported BRD.
4. Download and open the generated `.pptx` in PowerPoint.
5. If deployment fails, inspect the Render build and runtime logs for the first error.

For the production start command, do not use `--reload`. Render provides the port through the `$PORT` environment variable.

## How to use the app

1. Open the application.
2. Optionally enter the client name, industry, and tone.
3. Choose **one** requirements method:
   - paste/type requirements (at least 80 characters), or
   - upload a PDF, DOCX, or TXT BRD.
4. Generate the proposal.
5. Download the generated PowerPoint file.

The proposal content depends on the quality of the supplied requirements and the Groq response. Review all generated assumptions, technical recommendations, DAX, estimates, and client-facing claims before sharing the deck externally.

## Output and storage considerations

Generated files are written to the project's `outputs/` directory. The app returns a download URL after generation. On a hosting service with an ephemeral filesystem, locally generated PPTX files may be lost after a restart, redeployment, or instance replacement. Download generated files promptly. For reliable long-term retention or multi-user production use, add persistent/object storage and appropriate access controls.

## Troubleshooting

### `Groq is not configured`

- Check that `GROQ_API_KEY` is set in local `.env` or Render's Environment settings.
- Restart the local server or redeploy/restart the Render service after changing variables.
- Make sure the key is not the example placeholder.

### Groq API errors or model unavailable

- Verify the key and model are valid for your Groq account.
- Review the server logs and the error returned by the API.
- API access, rate limits, model availability, and pricing are controlled by Groq and may change.

### Upload rejected

- Only `.pdf`, `.docx`, and `.txt` files are supported.
- The default maximum file size is 25 MB; change `MAX_UPLOAD_MB` if needed.
- A scanned PDF without extractable text may need OCR before upload.
- Use either pasted requirements or an uploaded file, not both.

### Proposal generation fails or takes a long time

- Confirm the requirement source contains at least 80 characters.
- Check Groq API status, key permissions, model availability, and rate limits.
- Inspect local terminal output or Render runtime logs.

### PowerPoint download fails after a restart

The app currently writes generated files to local disk. If the hosting filesystem was cleared or the service restarted, generate the proposal again or configure persistent/object storage.

## Project structure

```text
.
├── backend/
│   ├── __init__.py
│   ├── main.py            # FastAPI routes, uploads, Groq request, downloads
│   ├── prompt.py          # Proposal-generation instructions
│   └── pptx_builder.py    # Editable PowerPoint creation and layout
├── frontend/
│   ├── index.html
│   ├── app.js
│   └── styles.css
├── knowledge/
├── outputs/
│   └── .gitkeep
├── .env.example
├── .gitignore
├── requirements.txt
└── README.md
```

## Important notes

- The fixed Xebia capability section is authored in `backend/pptx_builder.py`; update it when approved capability descriptions or branding change.
- Client-specific details not present in the requirements should be labelled as assumptions or recommendations and reviewed by a subject-matter expert.
- This is a proposal-generation starting point, not a complete secure multi-tenant document platform. Add authentication, authorization, rate limiting, logging controls, retention policies, and durable file storage before using it for sensitive client documents at scale.

## Official documentation

- [Render: Deploy a FastAPI app](https://render.com/docs/deploy-fastapi)
- [Render: Environment variables and secrets](https://render.com/docs/configure-environment-variables)
- [Groq API documentation](https://console.groq.com/docs)
