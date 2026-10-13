# Xebia Power BI Proposal Studio — Groq Editorial Edition

This is a fresh implementation focused on **substantive consulting proposals** and an editorial, presentation-first PowerPoint design. It does not use Gemini.

## Design principles
- No repeated grid of rounded content boxes.
- Typography-led slides with strong hierarchy, wide margins and whitespace.
- Full-width narrative, structured tables, clean process flows, roadmap timelines, KPI catalogues and risk registers.
- Eight fixed Xebia capability and accelerator slides first; a dedicated client requirements slide follows immediately, then the client-specific solution.
- AI output is validated. API failures are shown as errors; the app does not silently replace them with generic fallback content.
- Editable PowerPoint elements via `python-pptx`.
- Choose exactly one requirements input: upload a BRD (PDF/DOCX/TXT) OR type/paste plain text. Other form fields are optional.
- Fixed capability section covers Power BI consulting, Fabric/data engineering, semantic models/KPI governance, reporting/data storytelling, Tableau Workbook Validator, Qlik-to-Power BI Expression Translator, Power BI Playbook Builder, and the rationale for reusable metadata/migration accelerators.
- PowerPoint tables use native PowerPoint table elements with consistent column widths, cell alignment, typography and restrained row styling.

## Windows / VS Code setup
1. Extract the ZIP and open this folder in VS Code.
2. In a terminal at the project root:
   ```powershell
   py -m venv .venv
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\.venv\Scripts\Activate.ps1
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```
3. Create `.env` from the example:
   ```powershell
   Copy-Item .env.example .env
   notepad .env
   ```
4. Add your Groq key and model to `.env`. Keep the key private.
5. Start:
   ```powershell
   python -m uvicorn backend.main:app --reload
   ```
6. Browse to http://127.0.0.1:8000

## Important
- The model name must be available to your Groq account. The app surfaces API errors rather than silently generating a generic proposal.
- The fixed Xebia capability section is a carefully authored baseline. Update `backend/pptx_builder.py` if your approved brand wording changes.
- Client-specific technical details not supplied in requirements are labelled as recommendations or assumptions.
