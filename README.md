# ResumeFlow — AI Resume Screening Workspace

FastAPI web app with Google OAuth support, separate job campaigns, ZIP/multi-file upload, PDF/DOCX/TXT parsing, skill matching, TF-IDF relevance scoring, explainable weighted scores, reviewer decisions, CSV export, PostgreSQL support, and a warm editorial UI. Fonts: Newsreader, Plus Jakarta Sans, DM Mono (free alternatives; not proprietary Anthropic fonts).

## Run locally (Windows PowerShell)
```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
Copy-Item .env.example .env
uvicorn app.main:app --reload
```
Open http://127.0.0.1:8000. Without Google credentials, local demo login is available. Never deploy demo login publicly.

## Enable Google sign-in
1. In Google Cloud Console create/select a project and configure OAuth consent.
2. Create an OAuth Client ID (Web application).
3. Add redirect URI `http://127.0.0.1:8000/auth/google/callback` for local development.
4. For Render, add `https://YOUR-SERVICE.onrender.com/auth/google/callback` to authorized redirect URIs.
5. Set `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `GOOGLE_REDIRECT_URI`, `SECRET_KEY`, and `ALLOWED_EMAILS` in environment settings. `ALLOWED_EMAILS` is a comma-separated HR email allowlist. Set `APP_ENV=production` to disable demo login. Never commit secrets.

## Deploy on Render
Push this folder to GitHub, create a Render Blueprint from `render.yaml`, set OAuth variables, confirm database plan/pricing, register the exact HTTPS callback URI with Google, and test `/health`. Current plan availability and pricing can change; review before provisioning.

## Scoring
40% required skills, 30% TF-IDF text similarity, 20% experience fit, 10% preferred skills. This is a heuristic score, not a hiring probability. Validate against human-reviewed relevance labels and metrics such as Precision@K and NDCG@K.

## Production caveats
This is a deployable starter, not an audited enterprise recruitment system. Batch processing uses an in-process thread, not a durable queue; it can be interrupted by restarts/deploys. For large or concurrent workloads, use a durable worker (Celery/RQ + Redis or managed equivalent) and private object storage. Upload bytes are processed in memory; original files are not permanently stored for later review. Add CSRF protection, rate limiting, request size limits, malware scanning, encryption, retention/deletion, audit logging, backups, and security review before real candidate data. Scanned PDFs need OCR. Name and experience extraction are heuristics. Scores require human review. Do not use protected attributes or proxies to rank candidates.


## Campaign management updates
- Campaigns can be edited from the campaign page; state can be Active or Inactive. Inactive campaigns do not accept new screening uploads.
- Job descriptions support bold using the Bold toolbar (stored as **bold text**) and a Know more toggle on the campaign page.
- Repeat resumes in the same role are skipped using a normalized resume-text hash, email address, or detected applicant name. This is a practical duplicate check; common names can occasionally create false matches, so verify before relying on it.
- Workspace metrics show unique applicants, interview shortlist, and candidates needing review instead of raw screening-batch counts.
