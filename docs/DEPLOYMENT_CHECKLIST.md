# Deployment checklist
- Run `python -m py_compile app/main.py` and test `/health`.
- Set stable `SECRET_KEY`, production `DATABASE_URL`, and `APP_ENV=production`.
- Configure Google OAuth client and exact HTTPS callback URL.
- Set `ALLOWED_EMAILS` to authorized HR accounts.
- Never commit `.env`, resumes, or databases to GitHub.
- Use durable workers/private storage and add CSRF, rate limiting, malware scanning, audit logs, retention/deletion, and security review before real applicant data.
