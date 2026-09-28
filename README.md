# Training Desk

Training Desk is a public training-reminder website. Visitors can create an account, upload and manage their own training schedule, configure their own SMTP account, and enable automatic reminders. Each account has a separate workbook, SMTP settings file, and scheduler state.

## Deploy to Render

The repository includes `render.yaml` for Render's Blueprint flow.

1. In Render, choose **New + → Blueprint**, connect this GitHub repository, select the `main` branch, and deploy the Blueprint.
2. When prompted for `SMTP_ENCRYPTION_KEY`, generate a Fernet key on your computer:

   ```powershell
   python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
   ```

   Keep the key secret and unchanged. The app uses it to encrypt each user's saved SMTP password. If it is changed or lost, users must save their SMTP settings again.

3. Open the Render URL. Anyone can create an account from the sign-in page. Each user must upload their own training workbook and configure their own SMTP account before sending reminders.

The Blueprint creates a Starter web service with a 1 GB persistent disk mounted at `/var/data`. That disk stores account records and each user's app data across deployments. Keep the service at one instance because the SQLite account store and reminder scheduler are designed for a single active app process. Render provides HTTPS for the public URL.

## What each account can do

- Register with a unique username and a password of at least 12 characters.
- Upload a training plan. Imported rows start paused for review.
- Enter SMTP credentials for their own email account. Passwords are encrypted at rest; the dashboard never displays a saved password.
- Send due reminders manually or enable the minute-by-minute scheduler.

Each person's workbook and SMTP settings are isolated in their account folder. No central SMTP account is shared among visitors. Users need an SMTP provider that permits their account to send mail; Gmail may require an app password. See Google's [SMTP settings](https://support.google.com/mail/answer/7104828) and [app password help](https://support.google.com/accounts/answer/185833).

## Run locally on Windows

Install dependencies and start the app from PowerShell or Terminal:

```powershell
python -m pip install -r requirements.txt
python web_app.py
```

Open `http://127.0.0.1:8765`, create an account, then upload your workbook and configure SMTP. Keep the dashboard process running for automatic reminders.

## Public service notes

The sign-up page is open to the public. Do not use your personal SMTP password as a shared account credential; visitors configure their own SMTP settings. The GitHub repository contains source code only. `.gitignore` excludes local workbooks, SMTP settings, logs, environment files, and app databases.
