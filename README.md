# Training reminder dashboard

Upload an Excel training plan, review imported rows, configure SMTP, and enable automatic email reminders. Reminder times use India time (`Asia/Kolkata`). Uploaded rows start paused so you can review them before sending.

## Deploy to Render

The repository includes `render.yaml` for Render's Blueprint flow. The cloud version stores the workbook and SMTP settings on a persistent disk, protects the dashboard with a username and password, and runs its minute-by-minute reminder scheduler inside the web service.

1. Put this project in a private GitHub repository. Do not commit your local workbook, SMTP settings, passwords, or `.env` file; they are excluded by `.gitignore`.
2. In Render, choose **New + → Blueprint**, connect the repository, and deploy the `render.yaml` Blueprint. The Blueprint creates a Starter web service and a 1 GB disk mounted at `/var/data`.
3. When prompted for environment values, set a unique `DASHBOARD_USERNAME` and strong `DASHBOARD_PASSWORD`.
4. Generate an encryption key on your computer and set it as `SMTP_ENCRYPTION_KEY` in Render:

   ```powershell
   python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
   ```

   Keep this key private and unchanged. The app uses it to encrypt the SMTP password saved from the dashboard. If the key changes, the saved password must be entered again.

5. Open the Render URL and sign in at the browser prompt. Upload your training workbook, review and activate the rows, enter SMTP credentials, check the connection, and enable automatic sending.

Render provides HTTPS for the public service URL. The service's persistent disk keeps app data across restarts and deployments. Keep the service at one instance: the workbook and scheduler are designed for one active app process.

## Run locally on Windows

Install dependencies and start the app from PowerShell or Terminal:

```powershell
python -m pip install -r requirements.txt
python web_app.py
```

The dashboard opens at `http://127.0.0.1:8765`. Keep the process running while using the dashboard. **Enable automatic sending** checks reminders inside that server process once per minute; it does not open another Python window or install a Windows scheduled task.

## Import your Excel schedule

Click **Upload Excel** and choose an `.xlsx` workbook. The importer recognizes training plan columns for serial number, department, training topic, plan status, actual status, reminder date, actual training date, replan status, reminder time, training time, trainer name, trainee employee, and email. It accepts common header variations, including source workbook misspellings such as `Remainder Date`, `Remainder Time`, `Actual Tranining Date`, and `Trainner Name`.

Rows need a reminder date and time and an actual training date. The import replaces the current reminder list in `training_reminders.xlsx` and sets each imported row to Paused. Review the dates and email addresses, then set rows to Ready when approved. Rows without a valid email cannot be activated. The page reports row errors without partially importing the workbook.

## SMTP setup

Open **Configure SMTP** and enter your provider's server host, port, connection security, sender address, username, and password or app password. Click **Save settings**, then **Check connection**. On Render, the password is encrypted with `SMTP_ENCRYPTION_KEY`; locally, Windows Data Protection secures it for your signed-in account.

For Gmail, use `smtp.gmail.com`, port `587` with STARTTLS or port `465` with SSL, and an app password if Google requires one. See Google's [SMTP settings](https://support.google.com/mail/answer/7104828) and [app password help](https://support.google.com/accounts/answer/185833).

**Send due now** sends eligible reminders immediately. **Enable automatic sending** checks every minute. Sent rows are marked `Sent`; failures show `Error` and an explanation. A row left at `Sending` after an interruption may have been delivered, so check the Sent folder before setting it to Ready again.
