# Letterdrop

A small bulk email tool for sending individual, personalized messages to a team from an Excel `.xlsx` or `.csv` list. It uses Python's standard library plus `cryptography` to protect scheduled campaign data.

## Start

1. Install Python 3.9 or newer if it is not already installed.
2. In this folder, install dependencies with `python -m pip install -r requirements.txt`.
3. Run `python server.py` (Windows: `py server.py` also works).
4. Open `http://127.0.0.1:8000` in your browser.

The server listens only on your own computer. Keep the terminal open while using the app; press Ctrl+C to stop it.

## Prepare your spreadsheet

- Use the first worksheet of an `.xlsx` workbook, or a `.csv` file.
- Put column names in the first row and include an `Email` or `Email Address` column.
- Optional columns such as `First name`, `Last name`, and `Department` can be used in the message as `{{ First name }}` or `{{ Department }}`.
- Invalid email addresses and duplicates are skipped. Files are limited to 15 MB and 500 valid recipients per send.

## Connect Resend

Create a Resend account, add and verify a sending domain, and create an API key with permission to send email. Set the key as the server environment variable `RESEND_API_KEY`; never put it in the browser or commit it to Git. In the app, enter a sender address on that verified domain. Resend's test sender can only be used within the restrictions shown in your Resend account. The app calls Resend's HTTPS API, so the hosting provider does not need outbound SMTP access.

The API key stays on the server and is not included in browser requests or scheduled campaign data. Delivery is subject to Resend's account limits and policies. Send only to recipients who have agreed to receive your messages, and try a small test list first.

## Schedule email

Choose **Schedule for later**, pick a date and time in the browser device's local time zone, then confirm. The app sends the campaign automatically and shows scheduled, sending, sent, partial, failed, or cancelled status in **Your scheduled sends**. You can cancel a campaign until delivery starts. This is an email scheduler, not a push notification; the scheduled email still sends when the app page is closed, as long as its server is running.

Recipients and message content for queued campaigns are encrypted before being stored. The payload is erased after the campaign finishes or is cancelled; status and schedule metadata remain. Locally, the database and generated encryption key live in `data/`. Back up and protect that folder if you need to preserve pending scheduled sends across a move or reinstall. Keep the server running so due mail can be delivered.

## Deploy to Render

The `render.yaml` Blueprint describes the web service. Render needs this project in a GitHub, GitLab, or Bitbucket repository before it can build it. Push the project to a private repository, connect that provider in Render, then create a Blueprint Instance for the repository. Set `APP_USERNAME` and a strong `APP_PASSWORD` when prompted; these protect the public site with a browser sign-in. Do not commit real passwords or API keys.

The Blueprint uses a paid Starter web service and a persistent disk for the encrypted schedule database. The disk is required because Render's regular filesystem is ephemeral. Add `RESEND_API_KEY` as a secret environment variable in Render; do not commit it. Resend uses HTTPS, so this app does not require SMTP ports. Costs depend on current pricing; review all charges before deploying. The Blueprint generates an `APP_ENCRYPTION_KEY`; keep it secret and persistent or existing pending campaign data cannot be decrypted. The service binds to Render's assigned port and exposes only a health-check endpoint without sign-in.

## Privacy and deployment

Contacts uploaded for immediate sends are processed in memory and not written to disk. For scheduled sends, recipient data and message content are encrypted on the persistent disk until delivery. Render terminates HTTPS for public requests. The deployment requires a shared password to prevent anonymous use, but is intended for a trusted small group rather than as a multi-user service. Anyone with the sign-in can send through the configured Resend account, subject to provider limits and abuse protections.
