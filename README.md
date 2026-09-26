# Letterdrop

A small, local-first bulk email tool for sending individual, personalized messages to a team from an Excel `.xlsx` or `.csv` list. No account, database, or install step is required. The app uses Python's standard library.

## Start

1. Install Python 3.9 or newer if it is not already installed.
2. In this folder, run `python server.py` (Windows: `py server.py` also works).
3. Open `http://127.0.0.1:8000` in your browser.

The server listens only on your own computer. Keep the terminal open while using the app; press Ctrl+C to stop it.

## Prepare your spreadsheet

- Use the first worksheet of an `.xlsx` workbook, or a `.csv` file.
- Put column names in the first row and include an `Email` or `Email Address` column.
- Optional columns such as `First name`, `Last name`, and `Department` can be used in the message as `{{ First name }}` or `{{ Department }}`.
- Invalid email addresses and duplicates are skipped. Files are limited to 15 MB and 500 valid recipients per send.

## Connect email

Open **Set up SMTP** in the app and enter your SMTP host, email address, and app password. The app uses port `587` automatically. Common hosts are Gmail `smtp.gmail.com` and Microsoft 365 `smtp.office365.com`. Use your provider's app password or SMTP credential when required, not your normal account password. SMTP access may need to be enabled in your provider account. The tool sends a separate email for each person and uses TLS.

Credentials are sent only to the local server for the active send and are not saved. Email delivery is handled by your configured provider; its usage limits and anti-spam policies apply. Send only to people who have agreed to receive your messages, and try a small test list first.

## Deploy to Render

The `render.yaml` Blueprint describes the web service. Render needs this project in a GitHub, GitLab, or Bitbucket repository before it can build it. Push the project to a private repository, connect that provider in Render, then create a Blueprint Instance for the repository. Set `APP_USERNAME` and a strong `APP_PASSWORD` when prompted; these protect the public site with a browser sign-in. Do not commit real passwords or SMTP credentials.

The Blueprint uses a paid Starter web service because Render's free web services block outbound SMTP ports 25, 465, and 587. Costs depend on Render's current pricing. Review the plan and price in Render before confirming deployment. The service binds to Render's assigned port and exposes only a health-check endpoint without sign-in.

## Privacy and deployment

Uploaded contacts are processed in memory and not written to disk. Render terminates HTTPS for public requests. The deployment requires a shared password to prevent anonymous use, but is intended for a trusted small group rather than as a multi-user service. Anyone with the sign-in can send through the configured SMTP account, subject to provider limits and abuse protections.
