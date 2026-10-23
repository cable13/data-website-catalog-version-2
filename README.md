🌐 iHARP Automated Web Catalog Pipeline

A decoupled, automated static-site generator designed to securely pull restricted research data from institutional Google Sheets and deploy it as a fast, interactive web catalog via GitHub Pages.

🚀 The Problem & The Solution

The Problem: The iHARP catalog data (Literature, Datasets, Syllabi) is managed by researchers in UMBC Google Sheets. Because these sheets are locked behind strict institutional Single Sign-On (SSO) security, standard client-side web fetching (like JavaScript fetch()) gets blocked by the UMBC login screen.

The Solution: This pipeline completely decouples the database (Google Sheets) from the frontend (HTML/CSS).

A backend Python script authenticates via Google Cloud (OAuth 2.0 / Service Account).

It securely fetches the raw spreadsheet data and cleans messy metadata (e.g., standardizing timestamp regex).

It compiles the data directly into a static HTML/JSON bundle.

GitHub Actions CI/CD automatically runs this process in the cloud every hour, deploying the updated site to GitHub Pages.

Zero manual data entry or HTML coding is required by the researchers.

✨ Features

Real-Time Automated Sync: GitHub Actions rebuilds the catalog automatically based on a cron schedule or webhook.

Institutional Security Bypass: Securely authenticates with Google Cloud to read SSO-locked sheets without exposing them to the public internet.

Offline Fallback: If cloud credentials expire, the script gracefully falls back to parsing local .xlsx Excel downloads.

Modern UI/UX: Built with Tailwind CSS, featuring global search, dynamic multi-select DOM filtering (Year, Author, Venue), and Dark Mode.

Lightning Fast: Fully static architecture means zero database latency and instant page loads.

🛠️ Architecture Setup

Prerequisites

To run the builder locally, you need Python 3.10+ and the following libraries:

pip install pandas openpyxl gspread google-auth-oauthlib


1. Local Execution (Manual Build)

If you want to manually rebuild the index.html file on your local machine:

Ensure your Google Cloud client_secret.json and token.json files are in the root directory.

Run the build script:

python python_build_catalog.py


Open the newly generated index.html in your browser.

2. Cloud Automation (GitHub Actions)

This repository contains a GitHub Actions workflow (.github/workflows/update_catalog.yml) that automates the entire build process.

To enable this on a fork:

Go to your GitHub Repository Settings > Secrets and variables > Actions.

Add your Google Cloud credentials as two repository secrets:

CLIENT_SECRET (Paste the contents of your client_secret.json)

TOKEN_JSON (Paste the contents of your token.json)

The server will now spin up every hour, inject these secrets securely, run the Python script, and push the updated index.html to your main branch.

🔒 Security Note

Never upload your secret keys to GitHub.
This repository utilizes a .gitignore file to ensure client_secret.json, token.json, and umbc_bot_credentials.json are never committed to version control. The cloud automation strictly handles credentials via encrypted GitHub Secrets.
