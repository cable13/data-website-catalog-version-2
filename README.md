🌐 iHARP Automated Web Catalog Pipeline

A fully decoupled, cloud-automated static site generator designed to securely pull restricted research data from institutional Google Sheets and deploy it as a fast, interactive web catalog via GitHub Pages.

🚀 The Problem & The Solution

The Problem:
The iHARP catalog data (Literature, Datasets, Syllabi) is managed by researchers in UMBC Google Sheets. Because these sheets are locked behind strict institutional Single Sign-On (SSO) security, standard client-side web fetching (like JavaScript fetch()) gets blocked by the UMBC login screen.

The Solution:
This pipeline completely decouples the secure database (Google Sheets) from the public frontend (HTML/CSS) using a CI/CD cloud architecture.

Secure Extraction: A backend Python script authenticates via Google Cloud OAuth 2.0 to securely read the locked sheets.

Data Sanitization: The script cleans messy metadata (e.g., standardizing broken timestamp regex) and handles missing columns.

Static Compilation: The data is compiled directly into a static HTML/JSON bundle.

Cloud Automation: GitHub Actions automatically runs this entire process in the cloud every hour, deploying the updated site to GitHub Pages.

Zero manual data entry, downloading, or HTML coding is required by the researchers.

✨ Key Features

🤖 Real-Time Automated Sync: GitHub Actions rebuilds the catalog automatically based on an hourly cron schedule or manual webhook.

🔐 Institutional Security Bypass: Securely authenticates with Google Cloud OAuth to read SSO-locked sheets without exposing them to the public internet.

⚡ Lightning Fast UI: Built with Tailwind CSS, featuring global search, dynamic multi-select DOM filtering (Year, Author, Venue), and Dark Mode. Fully static architecture means zero database latency.

🛡️ Fallback Mechanisms: If cloud credentials expire, the script gracefully falls back to parsing local .xlsx Excel downloads to keep the site online.

🛠️ Architecture & Setup

Prerequisites

To run or test the builder locally, you need Python 3.10+ and the following libraries:

pip install pandas openpyxl gspread google-auth-oauthlib


1. Local Execution (Manual Build)

To manually rebuild the index.html file on your local machine:

Ensure your Google Cloud client_secret.json and token.json files are in the root directory.

Run the build script:

python python_build_catalog.py


Open the newly generated index.html in your browser.

2. Cloud Automation (GitHub Actions CI/CD)

This repository contains a GitHub Actions workflow (.github/workflows/update_catalog.yml) that automates the entire build process.

To enable this on a fork or new repository:

Go to your GitHub Repository Settings > Secrets and variables > Actions.

Add your Google Cloud credentials as two repository secrets:

CLIENT_SECRET (Paste the exact contents of your client_secret.json file)

TOKEN_JSON (Paste the exact contents of your token.json file)

The GitHub Actions server will now spin up every hour, securely inject these secrets, run the Python script, and push the updated index.html to your main branch.

🔒 Security & Credential Management

CRITICAL: Never upload your secret keys to GitHub.

This repository utilizes a strict .gitignore file to ensure client_secret.json and token.json are never committed to version control.

The cloud automation strictly handles these credentials via encrypted GitHub Secrets. The Python script generates these files dynamically in a secure container, runs the data fetch, and permanently deletes them before pushing the website code live.

Designed and engineered for the iHARP / CISAAD Data Science Scholars Program.
