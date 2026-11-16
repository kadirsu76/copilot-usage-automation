"""
Configuration module for Copilot Usage Automation.
Loads settings from environment variables or .env file.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Base directory of the project
BASE_DIR = Path(__file__).resolve().parent

# Load .env if present
load_dotenv(BASE_DIR / ".env")

# Azure Authentication
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID", "").strip()
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "").strip()
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "").strip()

# Managed Identity endpoints (automatically injected by Azure Automation / App Service / Functions)
IDENTITY_ENDPOINT = os.getenv("IDENTITY_ENDPOINT") or os.getenv("MSI_ENDPOINT")
IDENTITY_HEADER = os.getenv("IDENTITY_HEADER") or os.getenv("MSI_SECRET")

# Report Configuration
# Graph v2 supports: D7, D28, D90, D180. If D30 is passed, it is mapped to D28 automatically.
raw_period = os.getenv("REPORT_PERIOD", "D28").strip().upper()
if raw_period == "D30":
    REPORT_PERIOD = "D28"
else:
    REPORT_PERIOD = raw_period if raw_period in ["D7", "D28", "D90", "D180"] else "D28"
OUTPUT_DIR = BASE_DIR / os.getenv("OUTPUT_DIR", "output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Email Settings
MAIL_SENDER = os.getenv("MAIL_SENDER", "").strip()
MAIL_RECIPIENTS = [
    email.strip()
    for email in os.getenv("MAIL_RECIPIENTS", "").split(",")
    if email.strip()
]
MAIL_SUBJECT = os.getenv(
    "MAIL_SUBJECT", "Microsoft 365 Copilot Kullanım ve Benimseme Raporu"
).strip()

# Graph API Base Endpoints
GRAPH_API_V1 = "https://graph.microsoft.com/v1.0"
GRAPH_API_BETA = "https://graph.microsoft.com/beta"
LOGIN_MICROSOFT_ONLINE = "https://login.microsoftonline.com"
