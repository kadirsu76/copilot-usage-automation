# Microsoft 365 Copilot Adoption & Usage Automation

Automated, executive-grade reporting of Microsoft 365 Copilot usage, adoption trends, app penetration, and license optimization using Azure Automation, Managed Identity, and Microsoft Graph API.

- **Azure Automation Schedule:** Weekly on Mondays at 08:00 UTC (or Daily)
- **Azure Automation Python 3 Runbook:** Standalone execution with System-Assigned Managed Identity (zero stored secrets)
- **Executive Email Delivery:** 680px width optimized for Outlook Desktop, OWA, Gmail, and Apple Mail
- **Audit Roster Attachment:** Automatically generates and attaches an executive-styled `.xlsx` spreadsheet of all licensed accounts

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fkadirsu76%2Fcopilot-usage-automation%2Fmain%2Fazuredeploy.json)

---

## What It Reports

The automation gathers Microsoft Substrate telemetry via Microsoft Graph API v2 endpoints and generates an email-safe executive report:

1. **Executive KPI Cards:** Assigned seats, active users, adoption rate (%), total prompt volume, and inactive seats.
2. **Daily Activity & Prompt Momentum (Vector SVG):** Dual-metric time-series showing daily prompt submissions (bars) and active users (line) across the reporting period with peak dates and weekday averages.
3. **User Engagement Segmentation:**
   - **Power Users (50+ prompts):** Everyday advocates who integrate Copilot into core workflows.
   - **Regular Users (10–49 prompts):** Weekly active users ready for deeper adoption.
   - **Occasional Users (1–9 prompts):** Casual users needing targeted departmental training.
   - **Inactive (0 prompts):** Reallocation candidates for cost optimization.
4. **Application Penetration:** Active user counts and adoption percentages across 8 host apps: Microsoft Teams, Business Chat, Outlook, Word, Excel, PowerPoint, Loop, and OneNote.
5. **Top Active Contributors:** Leaderboard ranked by prompt volume, active days, and used apps.
6. **License Reallocation Candidates:** Accounts with zero recorded interactions in the reporting window.
7. **Detailed Excel Audit Roster (.xlsx):** Complete directory of all licensed seats with individual app engagement timestamps, zebra rows, freeze panes, and status badges.

---

## Architecture & Azure Resources Created

When deployed via ARM / Bicep template or the **Deploy to Azure** button, the following Azure resources are provisioned:

| Resource | Name Format | Description |
|---|---|---|
| **Automation Account** | `aa-<baseName>` | Serverless execution environment with **System-Assigned Managed Identity** enabled |
| **Automation Runbook** | `rb-<baseName>-report` | Python 3 runbook containing the report generator and Graph mail dispatcher |
| **Automation Variables** | `CopilotReport-*` | Configuration variables (`Period`, `MailSender`, `MailRecipients`, `MailSubject`) |
| **Automation Schedule** | `sch-<baseName>-weekly` | Recurring trigger (Mondays at 08:00 UTC or Daily) |
| **Job Schedule** | `jobSchedule-*` | Links the recurring schedule to the Python runbook |
| **Deployment Script Identity** | `mi-<baseName>-ds` | User-Assigned Identity used during deployment to publish the runbook and link the schedule |

---

## Prerequisites

- Azure Subscription with Contributor access.
- Microsoft Entra ID (Azure AD) Global Administrator or Privileged Role Administrator (for one-time Graph app role consent).
- Exchange Online mailbox for `mailSender` (e.g. `reports@yourdomain.com`).
- Local tools (for CLI deployment):
  - Azure CLI (`az`)
  - PowerShell 7+ (`pwsh`)

---

## Deployment

### Option 1: One-Click Deploy via Azure Portal

Click the button below to launch the deployment wizard in Azure Portal:

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fkadirsu76%2Fcopilot-usage-automation%2Fmain%2Fazuredeploy.json)

1. Select your **Subscription** and **Resource Group**.
2. Enter your `baseName` (e.g. `copilotrep`).
3. Set `mailSender` (e.g. `reports@yourdomain.com`) and `mailRecipients` (e.g. `cio@yourdomain.com, it-ops@yourdomain.com`).
4. Click **Review + Create**.

---

### Option 2: Deploy via PowerShell CLI

```powershell
pwsh ./scripts/Deploy-Azure.ps1 `
  -ResourceGroupName "rg-copilot-reporting" `
  -Location "westeurope" `
  -BaseName "copilotrep" `
  -ReportPeriod "D28" `
  -MailSender "reports@yourdomain.com" `
  -MailRecipients "management@yourdomain.com, it-lead@yourdomain.com" `
  -ScheduleFrequency "Week"
```

The script automatically derives all resource names, ensures the Resource Group exists, deploys the Bicep template, publishes the runbook, and returns the Managed Identity Principal (Object) ID.

---

## Grant Microsoft Graph Permissions

After deploying the infrastructure, grant the required Microsoft Graph application roles to the Automation Account's System-Assigned Managed Identity:

```powershell
pwsh ./scripts/Grant-GraphPermissions.ps1 -Mi <automation-principal-id>
```

> **Note:** `<automation-principal-id>` is printed at the end of `Deploy-Azure.ps1`, or can be copied from **Azure Portal > Automation Account (`aa-<baseName>`) > Identity > Object ID**.

### Permissions Granted:
- `Reports.Read.All`: Reads Microsoft 365 Copilot user activity telemetry and daily trends.
- `Mail.Send`: Dispatches the HTML report with Excel attachment via Microsoft Graph `sendMail`.

---

## Trigger Report On-Demand

You can trigger an immediate run from PowerShell without waiting for the schedule:

```powershell
pwsh ./scripts/Start-ReportNow.ps1 -ResourceGroupName "rg-copilot-reporting" -BaseName "copilotrep" -Wait
```

Use `-Wait` to follow the job in real time until completion.

---

## Local Development & Testing

You can also run the report generator locally on macOS / Linux / Windows:

```bash
# 1. Clone repository
git clone https://github.com/kadirsu76/copilot-usage-automation.git
cd copilot-usage-automation

# 2. Setup virtual environment
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 3. Test run with Mock Data (no credentials needed)
python main.py --mock

# 4. Live run with Service Principal (.env configured)
python main.py --period D28

# 5. Open generated files
open output/copilot_usage_report_latest.html
open output/copilot_usage_roster_latest.xlsx
```

---

## Repository Structure

```
.
├── azuredeploy.json                   # Root ARM entry template
├── azuredeploy.parameters.json        # Template parameters
├── infra/
│   ├── main.bicep                     # Bicep IaC source definition
│   ├── main.json                      # Compiled ARM template
│   └── publish-runbook.ps1            # Deployment script to publish runbook & link schedule
├── runbook/
│   └── Run-CopilotUsageReport.py      # Standalone single-file Python runbook
├── scripts/
│   ├── Deploy-Azure.ps1               # Azure CLI deployment script
│   ├── Grant-GraphPermissions.ps1     # Microsoft Graph app role assignment script
│   └── Start-ReportNow.ps1            # On-demand job trigger script
├── auth.py                            # Dual-mode authentication (MI / SP)
├── graph_client.py                    # Graph API telemetry client
├── report_generator.py                # 680px email-safe HTML generator
├── excel_generator.py                 # Styled Excel audit roster generator (.xlsx)
├── mailer.py                          # Graph API sendMail client
├── main.py                            # Local CLI entrypoint
├── requirements.txt                   # Local Python dependencies
└── README.md
```

---

## Design Principles

- **Zero Emojis:** Built for enterprise and executive briefings.
- **Zero AI Clichés:** Human-crafted, clear typography with tabular numerals (`tabular-nums`).
- **100% Inline Vector Assets:** Application icons and trend visuals are self-contained SVGs. No broken external image links or privacy block warnings in Outlook.
- **Email-Safe HTML:** Strictly table-structured, 680px max-width, no JavaScript, no inner scrollbars.
