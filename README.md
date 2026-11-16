# Microsoft 365 Copilot Adoption & Usage Automation

Automated Copilot usage and adoption reporting with:

- Azure Automation schedule (weekly Mondays 08:00 UTC or daily)
- Azure Automation Python 3 runbook (System-Assigned Managed Identity, zero secrets)
- Email report delivery (680px email-client optimized)
- Detailed user audit roster attached as Excel spreadsheet (`.xlsx`)

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fkadirsu76%2Fcopilot-usage-automation%2Fmain%2Fazuredeploy.json)

---

## What It Reports

- **KPI Cards:** Assigned seats, active users, adoption rate (%), total prompt volume, inactive seats.
- **Daily Trend (Inline SVG):** Daily prompt volume (bars) and active users (line) time-series.
- **User Engagement:** Power Users (50+ prompts), Regular (10–49), Occasional (1–9), Inactive (0).
- **Application Penetration:** Teams, Business Chat, Outlook, Word, Excel, PowerPoint, Loop, OneNote.
- **Top Contributors & Reallocation:** Top active accounts and inactive seats for license optimization.
- **Excel Audit Attachment:** Full 30-user directory with app timestamps, freeze panes, and zebra styling.

---

## Deploy

### 1. Deploy Azure Resources

Click **Deploy to Azure** above or run from CLI:

```powershell
pwsh ./scripts/Deploy-Azure.ps1 `
  -ResourceGroupName "rg-copilot-reporting" `
  -Location "westeurope" `
  -BaseName "copilotrep" `
  -MailSender "reports@yourdomain.com" `
  -MailRecipients "cio@yourdomain.com"
```

Resource names derived automatically:
- Automation Account: `aa-<baseName>`
- Runbook: `rb-<baseName>-report`
- Schedule: `sch-<baseName>-weekly`

---

### 2. Grant Graph Permissions

Assign required Graph app roles to the Automation Managed Identity:

```powershell
pwsh ./scripts/Grant-GraphPermissions.ps1 -Mi <automation-principal-id>
```

Permissions granted:
- `Reports.Read.All`: Copilot usage telemetry and daily trends.
- `Mail.Send`: Dispatching report email with Excel attachment.

---

### 3. Restrict Mail.Send to Single Mailbox (Least Privilege)

To restrict `Mail.Send` so the Managed Identity can ONLY send email from the designated sender mailbox:

```powershell
pwsh ./scripts/Set-MailboxAccessPolicy.ps1 `
  -ManagedIdentityId <automation-principal-id> `
  -Mailbox "reports@yourdomain.com"
```

Uses Exchange Online `New-ApplicationAccessPolicy` with `RestrictAccess`. All other mailboxes in the tenant remain protected.

---

## Trigger Report On-Demand

```powershell
pwsh ./scripts/Start-ReportNow.ps1 -ResourceGroupName "rg-copilot-reporting" -BaseName "copilotrep" -Wait
```

---

## Local Run

```bash
pip install -r requirements.txt
python main.py --mock        # Test run with sample data (no credentials needed)
python main.py --period D28  # Live run with .env credentials
```

---

## Notes

- Email subject is automatic with the current date: `Microsoft 365 Copilot Adoption & Usage Report - 16 Nov 2026`.
- 100% inline SVG assets — zero broken image links in Outlook.
- No client secrets or passwords required when running in Azure Automation.
