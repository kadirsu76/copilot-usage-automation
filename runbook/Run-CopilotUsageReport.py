#!/usr/bin/env python3
"""
Single-file Azure Automation Python Runbook.
Built to 'design-taste-frontend' standards:
  - Authentic Microsoft application vector marks
  - Daily Activity & Prompt Momentum dual-metric SVG time-series
  - Strict typographic hierarchy and tabular numbers
  - Zero emojis, zero AI slop words, zero external CSS/JS dependencies
Requires Azure Automation System-Assigned Managed Identity with 'Reports.Read.All' (and optionally 'Mail.Send') Graph permissions.
"""

import os
import io
import csv
import sys
import json
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional
import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("azure_runbook")

raw_period = os.getenv("REPORT_PERIOD", "D28").strip().upper()
REPORT_PERIOD = "D28" if raw_period in ["D28", "D30", ""] else raw_period

MAIL_SENDER = os.getenv("MAIL_SENDER", "").strip()
MAIL_RECIPIENTS = [
    e.strip() for e in os.getenv("MAIL_RECIPIENTS", "").split(",") if e.strip()
]

def get_default_subject() -> str:
    """Returns fixed email subject with current date (no time component)."""
    return f"Microsoft 365 Copilot Adoption & Usage Report - {datetime.now().strftime('%d %b %Y')}"

IDENTITY_ENDPOINT = os.getenv("IDENTITY_ENDPOINT") or os.getenv("MSI_ENDPOINT")
IDENTITY_HEADER = os.getenv("IDENTITY_HEADER") or os.getenv("MSI_SECRET")


def get_token() -> str:
    """Acquires a Microsoft Graph token strictly via Azure Automation Managed Identity."""
    if not IDENTITY_ENDPOINT or not IDENTITY_HEADER:
        raise RuntimeError(
            "Azure Managed Identity is required: IDENTITY_ENDPOINT / IDENTITY_HEADER environment variables not found. "
            "Ensure that System-Assigned Managed Identity is enabled on this Azure Automation Account."
        )

    logger.info("Acquiring Microsoft Graph token via Azure Automation System-Assigned Managed Identity...")
    url = f"{IDENTITY_ENDPOINT}?resource=https://graph.microsoft.com/&api-version=2019-08-01"
    headers = {"X-IDENTITY-HEADER": IDENTITY_HEADER}
    res = requests.get(url, headers=headers, timeout=25)
    res.raise_for_status()
    token = res.json().get("access_token")
    if not token:
        raise RuntimeError("Managed Identity endpoint did not return an access_token.")

    try:
        parts = token.split(".")
        if len(parts) >= 2:
            payload = parts[1]
            padded = payload + "=" * (-len(payload) % 4)
            claims = json.loads(base64.urlsafe_b64decode(padded).decode("utf-8"))
            roles = claims.get("roles", [])
            appid = claims.get("appid", "unknown")
            logger.info(f"Managed Identity Token Claims -> AppId: {appid}, Roles: {roles}")
            if "Reports.Read.All" not in roles:
                logger.warning(
                    f"Reports.Read.All permission not found in current token roles ({roles}). "
                    "Azure Managed Identity endpoint may still be returning a cached token. "
                    "Cache TTL usually clears within 30-60 minutes."
                )
    except Exception as parse_err:
        logger.debug(f"Could not parse token claims: {parse_err}")

    return token


def fetch_report(token: str, period: str = "D28") -> List[Dict[str, Any]]:
    """Fetches Copilot usage user detail report from Graph API v2."""
    if period.upper() in ["D30", ""]:
        period = "D28"
    url = f"https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='{period}')"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    logger.info(f"Querying Graph API user details: {url}")
    try:
        res = requests.get(url, headers=headers, allow_redirects=True, timeout=60)
        res.raise_for_status()
    except requests.exceptions.HTTPError as e:
        err_body = res.text if 'res' in locals() else str(e)
        logger.error(f"Graph API Request Failed (HTTP {res.status_code}): {err_body}")
        if res.status_code == 403:
            logger.error(
                "403 Forbidden: The Managed Identity is missing 'Reports.Read.All' application permission, "
                "or the Azure AD token cache has not propagated yet. "
                "Run: pwsh ./scripts/Grant-GraphPermissions.ps1 -Mi <ManagedIdentityPrincipalId>"
            )
        raise

    content_type = res.headers.get("Content-Type", "")
    raw_text = res.text.strip()

    if "application/json" in content_type or raw_text.startswith("{") or raw_text.startswith("["):
        try:
            data = res.json()
            raw_records = data.get("value", data) if isinstance(data, dict) else data
            return [_normalize_record(r) for r in raw_records]
        except Exception:
            pass

    reader = csv.DictReader(io.StringIO(raw_text))
    return [_normalize_record(row) for row in reader]


def fetch_trend(token: str, period: str = "D28") -> List[Dict[str, Any]]:
    """Fetches Copilot daily trend from Graph API."""
    if period.upper() in ["D30", ""]:
        period = "D28"
    url = f"https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountTrend(period='{period}')"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    logger.info(f"Querying Graph API daily trend: {url}")
    try:
        res = requests.get(url, headers=headers, allow_redirects=True, timeout=60)
        res.raise_for_status()
        val = res.json().get("value", [])
        if val:
            adoption = val[0].get("adoptionByDate", [])
            adoption.sort(key=lambda x: x.get("reportDate", ""))
            return adoption
    except Exception as e:
        logger.warning(f"Trend data could not be fetched: {e}")
    return []


def _normalize_record(raw: Dict[str, Any]) -> Dict[str, Any]:
    def get_val(*keys: str) -> Optional[str]:
        for k in keys:
            if k in raw and raw[k] is not None:
                v = str(raw[k]).strip()
                if v != "":
                    return v
            for rk, rv in raw.items():
                if rk.lower().replace(" ", "").replace("_", "") == k.lower().replace(" ", "").replace("_", ""):
                    if rv is not None and str(rv).strip() != "":
                        return str(rv).strip()
        return None

    upn = get_val("userPrincipalName", "User Principal Name", "upn") or "Unknown"
    display_name = get_val("displayName", "Display Name") or upn.split("@")[0]
    last_act = get_val("lastActivityDate", "Last Activity Date") or ""

    period_info = {}
    if "copilotActivityUserDetailsByPeriod" in raw and isinstance(raw["copilotActivityUserDetailsByPeriod"], list):
        if len(raw["copilotActivityUserDetailsByPeriod"]) > 0:
            period_info = raw["copilotActivityUserDetailsByPeriod"][0] or {}

    total_days = period_info.get("activeUsageDaysAllApps", 0) or 0
    prompts = period_info.get("promptsSubmittedAllApps", 0) or 0

    app_last_activity = {
        "Teams": get_val("microsoftTeamsCopilotLastActivityDate", "Teams Copilot Last Activity Date") or "",
        "Chat": get_val("copilotChatLastActivityDate", "copilotChatWorkLastActivityDate") or "",
        "Word": get_val("wordCopilotLastActivityDate", "Word Copilot Last Activity Date") or "",
        "Excel": get_val("excelCopilotLastActivityDate", "Excel Copilot Last Activity Date") or "",
        "PowerPoint": get_val("powerPointCopilotLastActivityDate", "PowerPoint Copilot Last Activity Date") or "",
        "Outlook": get_val("outlookCopilotLastActivityDate", "Outlook Copilot Last Activity Date") or "",
        "OneNote": get_val("oneNoteCopilotLastActivityDate", "OneNote Copilot Last Activity Date") or "",
        "Loop": get_val("loopCopilotLastActivityDate", "Loop Copilot Last Activity Date") or "",
    }
    apps_used = [app for app, act_date in app_last_activity.items() if act_date]
    is_active = bool(last_act or total_days > 0 or prompts > 0)

    return {
        "user_principal_name": upn,
        "display_name": display_name,
        "last_activity_date": last_act,
        "is_active": is_active,
        "apps_used": apps_used,
        "total_active_days": total_days,
        "prompts_submitted": prompts,
        "app_last_activity": app_last_activity,
    }


def _is_weekday(date_str: str) -> bool:
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").weekday() < 5
    except Exception:
        return True


def _render_app_icon(app_name: str) -> str:
    icons = {
        "Teams": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#5059C9"/><path d="M14.5 9h5v6h-5z" fill="#7B83EB"/><circle cx="17" cy="6.5" r="2" fill="#7B83EB"/><rect x="4.5" y="6" width="9" height="12" rx="2" fill="#4B53BC"/><path d="M7 9h4M9 9v6" stroke="#ffffff" stroke-width="1.6" stroke-linecap="round"/></svg>',
        "Chat": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#0284C7"/><path d="M6 8h12M6 12h8M6 16h5" stroke="#ffffff" stroke-width="1.8" stroke-linecap="round"/><circle cx="17.5" cy="15.5" r="2.5" fill="#38BDF8"/></svg>',
        "Outlook": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#0078D4"/><circle cx="14" cy="12" r="5" fill="#28A8EA"/><rect x="4" y="7" width="9" height="10" rx="1.5" fill="#005A9E"/><circle cx="8.5" cy="12" r="2.5" stroke="#ffffff" stroke-width="1.5"/></svg>',
        "Word": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#185ABD"/><rect x="10" y="5" width="10" height="14" rx="1" fill="#2B7CD3"/><rect x="4" y="7" width="9" height="10" rx="1.5" fill="#103F91"/><path d="M6 10l1.2 4 1.3-4 1.3 4 1.2-4" stroke="#ffffff" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>',
        "Excel": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#107C41"/><rect x="10" y="5" width="10" height="14" rx="1" fill="#18A058"/><rect x="4" y="7" width="9" height="10" rx="1.5" fill="#0B552C"/><path d="M6.5 10l4 4M10.5 10l-4 4" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/></svg>',
        "PowerPoint": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#C43E1C"/><rect x="10" y="5" width="10" height="14" rx="1" fill="#ED6C47"/><rect x="4" y="7" width="9" height="10" rx="1.5" fill="#93280F"/><path d="M7 14.5v-5h2.5a1.5 1.5 0 010 3H7" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/></svg>',
        "Loop": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#0891B2"/><circle cx="12" cy="12" r="5.5" stroke="#ffffff" stroke-width="1.8" stroke-dasharray="8 4"/></svg>',
        "OneNote": '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect width="24" height="24" rx="4" fill="#7719AA"/><rect x="10" y="5" width="10" height="14" rx="1" fill="#9B2FAF"/><rect x="4" y="7" width="9" height="10" rx="1.5" fill="#501075"/><path d="M6.5 14.5v-5l4 5v-5" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    }
    return icons.get(app_name, '<span style="display:inline-block;width:18px;height:18px;background:#e2e8f0;border-radius:4px;"></span>')


def render_trend_svg(trend_points: List[Dict[str, Any]]) -> str:
    if not trend_points:
        return ""
    n = len(trend_points)
    max_p = max((t.get("promptsSubmitted", 0) for t in trend_points), default=1) or 1
    max_u = max((t.get("anyAppActiveUsers", 0) for t in trend_points), default=1) or 1
    chart_left, chart_right, chart_top, chart_bottom = 40, 780, 24, 140
    chart_w, chart_h = chart_right - chart_left, chart_bottom - chart_top
    bar_spacing = chart_w / n
    bar_width = max(bar_spacing * 0.65, 4)

    bars, lines, dots, labels = [], [], [], []
    for i, t in enumerate(trend_points):
        xc = chart_left + (i + 0.5) * bar_spacing
        p = t.get("promptsSubmitted", 0)
        u = t.get("anyAppActiveUsers", 0)
        ds = t.get("reportDate", "")
        bh = (p / max_p) * (chart_h - 10)
        by = chart_bottom - bh
        fill = "#e2e8f0" if not _is_weekday(ds) else "#0f172a"
        bars.append(f'<rect x="{xc - bar_width/2:.1f}" y="{by:.1f}" width="{bar_width:.1f}" height="{max(bh, 1):.1f}" fill="{fill}" rx="2"><title>{ds}: {p} prompts, {u} active users</title></rect>')
        uy = chart_bottom - (u / max_u) * (chart_h - 10)
        lines.append(f"{xc:.1f},{uy:.1f}")
        dots.append(f'<circle cx="{xc:.1f}" cy="{uy:.1f}" r="3" fill="#0284c7" stroke="#ffffff" stroke-width="1.5"/>')
        if i % 7 == 0 or i == n - 1:
            try:
                sd = datetime.strptime(ds, "%Y-%m-%d").strftime("%d %b")
            except Exception:
                sd = ds[-5:]
            labels.append(f'<text x="{xc:.1f}" y="160" font-size="10" fill="#94a3b8" text-anchor="middle" font-family="-apple-system, sans-serif">{sd}</text>')

    return f"""
    <div style="background:#fff;border:1px solid #e2e8f0;border-radius:6px;padding:16px;margin-bottom:24px;">
        <div style="font-size:12px;font-weight:600;color:#64748b;margin-bottom:12px;display:flex;gap:18px;">
            <span><span style="display:inline-block;width:10px;height:10px;background:#0f172a;border-radius:2px;margin-right:4px;"></span>Daily Prompts (Max {max_p})</span>
            <span><span style="display:inline-block;width:12px;height:2px;background:#0284c7;margin-right:4px;"></span>Active Users (Max {max_u})</span>
        </div>
        <svg viewBox="0 0 800 170" style="width:100%;height:auto;display:block;">
            <line x1="40" y1="{chart_bottom}" x2="780" y2="{chart_bottom}" stroke="#e2e8f0"/>
            <text x="32" y="{chart_bottom}" font-size="10" fill="#94a3b8" text-anchor="end">0</text>
            <text x="32" y="{chart_top + 4}" font-size="10" fill="#94a3b8" text-anchor="end">{max_p}</text>
            {''.join(bars)}
            <polyline fill="none" stroke="#0284c7" stroke-width="2" points="{' '.join(lines)}"/>
            {''.join(dots)}
            {''.join(labels)}
        </svg>
    </div>
    """


def aggregate_metrics(users: List[Dict[str, Any]], period: str = "D28", trend: List[Dict[str, Any]] = None) -> Dict[str, Any]:
    total_users = len(users)
    active_users = sum(1 for u in users if u["is_active"])
    inactive_users = total_users - active_users
    adoption_rate = round((active_users / total_users * 100), 1) if total_users > 0 else 0.0
    total_prompts = sum(u.get("prompts_submitted", 0) for u in users)
    avg_prompts = round(total_prompts / active_users, 1) if active_users > 0 else 0.0

    apps = ["Teams", "Chat", "Outlook", "Word", "Excel", "PowerPoint", "Loop", "OneNote"]
    app_stats = {}
    for app in apps:
        count = sum(1 for u in users if u["app_last_activity"].get(app))
        rate = round((count / total_users * 100), 1) if total_users > 0 else 0.0
        app_stats[app] = {"active_count": count, "adoption_rate": rate}

    return {
        "period": period,
        "total_users": total_users,
        "active_users": active_users,
        "inactive_users": inactive_users,
        "adoption_rate": adoption_rate,
        "total_prompts": total_prompts,
        "avg_prompts": avg_prompts,
        "app_stats": app_stats,
        "trend_points": trend or [],
        "top_users": sorted(users, key=lambda x: (x.get("prompts_submitted", 0), x["total_active_days"]), reverse=True)[:10],
        "inactive_list": [u for u in users if not u["is_active"]],
    }


def generate_html(metrics: Dict[str, Any]) -> str:
    app_names = {
        "Teams": "Microsoft Teams", "Chat": "Business Chat", "Outlook": "Microsoft Outlook",
        "Word": "Microsoft Word", "Excel": "Microsoft Excel", "PowerPoint": "Microsoft PowerPoint",
        "Loop": "Microsoft Loop", "OneNote": "Microsoft OneNote",
    }
    apps_items = list(metrics["app_stats"].items())
    apps_rows = []
    for i in range(0, len(apps_items), 2):
        pair = apps_items[i:i+2]
        tds = []
        for app_name, stats in pair:
            title = app_names.get(app_name, app_name)
            rate = stats["adoption_rate"]
            icon_svg = _render_app_icon(app_name)
            tds.append(f"""
            <td style="width: 50%; vertical-align: top; padding: 4px; box-sizing: border-box;">
                <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px 12px;">
                    <table style="width: 100%; border-collapse: collapse; margin-bottom: 6px;">
                        <tr>
                            <td style="vertical-align: middle; padding: 0;">
                                <table style="border-collapse: collapse;">
                                    <tr>
                                        <td style="vertical-align: middle; padding-right: 7px; line-height: 0;">{icon_svg}</td>
                                        <td style="vertical-align: middle; font-size: 12px; font-weight: 600; color: #0f172a; white-space: nowrap;">{title}</td>
                                    </tr>
                                </table>
                            </td>
                            <td style="text-align: right; vertical-align: middle; font-size: 11px; font-weight: 600; color: #475569; font-variant-numeric: tabular-nums; white-space: nowrap; padding: 0;">
                                {stats['active_count']} <span style="color: #94a3b8; font-weight: 400;">({rate}%)</span>
                            </td>
                        </tr>
                    </table>
                    <div style="background: #f1f5f9; border-radius: 2px; height: 4px; overflow: hidden; width: 100%;">
                        <div style="background: #0f172a; width: {min(max(rate, 1), 100)}%; height: 4px; border-radius: 2px;"></div>
                    </div>
                </div>
            </td>
            """)
        if len(pair) == 1:
            tds.append('<td style="width: 50%; padding: 4px;"></td>')
        apps_rows.append(f"<tr>{''.join(tds)}</tr>")
    apps_table = f'<table style="width: 100%; border-collapse: collapse; margin-bottom: 24px;">{"".join(apps_rows)}</table>'

    top_rows = ""
    for idx, u in enumerate(metrics["top_users"], 1):
        badges = "".join(f'<span style="background:#f8fafc;color:#334155;border:1px solid #e2e8f0;font-size:10px;padding:1px 5px;border-radius:3px;margin:1px;display:inline-block;">{a}</span>' for a in u["apps_used"]) or '<span style="color:#94a3b8;font-size:11px;">None</span>'
        top_rows += f"""
        <tr style="border-bottom:1px solid #f1f5f9;font-size:12px;">
            <td style="padding:9px 10px;font-weight:600;color:#64748b;width:32px;">#{idx}</td>
            <td style="padding:9px 10px;font-weight:600;color:#0f172a;">
                {u['display_name']}
                <div style="font-size:11px;color:#64748b;font-family:ui-monospace,monospace;font-weight:400;">{u['user_principal_name']}</div>
            </td>
            <td style="padding:9px 10px;text-align:right;font-weight:600;color:#0f172a;font-variant-numeric:tabular-nums;">{u.get('prompts_submitted', 0):,}</td>
            <td style="padding:9px 10px;text-align:right;color:#475569;font-variant-numeric:tabular-nums;">{u['total_active_days']} d</td>
            <td style="padding:9px 10px;">{badges}</td>
        </tr>
        """

    inactive_rows = ""
    for u in metrics["inactive_list"][:8]:
        inactive_rows += f"""
        <tr style="border-bottom:1px solid #f8fafc;font-size:12px;">
            <td style="padding:8px 10px;color:#334155;font-weight:500;">{u['display_name']}</td>
            <td style="padding:8px 10px;color:#64748b;font-family:ui-monospace,monospace;font-size:11px;">{u['user_principal_name']}</td>
            <td style="padding:8px 10px;text-align:right;"><span style="background:#fef2f2;color:#991b1b;border:1px solid #fee2e2;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:600;">0 Prompts</span></td>
        </tr>
        """
    if not inactive_rows:
        inactive_rows = '<tr><td colspan="3" style="padding:14px 10px;text-align:center;color:#15803d;font-size:12px;">All licensed accounts have recorded Copilot activity during this period.</td></tr>'

    trend_html = render_trend_svg(metrics.get("trend_points", []))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;background:#f8fafc;margin:0;padding:24px 12px;color:#0f172a;-webkit-font-smoothing:antialiased;">
<div style="max-width:680px;margin:0 auto;background:#ffffff;border-radius:8px;border:1px solid #e2e8f0;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.04);">
    <div style="padding:24px 28px;border-bottom:1px solid #e2e8f0;">
        <div style="font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#64748b;margin-bottom:6px;display:inline-flex;align-items:center;gap:6px;">
            <span style="width:6px;height:6px;border-radius:50%;background:#16a34a;display:inline-block;"></span>
            <span>Live Telemetry &bull; Azure Automation</span>
        </div>
        <h1 style="margin:0 0 4px 0;font-size:20px;font-weight:700;letter-spacing:-0.02em;">Microsoft 365 Copilot Adoption</h1>
        <div style="font-size:12px;color:#64748b;">Reporting Window: <strong>Last 28 Days</strong> &bull; Generated {datetime.now().strftime('%d %b %Y, %H:%M')}</div>
    </div>
    <div style="padding:24px 28px;">
        <table style="width:100%;margin-bottom:24px;text-align:left;border-collapse:separate;border-spacing:6px 0;">
            <tr>
                <td style="padding:12px 10px;background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;width:20%;vertical-align:top;">
                    <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#64748b;">Assigned</div>
                    <div style="font-size:22px;font-weight:700;color:#0f172a;margin-top:4px;font-variant-numeric:tabular-nums;">{metrics['total_users']}</div>
                    <div style="font-size:10px;color:#94a3b8;margin-top:4px;">Seats</div>
                </td>
                <td style="padding:12px 10px;background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;width:20%;vertical-align:top;">
                    <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#64748b;">Active</div>
                    <div style="font-size:22px;font-weight:700;color:#0f172a;margin-top:4px;font-variant-numeric:tabular-nums;">{metrics['active_users']}</div>
                    <div style="font-size:10px;color:#94a3b8;margin-top:4px;">{metrics['adoption_rate']}% rate</div>
                </td>
                <td style="padding:12px 10px;background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;width:20%;vertical-align:top;">
                    <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#64748b;">Adoption</div>
                    <div style="font-size:22px;font-weight:700;color:#0284c7;margin-top:4px;font-variant-numeric:tabular-nums;">{metrics['adoption_rate']}%</div>
                    <div style="font-size:10px;color:#94a3b8;margin-top:4px;">Utilization</div>
                </td>
                <td style="padding:12px 10px;background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;width:20%;vertical-align:top;">
                    <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#64748b;">Prompts</div>
                    <div style="font-size:22px;font-weight:700;color:#0f172a;margin-top:4px;font-variant-numeric:tabular-nums;">{metrics['total_prompts']:,}</div>
                    <div style="font-size:10px;color:#94a3b8;margin-top:4px;">{metrics['avg_prompts']} / user</div>
                </td>
                <td style="padding:12px 10px;background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;width:20%;vertical-align:top;">
                    <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#64748b;">Inactive</div>
                    <div style="font-size:22px;font-weight:700;color:#991b1b;margin-top:4px;font-variant-numeric:tabular-nums;">{metrics['inactive_users']}</div>
                    <div style="font-size:10px;color:#94a3b8;margin-top:4px;">0 activity</div>
                </td>
            </tr>
        </table>

        {f'<div style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;color:#0f172a;margin:26px 0 10px 0;border-bottom:1px solid #e2e8f0;padding-bottom:6px;">Daily Activity & Prompt Momentum</div>{trend_html}' if trend_html else ''}

        <div style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;color:#0f172a;margin:26px 0 10px 0;border-bottom:1px solid #e2e8f0;padding-bottom:6px;">Application Penetration</div>
        {apps_table}

        <div style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;color:#0f172a;margin:30px 0 10px 0;border-bottom:1px solid #e2e8f0;padding-bottom:6px;">Top Active Contributors</div>
        <div style="border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-bottom:24px;">
            <table style="width:100%;border-collapse:collapse;text-align:left;">
                <thead><tr style="background:#f8fafc;font-size:10px;color:#475569;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e2e8f0;"><th style="padding:8px 10px;width:32px;">Rank</th><th style="padding:8px 10px;">User</th><th style="padding:8px 10px;text-align:right;">Prompts</th><th style="padding:8px 10px;text-align:right;">Days</th><th style="padding:8px 10px;">Applications</th></tr></thead>
                <tbody>{top_rows}</tbody>
            </table>
        </div>

        {f'<div style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:0.04em;color:#991b1b;margin:30px 0 10px 0;border-bottom:1px solid #e2e8f0;padding-bottom:6px;">License Optimization & Reallocation Candidates</div><div style="border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-bottom:24px;"><table style="width:100%;border-collapse:collapse;text-align:left;"><thead><tr style="background:#f8fafc;font-size:10px;color:#475569;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e2e8f0;"><th style="padding:8px 10px;">User</th><th style="padding:8px 10px;">Email</th><th style="padding:8px 10px;text-align:right;">Status</th></tr></thead><tbody>{inactive_rows}</tbody></table></div>' if inactive_rows else ''}

        <!-- Attachment Callout Card -->
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:16px;margin-top:24px;">
            <table style="width:100%;border-collapse:collapse;">
                <tr>
                    <td style="width:24px;vertical-align:top;padding-right:10px;line-height:0;">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
                            <rect width="24" height="24" rx="4" fill="#107C41"/>
                            <rect x="10" y="5" width="10" height="14" rx="1" fill="#18A058"/>
                            <rect x="4" y="7" width="9" height="10" rx="1.5" fill="#0B552C"/>
                            <path d="M6.5 10l4 4M10.5 10l-4 4" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>
                        </svg>
                    </td>
                    <td style="vertical-align:top;">
                        <div style="font-size:13px;font-weight:700;color:#0f172a;margin-bottom:4px;">Detailed User Audit Roster Attached</div>
                        <p style="font-size:12px;color:#475569;line-height:1.5;margin:0;">
                            For detailed departmental audits and user-level analysis, the complete roster of all <strong>{metrics['total_users']} licensed accounts</strong> with individual application engagement timestamps is attached as a spreadsheet (<code>copilot_usage_roster_{metrics['period']}.xlsx</code>).
                        </p>
                    </td>
                </tr>
            </table>
        </div>
    </div>
    <div style="background:#f8fafc;padding:16px 28px;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;display:flex;justify-content:space-between;">
        <span>Copilot Usage Automation &bull; Microsoft Graph Telemetry</span>
        <span>Generated on {datetime.now().strftime('%d %b %Y, %H:%M')}</span>
    </div>
</div>
</body>
</html>"""


def build_audit_attachment(users: List[Dict[str, Any]], period: str = "D28"):
    """
    Generates an Excel roster (if openpyxl available) or CSV fallback in memory.
    Returns (filename, contentType, base64_content).
    """
    import base64
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Copilot User Roster"
        ws.views.sheetView[0].showGridLines = True

        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        regular_font = Font(name="Segoe UI", size=10, color="0F172A")
        bold_font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
        muted_font = Font(name="Segoe UI", size=9, color="64748B")
        active_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        active_font = Font(name="Segoe UI", size=9, bold=True, color="166534")
        inactive_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        inactive_font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

        headers = ["Rank", "Display Name", "Email (UPN)", "Status", "Total Prompts", "Active Days", "Last Activity", "Teams", "Chat", "Outlook", "Word", "Excel", "PowerPoint", "Loop", "OneNote"]
        ws.append(headers)
        ws.row_dimensions[1].height = 24
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        sorted_users = sorted(users, key=lambda x: (x["is_active"], x.get("prompts_submitted", 0), x["total_active_days"]), reverse=True)
        for idx, u in enumerate(sorted_users, 1):
            act = u.get("app_last_activity", {})
            ws.append([
                idx, u["display_name"], u["user_principal_name"],
                "Active" if u["is_active"] else "Inactive",
                u.get("prompts_submitted", 0), u["total_active_days"],
                u["last_activity_date"] or "-",
                act.get("Teams") or "-", act.get("Chat") or "-", act.get("Outlook") or "-",
                act.get("Word") or "-", act.get("Excel") or "-", act.get("PowerPoint") or "-",
                act.get("Loop") or "-", act.get("OneNote") or "-",
            ])
            r_idx = idx + 1
            ws.row_dimensions[r_idx].height = 20
            for c_idx, c in enumerate(ws[r_idx], 1):
                c.font = regular_font
                if c_idx == 4:
                    c.fill = active_fill if u["is_active"] else inactive_fill
                    c.font = active_font if u["is_active"] else inactive_font
                elif c_idx == 5 and u.get("prompts_submitted", 0) > 0:
                    c.font = bold_font

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

        ws.freeze_panes = "A2"
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        b64 = base64.b64encode(buf.read()).decode("utf-8")
        return f"copilot_usage_roster_{period}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", b64
    except ImportError:
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(["Rank", "Display Name", "Email", "Status", "Prompts", "Active Days", "Last Activity", "Applications"])
        sorted_users = sorted(users, key=lambda x: (x["is_active"], x.get("prompts_submitted", 0), x["total_active_days"]), reverse=True)
        for idx, u in enumerate(sorted_users, 1):
            writer.writerow([idx, u["display_name"], u["user_principal_name"], "Active" if u["is_active"] else "Inactive", u.get("prompts_submitted", 0), u["total_active_days"], u["last_activity_date"], ", ".join(u.get("apps_used", []))])
        b64 = base64.b64encode(buf.getvalue().encode("utf-8")).decode("utf-8")
        return f"copilot_usage_roster_{period}.csv", "text/csv", b64


def send_email(token: str, html_body: str, sender: str = None, recipients: List[str] = None, subject: str = None, attachment: Optional[tuple] = None):
    sender = sender or MAIL_SENDER
    recipients = recipients or MAIL_RECIPIENTS
    subject = subject or get_default_subject()

    if not sender or not recipients:
        logger.warning("Sender or recipients not configured; skipping email dispatch.")
        return

    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    msg = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [{"emailAddress": {"address": r}} for r in recipients],
    }
    if attachment:
        fname, ctype, b64content = attachment
        msg["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": fname,
            "contentType": ctype,
            "contentBytes": b64content,
        }]

    payload = {"message": msg, "saveToSentItems": "true"}
    logger.info(f"Dispatching email: {sender} -> {recipients}")
    res = requests.post(url, headers=headers, json=payload, timeout=40)
    if res.status_code in (200, 202):
        logger.info("Report email with audit attachment dispatched successfully.")
    else:
        res.raise_for_status()


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Azure Automation Copilot Usage Runbook")
    parser.add_argument("--period", default=REPORT_PERIOD, help="Report period (D7, D28, D90, D180)")
    parser.add_argument("--sender", default=MAIL_SENDER, help="Sender mailbox UPN")
    parser.add_argument("--recipients", default=",".join(MAIL_RECIPIENTS), help="Comma-separated recipient emails")
    args, _ = parser.parse_known_args()

    period = args.period or REPORT_PERIOD
    sender = args.sender or MAIL_SENDER
    recipients = [r.strip() for r in (args.recipients or "").split(",") if r.strip()] or MAIL_RECIPIENTS
    subject = get_default_subject()

    logger.info("=== Azure Automation Copilot Usage Runbook Initialized ===")
    logger.info("Authentication Mode: Azure System-Assigned Managed Identity (Credential-free)")
    logger.info(f"Report Period: {period}")

    token = get_token()
    users = fetch_report(token, period=period)
    trend = fetch_trend(token, period=period)
    logger.info(f"Retrieved {len(users)} user records and {len(trend)} trend points.")

    metrics = aggregate_metrics(users, period=period, trend=trend)
    html_content = generate_html(metrics)
    attachment = build_audit_attachment(users, period=period)

    if sender and recipients:
        send_email(token, html_content, sender=sender, recipients=recipients, subject=subject, attachment=attachment)
    else:
        logger.info("Dry-run / Preview Mode: Sender or recipients not configured. Email dispatch skipped.")
    logger.info("=== Runbook Execution Completed ===")


if __name__ == "__main__":
    main()
