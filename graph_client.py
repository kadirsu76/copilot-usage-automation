"""
Microsoft Graph API Client for Copilot Usage & Trend Reports.
Handles endpoint requests, redirects, JSON/CSV parsing, and data normalization.
Supports Microsoft Graph Copilot Reporting API (v2 schema) and Daily Trends.
"""

import io
import csv
import logging
from typing import List, Dict, Any, Optional
import requests
import config

logger = logging.getLogger(__name__)


class GraphReportError(Exception):
    """Exception raised when Graph reports fail to fetch or parse."""
    pass


def fetch_copilot_usage_user_detail(token: str, period: str = "D28") -> List[Dict[str, Any]]:
    """
    Fetches user-level detail for Microsoft 365 Copilot usage.
    Endpoint: /beta/reports/getMicrosoft365CopilotUsageUserDetail(period='{period}')
    
    Supported periods in Graph v2: D7, D28, D90, D180.
    (If D30 is provided, it is automatically converted to D28).
    """
    period_upper = period.upper()
    if period_upper in ["D30", ""]:
        period_upper = "D28"

    valid_periods = ["D7", "D28", "D90", "D180"]
    if period_upper not in valid_periods:
        logger.warning(f"Invalid period '{period}', defaulting to 'D28'.")
        period_upper = "D28"

    url = f"{config.GRAPH_API_BETA}/reports/getMicrosoft365CopilotUsageUserDetail(period='{period_upper}')"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    logger.info(f"Fetching Copilot user details: {url}")
    try:
        response = requests.get(url, headers=headers, allow_redirects=True, timeout=60)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        raw_text = response.text.strip()

        if "application/json" in content_type or raw_text.startswith("{") or raw_text.startswith("["):
            try:
                data = response.json()
                if isinstance(data, dict) and "value" in data:
                    raw_records = data["value"]
                elif isinstance(data, list):
                    raw_records = data
                else:
                    raw_records = [data]
                logger.info(f"Retrieved {len(raw_records)} user records from Graph API.")
                return [_normalize_record(r) for r in raw_records]
            except Exception as json_err:
                logger.debug(f"JSON parsing failed, attempting CSV: {json_err}")

        reader = csv.DictReader(io.StringIO(raw_text))
        records = [_normalize_record(row) for row in reader]
        logger.info(f"Parsed {len(records)} user records from CSV.")
        return records

    except requests.exceptions.HTTPError as e:
        status_code = response.status_code if "response" in locals() else None
        err_msg = response.text if "response" in locals() else str(e)
        if status_code == 403:
            raise GraphReportError(
                "Graph API 403 Forbidden! Ensure the application has 'Reports.Read.All' permission "
                "with admin consent granted."
            ) from e
        raise GraphReportError(f"Failed to fetch report (HTTP {status_code}): {err_msg}") from e
    except requests.exceptions.RequestException as e:
        raise GraphReportError(f"Graph API network error: {e}") from e


def fetch_copilot_user_count_trend(token: str, period: str = "D28") -> List[Dict[str, Any]]:
    """
    Fetches daily trend data for Copilot active users and prompt volume.
    Endpoint: /beta/reports/getMicrosoft365CopilotUserCountTrend(period='{period}')
    """
    period_upper = period.upper()
    if period_upper in ["D30", ""]:
        period_upper = "D28"

    valid_periods = ["D7", "D28", "D90", "D180"]
    if period_upper not in valid_periods:
        period_upper = "D28"

    url = f"{config.GRAPH_API_BETA}/reports/getMicrosoft365CopilotUserCountTrend(period='{period_upper}')"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    logger.info(f"Fetching Copilot daily trend: {url}")
    try:
        response = requests.get(url, headers=headers, allow_redirects=True, timeout=60)
        response.raise_for_status()

        data = response.json()
        val = data.get("value", [])
        if not val:
            logger.warning("Trend response contained empty value array.")
            return []

        adoption_by_date = val[0].get("adoptionByDate", [])
        # Sort chronologically ascending
        adoption_by_date.sort(key=lambda x: x.get("reportDate", ""))
        logger.info(f"Retrieved {len(adoption_by_date)} daily trend data points.")
        return adoption_by_date
    except Exception as e:
        logger.warning(f"Could not retrieve trend data: {e}. Report will proceed without trend.")
        return []


def _normalize_record(raw: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalizes both v2 JSON schema and legacy CSV into a uniform dictionary schema.
    """
    def get_val(*candidates: str) -> Optional[str]:
        for c in candidates:
            if c in raw and raw[c] is not None:
                val = str(raw[c]).strip()
                if val != "":
                    return val
            for k, v in raw.items():
                if k.lower().replace(" ", "").replace("_", "") == c.lower().replace(" ", "").replace("_", ""):
                    if v is not None:
                        val = str(v).strip()
                        if val != "":
                            return val
        return None

    def get_int(*candidates: str) -> int:
        val = get_val(*candidates)
        if val is None:
            return 0
        try:
            return int(float(val))
        except (ValueError, TypeError):
            return 0

    upn = get_val("userPrincipalName", "User Principal Name", "upn") or "Unknown"
    display_name = get_val("displayName", "Display Name", "name") or upn.split("@")[0]
    report_refresh_date = get_val("reportRefreshDate", "Report Refresh Date") or ""
    last_activity_date = get_val("lastActivityDate", "Last Activity Date") or ""

    period_info = {}
    if "copilotActivityUserDetailsByPeriod" in raw and isinstance(raw["copilotActivityUserDetailsByPeriod"], list):
        if len(raw["copilotActivityUserDetailsByPeriod"]) > 0:
            period_info = raw["copilotActivityUserDetailsByPeriod"][0] or {}

    total_active_days = period_info.get("activeUsageDaysAllApps")
    if total_active_days is None:
        total_active_days = get_int("Copilot Chat Active Days", "activeUsageDaysAllApps", "copilotChatActiveDays")
    else:
        total_active_days = int(total_active_days)

    prompts_submitted = period_info.get("promptsSubmittedAllApps", 0) or 0
    prompts_chat_web = period_info.get("promptsSubmittedBusinessChatWeb", 0) or 0
    prompts_chat_work = period_info.get("promptsSubmittedBusinessChatWork", 0) or 0

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
    is_active = bool(last_activity_date or total_active_days > 0 or prompts_submitted > 0)

    return {
        "user_principal_name": upn,
        "display_name": display_name,
        "report_refresh_date": report_refresh_date,
        "last_activity_date": last_activity_date,
        "is_active": is_active,
        "apps_used": apps_used,
        "total_active_days": total_active_days,
        "prompts_submitted": prompts_submitted,
        "prompts_chat_web": prompts_chat_web,
        "prompts_chat_work": prompts_chat_work,
        "app_last_activity": app_last_activity,
    }
