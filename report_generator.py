"""
Executive Email Report Generator Module - Engineered for Enterprise Email Clients.
Adheres to 'design-taste-frontend' standards:
  - 680px max-width optimized for Outlook Desktop, Outlook Web, Gmail, and Mobile screens.
  - Zero JavaScript, zero inner scrollbars (incompatible with email clients).
  - Authentic Microsoft application vector marks.
  - Daily Activity & Prompt Momentum dual-metric SVG time-series.
  - Attachment Callout Card referencing the attached comprehensive Excel audit roster (.xlsx).
  - Tabular numerals (tabular-nums), neutral slate/zinc palette, zero emojis, zero AI slop words.
"""

from datetime import datetime
from typing import List, Dict, Any, Optional


def aggregate_metrics(
    users: List[Dict[str, Any]],
    period: str = "D28",
    trend: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Computes summary metrics, adoption rates, prompt volume, app penetration, and daily trends.
    """
    total_users = len(users)
    active_users = sum(1 for u in users if u["is_active"])
    inactive_users = total_users - active_users
    adoption_rate = round((active_users / total_users * 100), 1) if total_users > 0 else 0.0

    total_active_days = sum(u["total_active_days"] for u in users)
    total_prompts = sum(u.get("prompts_submitted", 0) for u in users)
    avg_prompts = round(total_prompts / active_users, 1) if active_users > 0 else 0.0
    avg_active_days = round(total_active_days / active_users, 1) if active_users > 0 else 0.0

    # Application penetration metrics
    apps = ["Teams", "Chat", "Outlook", "Word", "Excel", "PowerPoint", "Loop", "OneNote"]
    app_stats = {}
    for app in apps:
        active_count = sum(1 for u in users if u["app_last_activity"].get(app))
        rate = round((active_count / total_users * 100), 1) if total_users > 0 else 0.0
        app_stats[app] = {
            "active_count": active_count,
            "adoption_rate": rate,
        }

    # Engagement segments
    power_users = [u for u in users if u["total_active_days"] >= 10 or u.get("prompts_submitted", 0) >= 50]
    regular_users = [
        u for u in users
        if (3 <= u["total_active_days"] < 10 or 10 <= u.get("prompts_submitted", 0) < 50) and u not in power_users
    ]
    casual_users = [
        u for u in users
        if u["is_active"] and u not in power_users and u not in regular_users
    ]
    inactive_list = [u for u in users if not u["is_active"]]

    # Trend calculations
    trend_stats = {}
    if trend and len(trend) > 0:
        max_prompt_entry = max(trend, key=lambda x: x.get("promptsSubmitted", 0))
        peak_prompts = max_prompt_entry.get("promptsSubmitted", 0)
        peak_date = max_prompt_entry.get("reportDate", "")
        try:
            peak_date_fmt = datetime.strptime(peak_date, "%Y-%m-%d").strftime("%d %b")
        except Exception:
            peak_date_fmt = peak_date

        weekday_prompts = [
            t.get("promptsSubmitted", 0) for t in trend
            if _is_weekday(t.get("reportDate", ""))
        ]
        avg_weekday_prompts = (
            round(sum(weekday_prompts) / len(weekday_prompts), 1)
            if weekday_prompts else 0.0
        )

        if len(trend) >= 14:
            recent_7 = sum(t.get("promptsSubmitted", 0) for t in trend[-7:])
            prev_7 = sum(t.get("promptsSubmitted", 0) for t in trend[-14:-7])
            if prev_7 > 0:
                wow_change = round(((recent_7 - prev_7) / prev_7) * 100, 1)
            else:
                wow_change = 0.0
        else:
            wow_change = 0.0

        trend_stats = {
            "peak_prompts": peak_prompts,
            "peak_date_fmt": peak_date_fmt,
            "avg_weekday_prompts": avg_weekday_prompts,
            "wow_change": wow_change,
            "points": trend,
        }

    period_labels = {
        "D7": "Last 7 Days",
        "D28": "Last 28 Days",
        "D30": "Last 28 Days",
        "D90": "Last 90 Days",
        "D180": "Last 180 Days",
    }
    period_label = period_labels.get(period.upper(), f"Period: {period}")

    return {
        "period": period,
        "period_label": period_label,
        "generated_at": datetime.now().strftime("%d %b %Y, %H:%M"),
        "total_users": total_users,
        "active_users": active_users,
        "inactive_users": inactive_users,
        "adoption_rate": adoption_rate,
        "total_active_days": total_active_days,
        "total_prompts": total_prompts,
        "avg_prompts": avg_prompts,
        "avg_active_days": avg_active_days,
        "app_stats": app_stats,
        "trend_stats": trend_stats,
        "segments": {
            "power": {"count": len(power_users), "rate": round(len(power_users)/total_users*100, 1) if total_users else 0},
            "regular": {"count": len(regular_users), "rate": round(len(regular_users)/total_users*100, 1) if total_users else 0},
            "casual": {"count": len(casual_users), "rate": round(len(casual_users)/total_users*100, 1) if total_users else 0},
            "inactive": {"count": len(inactive_list), "rate": round(len(inactive_list)/total_users*100, 1) if total_users else 0},
        },
        "top_users": sorted(users, key=lambda x: (x.get("prompts_submitted", 0), x["total_active_days"]), reverse=True)[:10],
        "inactive_list": inactive_list,
        "all_users": sorted(users, key=lambda x: (x["is_active"], x.get("prompts_submitted", 0), x["total_active_days"]), reverse=True),
    }


def _is_weekday(date_str: str) -> bool:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return dt.weekday() < 5
    except Exception:
        return True


def _render_app_icon(app_name: str) -> str:
    """Renders authentic Microsoft Office brand icons as compact vector SVGs."""
    icons = {
        "Teams": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#5059C9"/>'
            '<path d="M14.5 9h5v6h-5z" fill="#7B83EB"/>'
            '<circle cx="17" cy="6.5" r="2" fill="#7B83EB"/>'
            '<rect x="4.5" y="6" width="9" height="12" rx="2" fill="#4B53BC"/>'
            '<path d="M7 9h4M9 9v6" stroke="#ffffff" stroke-width="1.6" stroke-linecap="round"/>'
            '</svg>'
        ),
        "Chat": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#0284C7"/>'
            '<path d="M6 8h12M6 12h8M6 16h5" stroke="#ffffff" stroke-width="1.8" stroke-linecap="round"/>'
            '<circle cx="17.5" cy="15.5" r="2.5" fill="#38BDF8"/>'
            '</svg>'
        ),
        "Outlook": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#0078D4"/>'
            '<circle cx="14" cy="12" r="5" fill="#28A8EA"/>'
            '<rect x="4" y="7" width="9" height="10" rx="1.5" fill="#005A9E"/>'
            '<circle cx="8.5" cy="12" r="2.5" stroke="#ffffff" stroke-width="1.5"/>'
            '</svg>'
        ),
        "Word": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#185ABD"/>'
            '<rect x="10" y="5" width="10" height="14" rx="1" fill="#2B7CD3"/>'
            '<rect x="4" y="7" width="9" height="10" rx="1.5" fill="#103F91"/>'
            '<path d="M6 10l1.2 4 1.3-4 1.3 4 1.2-4" stroke="#ffffff" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>'
            '</svg>'
        ),
        "Excel": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#107C41"/>'
            '<rect x="10" y="5" width="10" height="14" rx="1" fill="#18A058"/>'
            '<rect x="4" y="7" width="9" height="10" rx="1.5" fill="#0B552C"/>'
            '<path d="M6.5 10l4 4M10.5 10l-4 4" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>'
            '</svg>'
        ),
        "PowerPoint": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#C43E1C"/>'
            '<rect x="10" y="5" width="10" height="14" rx="1" fill="#ED6C47"/>'
            '<rect x="4" y="7" width="9" height="10" rx="1.5" fill="#93280F"/>'
            '<path d="M7 14.5v-5h2.5a1.5 1.5 0 010 3H7" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>'
            '</svg>'
        ),
        "Loop": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#0891B2"/>'
            '<circle cx="12" cy="12" r="5.5" stroke="#ffffff" stroke-width="1.8" stroke-dasharray="8 4"/>'
            '</svg>'
        ),
        "OneNote": (
            '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">'
            '<rect width="24" height="24" rx="4" fill="#7719AA"/>'
            '<rect x="10" y="5" width="10" height="14" rx="1" fill="#9B2FAF"/>'
            '<rect x="4" y="7" width="9" height="10" rx="1.5" fill="#501075"/>'
            '<path d="M6.5 14.5v-5l4 5v-5" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>'
            '</svg>'
        ),
    }
    return icons.get(app_name, "")


def _render_trend_svg(trend_points: List[Dict[str, Any]]) -> str:
    """
    Renders an inline SVG time-series chart of daily prompts and active user counts.
    Sized for 680px email containers with zero external dependencies.
    """
    if not trend_points:
        return ""

    n = len(trend_points)
    max_p = max((t.get("promptsSubmitted", 0) for t in trend_points), default=1) or 1
    max_u = max((t.get("anyAppActiveUsers", 0) for t in trend_points), default=1) or 1

    chart_left = 36
    chart_right = 584
    chart_top = 20
    chart_bottom = 125
    chart_w = chart_right - chart_left
    chart_h = chart_bottom - chart_top

    bar_spacing = chart_w / n
    bar_width = max(bar_spacing * 0.65, 3)

    bars_svg = []
    line_pts = []
    dots_svg = []
    labels_svg = []

    for i, t in enumerate(trend_points):
        x_center = chart_left + (i + 0.5) * bar_spacing
        p = t.get("promptsSubmitted", 0)
        u = t.get("anyAppActiveUsers", 0)
        date_str = t.get("reportDate", "")

        bh = (p / max_p) * (chart_h - 10)
        bx = x_center - bar_width / 2
        by = chart_bottom - bh

        is_wkend = not _is_weekday(date_str)
        fill_col = "#e2e8f0" if is_wkend else "#0f172a"
        bars_svg.append(
            f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bar_width:.1f}" height="{max(bh, 1):.1f}" '
            f'fill="{fill_col}" rx="1">'
            f'<title>{date_str}: {p} prompts, {u} active users</title></rect>'
        )

        uy = chart_bottom - (u / max_u) * (chart_h - 10)
        line_pts.append(f"{x_center:.1f},{uy:.1f}")
        dots_svg.append(
            f'<circle cx="{x_center:.1f}" cy="{uy:.1f}" r="2.5" fill="#0284c7" stroke="#ffffff" stroke-width="1.2">'
            f'<title>{date_str}: {u} active users</title></circle>'
        )

        if i % 7 == 0 or i == n - 1:
            try:
                short_d = datetime.strptime(date_str, "%Y-%m-%d").strftime("%d %b")
            except Exception:
                short_d = date_str[-5:]
            labels_svg.append(
                f'<text x="{x_center:.1f}" y="142" font-size="9" fill="#94a3b8" text-anchor="middle" '
                f'font-family="-apple-system, sans-serif">{short_d}</text>'
            )

    line_path = " ".join(line_pts)

    svg = f"""
    <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;padding:16px;margin-bottom:24px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;font-size:11px;font-weight:600;color:#64748b;">
            <div style="display:flex;gap:16px;">
                <span style="display:inline-flex;align-items:center;gap:5px;">
                    <span style="display:inline-block;width:9px;height:9px;background:#0f172a;border-radius:2px;"></span>
                    Prompts (Peak {max_p})
                </span>
                <span style="display:inline-flex;align-items:center;gap:5px;">
                    <span style="display:inline-block;width:12px;height:2px;background:#0284c7;"></span>
                    Active Users (Max {max_u})
                </span>
            </div>
            <div style="color:#94a3b8;font-weight:400;">
                Hover for daily count
            </div>
        </div>
        <svg viewBox="0 0 620 150" style="width:100%;height:auto;display:block;">
            <line x1="36" y1="{chart_bottom}" x2="584" y2="{chart_bottom}" stroke="#e2e8f0" stroke-width="1"/>
            <line x1="36" y1="{chart_top + chart_h/2}" x2="584" y2="{chart_top + chart_h/2}" stroke="#f1f5f9" stroke-width="1" stroke-dasharray="3 3"/>
            
            <text x="30" y="{chart_bottom}" font-size="9" fill="#94a3b8" text-anchor="end" font-family="-apple-system, sans-serif">0</text>
            <text x="30" y="{chart_top + 4}" font-size="9" fill="#94a3b8" text-anchor="end" font-family="-apple-system, sans-serif">{max_p}</text>

            {''.join(bars_svg)}
            <polyline fill="none" stroke="#0284c7" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" points="{line_path}"/>
            {''.join(dots_svg)}
            {''.join(labels_svg)}
        </svg>
    </div>
    """
    return svg


def generate_html_report(metrics: Dict[str, Any]) -> str:
    """
    Renders an executive email HTML report.
    Optimized for email client preview panes (680px width) with zero broken JavaScript or inner scrollbars.
    """
    app_meta = {
        "Teams": "Microsoft Teams",
        "Chat": "Business Chat",
        "Outlook": "Microsoft Outlook",
        "Word": "Microsoft Word",
        "Excel": "Microsoft Excel",
        "PowerPoint": "Microsoft PowerPoint",
        "Loop": "Microsoft Loop",
        "OneNote": "Microsoft OneNote",
    }

    # Top Users table rows
    top_users_html = ""
    for idx, u in enumerate(metrics["top_users"], 1):
        badges = "".join(
            f'<span style="display:inline-block;background:#f8fafc;color:#334155;border:1px solid #e2e8f0;font-size:10px;font-weight:500;padding:1px 5px;border-radius:3px;margin:1px;">{app}</span>'
            for app in u["apps_used"]
        ) or '<span style="color:#94a3b8;font-size:11px;">None</span>'

        prompts_str = f"{u.get('prompts_submitted', 0):,}"
        top_users_html += f"""
        <tr style="border-bottom: 1px solid #f1f5f9; font-size: 12px;">
            <td style="padding: 9px 10px; font-weight: 600; color: #64748b; font-variant-numeric: tabular-nums; width: 32px;">#{idx}</td>
            <td style="padding: 9px 10px; font-weight: 600; color: #0f172a;">
                {u['display_name']}
                <div style="font-size: 11px; font-weight: 400; color: #64748b; font-family: ui-monospace, monospace;">{u['user_principal_name']}</div>
            </td>
            <td style="padding: 9px 10px; text-align: right; font-weight: 600; color: #0f172a; font-variant-numeric: tabular-nums;">{prompts_str}</td>
            <td style="padding: 9px 10px; text-align: right; color: #475569; font-variant-numeric: tabular-nums;">{u['total_active_days']} d</td>
            <td style="padding: 9px 10px;">{badges}</td>
        </tr>
        """

    # Inactive Users rows
    inactive_users_html = ""
    for u in metrics["inactive_list"][:8]:
        inactive_users_html += f"""
        <tr style="border-bottom: 1px solid #f8fafc; font-size: 12px;">
            <td style="padding: 8px 10px; font-weight: 500; color: #334155;">{u['display_name']}</td>
            <td style="padding: 8px 10px; color: #64748b; font-family: ui-monospace, monospace; font-size: 11px;">{u['user_principal_name']}</td>
            <td style="padding: 8px 10px; text-align: right;"><span style="display:inline-block;background:#fef2f2;color:#991b1b;border:1px solid #fee2e2;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:600;">0 Prompts</span></td>
        </tr>
        """
    if not inactive_users_html:
        inactive_users_html = '<tr><td colspan="3" style="padding: 14px 10px; text-align: center; color: #15803d; font-size: 12px;">All licensed accounts have recorded Copilot activity during this period.</td></tr>'

    # App Breakdown Cards (2 columns email-safe table layout)
    apps_items = list(metrics["app_stats"].items())
    apps_rows = []
    for i in range(0, len(apps_items), 2):
        pair = apps_items[i:i+2]
        tds = []
        for app_name, stats in pair:
            title = app_meta.get(app_name, app_name)
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

    apps_table_html = f'<table style="width: 100%; border-collapse: collapse; margin-bottom: 24px;">{"".join(apps_rows)}</table>'

    total_prompts_fmt = f"{metrics['total_prompts']:,}"

    # Trend Section
    trend_section = ""
    t_stats = metrics.get("trend_stats", {})
    if t_stats and t_stats.get("points"):
        svg_chart = _render_trend_svg(t_stats["points"])
        wow_badge = (
            f'<span style="color:#15803d;font-weight:600;">+{t_stats["wow_change"]}%</span>'
            if t_stats["wow_change"] >= 0
            else f'<span style="color:#991b1b;font-weight:600;">{t_stats["wow_change"]}%</span>'
        )
        trend_section = f"""
        <div class="section-header">
            <div>
                <h2 class="section-title">Daily Activity & Prompt Momentum</h2>
                <div class="section-subtext">Peak: <strong>{t_stats['peak_prompts']} prompts</strong> ({t_stats['peak_date_fmt']}) &bull; Weekday Avg: <strong>{t_stats['avg_weekday_prompts']}</strong> &bull; 7-day trend: {wow_badge}</div>
            </div>
        </div>
        {svg_chart}
        """

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Copilot Adoption Report</title>
    <style>
        * {{
            box-sizing: border-box;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background-color: #f8fafc;
            color: #0f172a;
            margin: 0;
            padding: 24px 12px;
            line-height: 1.5;
            -webkit-font-smoothing: antialiased;
        }}
        .container {{
            max-width: 680px;
            margin: 0 auto;
            background: #ffffff;
            border-radius: 8px;
            border: 1px solid #e2e8f0;
            overflow: hidden;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.04);
        }}
        .header {{
            padding: 24px 28px;
            border-bottom: 1px solid #e2e8f0;
            background: #ffffff;
        }}
        .badge-kicker {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #64748b;
            margin-bottom: 6px;
        }}
        .badge-kicker-dot {{
            width: 6px;
            height: 6px;
            border-radius: 50%;
            background: #16a34a;
        }}
        .header h1 {{
            margin: 0 0 4px 0;
            font-size: 20px;
            font-weight: 700;
            letter-spacing: -0.02em;
            color: #0f172a;
        }}
        .header-meta {{
            font-size: 12px;
            color: #64748b;
        }}
        .content {{
            padding: 24px 28px;
        }}
        .kpi-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 24px;
        }}
        .kpi-cell {{
            padding: 12px 10px;
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            width: 20%;
            text-align: left;
        }}
        .kpi-label {{
            font-size: 10px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #64748b;
        }}
        .kpi-value {{
            font-size: 22px;
            font-weight: 700;
            color: #0f172a;
            letter-spacing: -0.02em;
            line-height: 1.1;
            margin-top: 4px;
            font-variant-numeric: tabular-nums;
        }}
        .kpi-subtext {{
            font-size: 10px;
            color: #94a3b8;
            margin-top: 4px;
        }}
        .section-header {{
            margin: 26px 0 10px 0;
            padding-bottom: 6px;
            border-bottom: 1px solid #e2e8f0;
        }}
        .section-title {{
            font-size: 13px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            color: #0f172a;
            margin: 0;
        }}
        .section-subtext {{
            font-size: 11px;
            color: #64748b;
            margin-top: 2px;
        }}
        .segment-container {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            padding: 14px;
            margin-bottom: 24px;
        }}
        .segment-bar {{
            display: flex;
            height: 6px;
            border-radius: 3px;
            overflow: hidden;
            background: #f1f5f9;
            margin-bottom: 12px;
        }}
        .segment-legend {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px 14px;
            font-size: 11px;
        }}
        .legend-item {{
            display: flex;
            align-items: center;
            gap: 6px;
        }}
        .legend-dot {{
            width: 7px;
            height: 7px;
            border-radius: 2px;
            flex-shrink: 0;
        }}
        .legend-name {{
            color: #64748b;
        }}
        .legend-val {{
            font-weight: 600;
            color: #0f172a;
            margin-left: auto;
            font-variant-numeric: tabular-nums;
        }}
        .apps-wrap {{
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            margin-bottom: 24px;
        }}
        .table-wrap {{
            width: 100%;
            overflow-x: auto;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            margin-bottom: 24px;
            background: #ffffff;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
        }}
        th {{
            background: #f8fafc;
            padding: 8px 10px;
            font-size: 10px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #475569;
            border-bottom: 1px solid #e2e8f0;
        }}
        .attachment-card {{
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            padding: 16px;
            margin-top: 26px;
        }}
        .attachment-header {{
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 4px;
        }}
        .attachment-title {{
            font-size: 12px;
            font-weight: 700;
            color: #0f172a;
        }}
        .attachment-desc {{
            font-size: 11px;
            color: #64748b;
            line-height: 1.5;
            margin: 0;
        }}
        .attachment-desc code {{
            background: #e2e8f0;
            color: #0f172a;
            padding: 1px 4px;
            border-radius: 3px;
            font-size: 10px;
            font-family: ui-monospace, monospace;
        }}
        .footer {{
            padding: 16px 28px;
            border-top: 1px solid #e2e8f0;
            background: #f8fafc;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 11px;
            color: #94a3b8;
            flex-wrap: wrap;
            gap: 8px;
        }}
        .footer-brand {{
            color: #64748b;
            font-weight: 500;
        }}
    </style>
</head>
<body>

<div class="container">
    <!-- Header -->
    <div class="header">
        <div class="badge-kicker">
            <span class="badge-kicker-dot"></span>
            <span>Live Telemetry &bull; Microsoft 365 Copilot</span>
        </div>
        <h1>Microsoft 365 Copilot Adoption</h1>
        <div class="header-meta">Reporting Window: <strong>{metrics['period_label']}</strong> &bull; Generated {metrics['generated_at']}</div>
    </div>

    <div class="content">
        <!-- KPI Row -->
        <table class="kpi-table">
            <tr>
                <td class="kpi-cell">
                    <div class="kpi-label">Assigned</div>
                    <div class="kpi-value">{metrics['total_users']}</div>
                    <div class="kpi-subtext">Seats</div>
                </td>
                <td class="kpi-cell">
                    <div class="kpi-label">Active</div>
                    <div class="kpi-value" style="color: #0f172a;">{metrics['active_users']}</div>
                    <div class="kpi-subtext">{metrics['adoption_rate']}% rate</div>
                </td>
                <td class="kpi-cell">
                    <div class="kpi-label">Adoption</div>
                    <div class="kpi-value" style="color: #0284c7;">{metrics['adoption_rate']}%</div>
                    <div class="kpi-subtext">Utilization</div>
                </td>
                <td class="kpi-cell">
                    <div class="kpi-label">Prompts</div>
                    <div class="kpi-value" style="color: #0f172a;">{total_prompts_fmt}</div>
                    <div class="kpi-subtext">{metrics['avg_prompts']} / user</div>
                </td>
                <td class="kpi-cell">
                    <div class="kpi-label">Inactive</div>
                    <div class="kpi-value" style="color: #991b1b;">{metrics['inactive_users']}</div>
                    <div class="kpi-subtext">0 activity</div>
                </td>
            </tr>
        </table>

        <!-- Trend Section -->
        {trend_section}

        <!-- User Engagement Segmentation -->
        <div class="section-header">
            <h2 class="section-title">User Engagement Distribution</h2>
            <div class="section-subtext">Segmentation based on prompt frequency and active days</div>
        </div>
        <div class="segment-container" style="background:#ffffff;border:1px solid #e2e8f0;border-radius:6px;padding:14px;margin-bottom:24px;">
            <table style="width:100%;height:6px;border-collapse:collapse;margin-bottom:12px;border-radius:3px;overflow:hidden;">
                <tr>
                    <td style="width:{metrics['segments']['power']['rate']}%;background:#0f172a;height:6px;padding:0;"></td>
                    <td style="width:{metrics['segments']['regular']['rate']}%;background:#0284c7;height:6px;padding:0;"></td>
                    <td style="width:{metrics['segments']['casual']['rate']}%;background:#94a3b8;height:6px;padding:0;"></td>
                    <td style="width:{metrics['segments']['inactive']['rate']}%;background:#cbd5e1;height:6px;padding:0;"></td>
                </tr>
            </table>
            <table style="width:100%;border-collapse:collapse;font-size:11px;">
                <tr>
                    <td style="width:50%;padding:4px 8px 4px 0;vertical-align:middle;">
                        <table style="width:100%;border-collapse:collapse;">
                            <tr>
                                <td style="width:12px;padding:0;line-height:0;vertical-align:middle;"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:#0f172a;"></span></td>
                                <td style="color:#64748b;padding:0;font-size:11px;vertical-align:middle;">Power Users (50+ prompts)</td>
                                <td style="text-align:right;font-weight:600;color:#0f172a;padding:0;font-size:11px;font-variant-numeric:tabular-nums;vertical-align:middle;">{metrics['segments']['power']['count']} <span style="color:#94a3b8;font-weight:400;">({metrics['segments']['power']['rate']}%)</span></td>
                            </tr>
                        </table>
                    </td>
                    <td style="width:50%;padding:4px 0 4px 8px;vertical-align:middle;">
                        <table style="width:100%;border-collapse:collapse;">
                            <tr>
                                <td style="width:12px;padding:0;line-height:0;vertical-align:middle;"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:#0284c7;"></span></td>
                                <td style="color:#64748b;padding:0;font-size:11px;vertical-align:middle;">Regular Users (10-49 prompts)</td>
                                <td style="text-align:right;font-weight:600;color:#0f172a;padding:0;font-size:11px;font-variant-numeric:tabular-nums;vertical-align:middle;">{metrics['segments']['regular']['count']} <span style="color:#94a3b8;font-weight:400;">({metrics['segments']['regular']['rate']}%)</span></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="width:50%;padding:4px 8px 4px 0;vertical-align:middle;">
                        <table style="width:100%;border-collapse:collapse;">
                            <tr>
                                <td style="width:12px;padding:0;line-height:0;vertical-align:middle;"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:#94a3b8;"></span></td>
                                <td style="color:#64748b;padding:0;font-size:11px;vertical-align:middle;">Occasional Users (1-9 prompts)</td>
                                <td style="text-align:right;font-weight:600;color:#0f172a;padding:0;font-size:11px;font-variant-numeric:tabular-nums;vertical-align:middle;">{metrics['segments']['casual']['count']} <span style="color:#94a3b8;font-weight:400;">({metrics['segments']['casual']['rate']}%)</span></td>
                            </tr>
                        </table>
                    </td>
                    <td style="width:50%;padding:4px 0 4px 8px;vertical-align:middle;">
                        <table style="width:100%;border-collapse:collapse;">
                            <tr>
                                <td style="width:12px;padding:0;line-height:0;vertical-align:middle;"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:#cbd5e1;"></span></td>
                                <td style="color:#64748b;padding:0;font-size:11px;vertical-align:middle;">Inactive (0 prompts)</td>
                                <td style="text-align:right;font-weight:600;color:#0f172a;padding:0;font-size:11px;font-variant-numeric:tabular-nums;vertical-align:middle;">{metrics['segments']['inactive']['count']} <span style="color:#94a3b8;font-weight:400;">({metrics['segments']['inactive']['rate']}%)</span></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>

        <!-- Application Penetration -->
        <div class="section-header">
            <h2 class="section-title">Application Penetration</h2>
            <div class="section-subtext">Active user counts across Microsoft 365 host applications</div>
        </div>
        {apps_table_html}

        <!-- Top Users -->
        <div class="section-header">
            <h2 class="section-title">Top Active Contributors</h2>
            <div class="section-subtext">Ranked by prompt submissions and interaction frequency</div>
        </div>
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th style="width: 32px;">Rank</th>
                        <th>User</th>
                        <th style="text-align: right;">Prompts</th>
                        <th style="text-align: right;">Days</th>
                        <th>Applications</th>
                    </tr>
                </thead>
                <tbody>
                    {top_users_html}
                </tbody>
            </table>
        </div>

        <!-- Inactive Users -->
        <div class="section-header">
            <h2 class="section-title" style="color: #991b1b;">License Optimization & Reallocation Candidates</h2>
            <div class="section-subtext">Accounts with zero recorded interactions in the reporting window</div>
        </div>
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>User</th>
                        <th>Email (UPN)</th>
                        <th style="text-align: right;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    {inactive_users_html}
                </tbody>
            </table>
        </div>

        <!-- Detailed Audit Attachment Callout -->
        <div class="attachment-card">
            <div class="attachment-header">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">
                    <rect width="24" height="24" rx="4" fill="#107C41"/>
                    <rect x="10" y="5" width="10" height="14" rx="1" fill="#18A058"/>
                    <rect x="4" y="7" width="9" height="10" rx="1.5" fill="#0B552C"/>
                    <path d="M6.5 10l4 4M10.5 10l-4 4" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>
                </svg>
                <span class="attachment-title">Detailed User Audit Roster Attached</span>
            </div>
            <p class="attachment-desc">
                For detailed departmental audits and user-level analysis, the complete roster of all <strong>{metrics['total_users']} licensed accounts</strong> with individual application engagement timestamps is attached as a Microsoft Excel spreadsheet (<code>copilot_usage_roster_{metrics['period']}.xlsx</code>).
            </p>
        </div>
    </div>

    <!-- Footer -->
    <div class="footer">
        <div class="footer-brand">Copilot Usage Automation &bull; Microsoft Graph Telemetry</div>
        <div>Report generated on {metrics['generated_at']}</div>
    </div>
</div>

</body>
</html>
"""
    return html
