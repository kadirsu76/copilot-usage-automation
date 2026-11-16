"""
Excel & CSV Exporter Module for Copilot Usage Telemetry.
Generates an executive-grade .xlsx spreadsheet (with .csv fallback) of the complete user roster.
Formatted with professional styling, column widths, freeze panes, and alignment.
"""

import csv
import logging
from pathlib import Path
from typing import List, Dict, Any

logger = logging.getLogger(__name__)


def generate_excel_roster(users: List[Dict[str, Any]], output_path: Path) -> Path:
    """
    Generates a professionally formatted Excel spreadsheet of all users.
    Falls back to CSV if openpyxl is not available.
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Copilot User Roster"
        ws.views.sheetView[0].showGridLines = True

        # Styles
        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        regular_font = Font(name="Segoe UI", size=10, color="0F172A")
        bold_font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
        muted_font = Font(name="Segoe UI", size=9, color="64748B")
        active_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        active_font = Font(name="Segoe UI", size=9, bold=True, color="166534")
        inactive_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        inactive_font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0"),
        )

        headers = [
            "Rank",
            "Display Name",
            "User Principal Name (Email)",
            "Status",
            "Total Prompts",
            "Active Days",
            "Engagement Segment",
            "Last Activity Date",
            "Teams",
            "Business Chat",
            "Outlook",
            "Word",
            "Excel",
            "PowerPoint",
            "Loop",
            "OneNote",
        ]

        ws.append(headers)
        ws.row_dimensions[1].height = 26

        for col_idx, cell in enumerate(ws[1], 1):
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center" if col_idx in (1, 4, 5, 6) else "left", vertical="center")

        # Sort users by activity
        sorted_users = sorted(
            users,
            key=lambda x: (x["is_active"], x.get("prompts_submitted", 0), x["total_active_days"]),
            reverse=True,
        )

        for idx, u in enumerate(sorted_users, 1):
            prompts = u.get("prompts_submitted", 0)
            days = u["total_active_days"]

            if not u["is_active"]:
                segment = "Inactive"
            elif prompts >= 50 or days >= 10:
                segment = "Power User"
            elif prompts >= 10 or days >= 3:
                segment = "Regular User"
            else:
                segment = "Occasional"

            app_act = u.get("app_last_activity", {})

            row_data = [
                idx,
                u["display_name"],
                u["user_principal_name"],
                "Active" if u["is_active"] else "Inactive",
                prompts,
                days,
                segment,
                u["last_activity_date"] or "-",
                app_act.get("Teams") or "-",
                app_act.get("Chat") or "-",
                app_act.get("Outlook") or "-",
                app_act.get("Word") or "-",
                app_act.get("Excel") or "-",
                app_act.get("PowerPoint") or "-",
                app_act.get("Loop") or "-",
                app_act.get("OneNote") or "-",
            ]
            ws.append(row_data)
            row_idx = idx + 1
            ws.row_dimensions[row_idx].height = 20

            # Cell formatting
            is_even = (idx % 2 == 0)
            for c_idx, cell in enumerate(ws[row_idx], 1):
                cell.font = regular_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

                if is_even:
                    cell.fill = zebra_fill

                # Alignments
                if c_idx == 1:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = muted_font
                elif c_idx == 4:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if u["is_active"]:
                        cell.fill = active_fill
                        cell.font = active_font
                    else:
                        cell.fill = inactive_fill
                        cell.font = inactive_font
                elif c_idx in (5, 6):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0"
                    if c_idx == 5 and prompts > 0:
                        cell.font = bold_font
                elif c_idx >= 8:
                    cell.font = muted_font

        # Auto column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

        # Freeze header row
        ws.freeze_panes = "A2"

        # Save xlsx
        xlsx_path = output_path.with_suffix(".xlsx")
        wb.save(xlsx_path)
        logger.info(f"Excel roster generated: {xlsx_path}")
        return xlsx_path

    except ImportError:
        logger.warning("openpyxl not found; falling back to CSV export.")
        return export_csv_roster(users, output_path.with_suffix(".csv"))


def export_csv_roster(users: List[Dict[str, Any]], output_path: Path) -> Path:
    """Exports raw roster to CSV using built-in csv module."""
    sorted_users = sorted(
        users,
        key=lambda x: (x["is_active"], x.get("prompts_submitted", 0), x["total_active_days"]),
        reverse=True,
    )
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Rank", "Display Name", "Email (UPN)", "Status", "Total Prompts",
            "Active Days", "Last Activity Date", "Applications"
        ])
        for idx, u in enumerate(sorted_users, 1):
            writer.writerow([
                idx,
                u["display_name"],
                u["user_principal_name"],
                "Active" if u["is_active"] else "Inactive",
                u.get("prompts_submitted", 0),
                u["total_active_days"],
                u["last_activity_date"],
                ", ".join(u.get("apps_used", []))
            ])
    logger.info(f"CSV roster generated: {output_path}")
    return output_path
