#!/usr/bin/env python3
"""
Microsoft 365 Copilot Usage Automation CLI & Runbook Entrypoint.

Usage:
  # 1. Local Mock Run (No credentials required - generates HTML report directly):
  python main.py --mock

  # 2. Local Live Run with Service Principal (.env configured):
  python main.py --period D30

  # 3. Azure Automation or Scheduled Run with Email Dispatch:
  python main.py --send-mail
"""

import sys
import argparse
import logging
from datetime import datetime
from pathlib import Path

import config
from auth import get_graph_access_token, GraphAuthError
from graph_client import (
    fetch_copilot_usage_user_detail,
    fetch_copilot_user_count_trend,
    GraphReportError,
)
from report_generator import aggregate_metrics, generate_html_report
from excel_generator import generate_excel_roster
from mock_data import generate_mock_copilot_users, generate_mock_copilot_trend
from mailer import send_report_email, MailerError

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("copilot_automation")


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Microsoft 365 Copilot Usage Report Generator & Email Automation"
    )
    parser.add_argument(
        "--mock",
        action="store_true",
        help="Generate report using realistic mock telemetry without calling live Graph API.",
    )
    parser.add_argument(
        "--period",
        type=str,
        default=config.REPORT_PERIOD,
        choices=["D7", "D28", "D30", "D90", "D180"],
        help="Reporting window: D7, D28, D30, D90, D180 (Default: D28)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Destination path for HTML report (Default: output/copilot_usage_report_<period>_<timestamp>.html)",
    )
    parser.add_argument(
        "--send-mail",
        action="store_true",
        help="Dispatch generated report to configured recipients via Microsoft Graph sendMail.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable detailed DEBUG logging.",
    )
    return parser.parse_args()


def run():
    args = parse_arguments()
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    logger.info("=== Microsoft 365 Copilot Usage Automation Initialized ===")
    logger.info(f"Reporting Window: {args.period}")

    token = None
    users = []
    trend = []

    # 1. Fetch Data (Mock or Live Graph API)
    if args.mock:
        logger.info("MOCK MODE ACTIVE: Generating 45 realistic sample users and trend points...")
        users = generate_mock_copilot_users(count=45)
        trend = generate_mock_copilot_trend(days=28)
    else:
        try:
            logger.info("Authenticating to Microsoft Graph API...")
            token = get_graph_access_token()
            logger.info("Fetching Copilot user activity details from Graph API...")
            users = fetch_copilot_usage_user_detail(token, period=args.period)
            logger.info("Fetching daily active user and prompt trend...")
            trend = fetch_copilot_user_count_trend(token, period=args.period)
        except (GraphAuthError, GraphReportError) as e:
            logger.error(f"ERROR: {e}")
            sys.exit(1)
        except Exception as e:
            logger.exception(f"Unexpected error encountered: {e}")
            sys.exit(1)

    if not users:
        logger.warning("No user activity records found!")
        sys.exit(0)

    # 2. Aggregate Metrics & Generate HTML
    logger.info(f"Aggregating metrics for {len(users)} users and {len(trend)} trend points...")
    metrics = aggregate_metrics(users, period=args.period, trend=trend)
    logger.info(
        f"Summary: {metrics['total_users']} Assigned | {metrics['active_users']} Active "
        f"({metrics['adoption_rate']}%) | {metrics['inactive_users']} Inactive"
    )

    logger.info("Generating executive email HTML report...")
    html_content = generate_html_report(metrics)

    # 3. Save to Local Output HTML & Excel Files
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if args.output:
        output_file = Path(args.output)
    else:
        output_file = config.OUTPUT_DIR / f"copilot_usage_report_{args.period}_{timestamp}.html"

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Save canonical 'latest.html' for quick opening
    latest_file = config.OUTPUT_DIR / "copilot_usage_report_latest.html"
    with open(latest_file, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Generate executive Excel audit roster
    excel_file = config.OUTPUT_DIR / f"copilot_usage_roster_{args.period}_{timestamp}.xlsx"
    generate_excel_roster(users, excel_file)
    latest_excel = config.OUTPUT_DIR / "copilot_usage_roster_latest.xlsx"
    generate_excel_roster(users, latest_excel)

    logger.info("HTML Report and Excel Roster successfully generated:")
    logger.info(f"   HTML:  {output_file.resolve()}")
    logger.info(f"   Excel: {excel_file.resolve()}")
    logger.info(f"   Quick HTML Preview:  {latest_file.resolve()}")
    logger.info(f"   Quick Excel Preview: {latest_excel.resolve()}")

    # 4. Optional: Send Email
    if args.send_mail:
        logger.info("Email dispatch requested...")
        if args.mock and not token:
            try:
                # Try acquiring token for sending test mail if configured
                token = get_graph_access_token()
            except Exception as e:
                logger.warning(f"Could not acquire token to send mail in mock mode ({e}). Skipping email dispatch.")
                return

        if token:
            try:
                send_report_email(token, html_content, attachment_paths=[latest_excel])
                logger.info("Report email dispatched successfully.")
            except MailerError as e:
                logger.error(f"Email dispatch failed: {e}")
                sys.exit(1)
    else:
        logger.info("Local mode: Email not dispatched (HTML and Excel generated only). Add --send-mail to dispatch.")

    logger.info("=== Automation Completed Successfully ===")


if __name__ == "__main__":
    run()
