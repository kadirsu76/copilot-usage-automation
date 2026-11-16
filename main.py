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
        help="Gerçek Graph API çağrısı yapmadan örnek mock verilerle test HTML raporu üretir.",
    )
    parser.add_argument(
        "--period",
        type=str,
        default=config.REPORT_PERIOD,
        choices=["D7", "D28", "D30", "D90", "D180"],
        help="Raporlama dönemi: D7, D28, D30, D90, D180 (Varsayılan: config.py veya D28)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Oluşturulacak HTML raporunun dosya yolu (Varsayılan: output/copilot_usage_report_YYYYMMDD.html)",
    )
    parser.add_argument(
        "--send-mail",
        action="store_true",
        help="Üretilen raporu tanımlı alıcılara e-posta olarak gönderir (Graph API Mail.Send gerektirir).",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Ayrıntılı hata ayıklama (DEBUG) loglarını açar.",
    )
    return parser.parse_args()


def run():
    args = parse_arguments()
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    logger.info("=== Microsoft 365 Copilot Kullanım Otomasyonu Başlatıldı ===")
    logger.info(f"Rapor Dönemi: {args.period}")

    token = None
    users = []
    trend = []

    # 1. Fetch Data (Mock or Live Graph API)
    if args.mock:
        logger.info("🧪 MOCK MODU AKTİF: 45 adet gerçekçi örnek kullanıcı verisi ve trend üretiliyor...")
        users = generate_mock_copilot_users(count=45)
        trend = generate_mock_copilot_trend(days=28)
    else:
        try:
            logger.info("Graph API kimlik doğrulaması yapılıyor...")
            token = get_graph_access_token()
            logger.info("Kullanım verileri Graph API'den çekiliyor...")
            users = fetch_copilot_usage_user_detail(token, period=args.period)
            logger.info("Günlük kullanım ve prompt trendi çekiliyor...")
            trend = fetch_copilot_user_count_trend(token, period=args.period)
        except (GraphAuthError, GraphReportError) as e:
            logger.error(f"HATA: {e}")
            sys.exit(1)
        except Exception as e:
            logger.exception(f"Beklenmeyen bir hata oluştu: {e}")
            sys.exit(1)

    if not users:
        logger.warning("Herhangi bir kullanıcı kullanım kaydı bulunamadı!")
        sys.exit(0)

    # 2. Aggregate Metrics & Generate HTML
    logger.info(f"Toplam {len(users)} kullanıcı ve {len(trend)} günlük trend için metrikler hesaplanıyor...")
    metrics = aggregate_metrics(users, period=args.period, trend=trend)
    logger.info(
        f"Özet: {metrics['total_users']} Lisanslı | {metrics['active_users']} Aktif "
        f"({metrics['adoption_rate']}%) | {metrics['inactive_users']} İnaktif"
    )

    logger.info("HTML raporu oluşturuluyor...")
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

    logger.info(f"✅ HTML Raporu ve Excel Listesi başarıyla kaydedildi:")
    logger.info(f"   📄 HTML: {output_file.resolve()}")
    logger.info(f"   📊 Excel: {excel_file.resolve()}")
    logger.info(f"   🌐 Hızlı HTML İnceleme: {latest_file.resolve()}")
    logger.info(f"   📑 Hızlı Excel İnceleme: {latest_excel.resolve()}")

    # 4. Optional: Send Email
    if args.send_mail:
        logger.info("E-posta gönderimi tetiklendi...")
        if args.mock and not token:
            try:
                # Try acquiring token for sending test mail if configured
                token = get_graph_access_token()
            except Exception as e:
                logger.warning(f"Mock modunda mail göndermek için geçerli token alınamadı ({e}). Mail gönderimi atlanıyor.")
                return

        if token:
            try:
                send_report_email(token, html_content, attachment_paths=[latest_excel])
                logger.info("✅ Mail gönderimi tamamlandı.")
            except MailerError as e:
                logger.error(f"Mail gönderimi başarısız: {e}")
                sys.exit(1)
    else:
        logger.info("ℹ️ Local mod: Mail gönderilmedi (sadece HTML ve Excel üretildi). Mail göndermek için --send-mail ekleyin.")

    logger.info("=== Otomasyon Tamamlandı ===")


if __name__ == "__main__":
    run()
