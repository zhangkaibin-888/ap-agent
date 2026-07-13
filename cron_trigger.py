#!/usr/bin/env python3
"""
Daily Cron Trigger
===================
Called by cron (or any scheduler) to trigger the Logic App
to check the shared mailbox for invoices.

Usage:
  python3 cron_trigger.py

This triggers the Logic App which then POSTs invoice data
to the server.py HTTP endpoint for processing.
"""

import json
import requests
import logging
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.json"
with open(CONFIG_PATH) as f:
    cfg = json.load(f)

APP_NAME = cfg.get("app_name", "InvoiceXero")

logging.basicConfig(
    level=getattr(logging, cfg.get("log_level", "INFO")),
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(cfg.get("log_file", "/var/log/invoice-xero-bridge.log")),
    ],
)
log = logging.getLogger(APP_NAME.lower().replace(" ", "-cron"))

LOGIC_APP_URL = cfg.get("logic_app_url", "")


def main():
    if not LOGIC_APP_URL:
        log.error("No logic_app_url in config.json")
        return

    log.info("Triggering Logic App to check for invoices...")
    try:
        resp = requests.post(LOGIC_APP_URL, json={"action": "check_inbox"}, timeout=120)
        log.info(f"Logic App responded: {resp.status_code}")
        log.info(f"Response: {resp.text[:500]}")
    except Exception as e:
        log.error(f"Failed to trigger Logic App: {e}")


if __name__ == "__main__":
    main()
