#!/usr/bin/env python3
"""
Invoice-to-Xero HTTP Server
============================
Runs as a PM2 service. Listens for invoice data forwarded by a Logic App.
Handles Xero contact/invoice creation and sends the approval email.

Expected POST body from Logic App:
{
  "invoice_pdf_base64": "<base64 of PDF>",
  "invoice_pdf_filename": "invoice_123.pdf",
  "email_subject": "Invoice from Supplier XYZ",
  "from_name": "Supplier XYZ Pty Ltd",
  "from_email": "billing@supplierxyz.com",
  "approver_email": "approver@example.com"
}
"""

import json
import base64
import logging
import re
from datetime import datetime
from pathlib import Path
from io import BytesIO
from http.server import HTTPServer, BaseHTTPRequestHandler

import requests
from pdfplumber import open as open_pdf

# Load config
CONFIG_PATH = Path(__file__).parent / "config.json"
if not CONFIG_PATH.exists():
    raise RuntimeError(f"config.json not found at {CONFIG_PATH}")

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
log = logging.getLogger(APP_NAME.lower().replace(" ", "-server"))

PORT = cfg.get("port", 8088)

# ─────────────────────────────────────────────────────────────────────
# Xero Client
# ─────────────────────────────────────────────────────────────────────

class XeroClient:
    BASE_URL = "https://api.xero.com/api.xro/2.0"

    def __init__(self, config):
        self.client_id = config["xero"]["client_id"]
        self.client_secret = config["xero"]["client_secret"]
        self.tenant_id = config["xero"]["tenant_id"]
        self._tokens = None
        token_path = Path(__file__).parent / "xero_tokens.json"
        if token_path.exists():
            with open(token_path) as f:
                self._tokens = json.load(f)
        self._ensure_token()

    def _ensure_token(self):
        if not self._tokens:
            raise RuntimeError("Xero not authenticated. Run setup_xero_auth.py first.")
        if self._tokens.get("expires_at", 0) > datetime.now().timestamp():
            return

        # Try refresh_token flow first
        if "refresh_token" in self._tokens:
            resp = requests.post(
                "https://identity.xero.com/connect/token",
                data={
                    "grant_type": "refresh_token",
                    "refresh_token": self._tokens["refresh_token"],
                    "client_id": self.client_id,
                    "client_secret": self.client_secret,
                },
            )
            if resp.ok:
                data = resp.json()
                data["expires_at"] = datetime.now().timestamp() + data.get("expires_in", 3600) - 60
                self._tokens = data
                self._save_tokens()
                return

        # Fall back to client_credentials
        raw_auth = base64.b64encode(f"{self.client_id}:{self.client_secret}".encode()).decode()
        resp = requests.post(
            "https://identity.xero.com/connect/token",
            headers={
                "Authorization": f"Basic {raw_auth}",
                "Content-Type": "application/x-www-form-urlencoded",
            },
            data="grant_type=client_credentials&scope=accounting.invoices accounting.contacts accounting.attachments",
        )
        resp.raise_for_status()
        data = resp.json()
        data["expires_at"] = datetime.now().timestamp() + data.get("expires_in", 1800) - 60
        self._tokens = data
        self._save_tokens()

    def _save_tokens(self):
        with open(Path(__file__).parent / "xero_tokens.json", "w") as f:
            json.dump(self._tokens, f, indent=2)

    def _headers(self):
        return {
            "Authorization": f"Bearer {self._tokens['access_token']}",
            "Xero-tenant-id": self.tenant_id,
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def find_or_create_contact(self, name, email=""):
        url = f"{self.BASE_URL}/Contacts"
        params = {"where": f'Name=="{name}"'}
        resp = requests.get(url, headers=self._headers(), params=params)
        resp.raise_for_status()
        contacts = resp.json().get("Contacts", [])
        if contacts:
            log.info(f"  Existing Xero contact: {contacts[0]['Name']} (ID: {contacts[0]['ContactID']})")
            return contacts[0]
        payload = {"Contacts": [{"Name": name, "IsSupplier": True, "EmailAddress": email}]}
        resp = requests.post(url, headers=self._headers(), json=payload)
        resp.raise_for_status()
        contact = resp.json()["Contacts"][0]
        log.info(f"  ✓ Created Xero contact: {contact['Name']} (ID: {contact['ContactID']})")
        return contact

    def create_bill(self, contact_id, invoice_num, amount, due_date, description, pdf_bytes=None):
        url = f"{self.BASE_URL}/Invoices"
        invoice = {
            "Type": "ACCPAY",
            "Contact": {"ContactID": contact_id},
            "InvoiceNumber": invoice_num,
            "Date": datetime.now().strftime("%Y-%m-%d"),
            "DueDate": due_date,
            "Status": "SUBMITTED",
            "LineItems": [{
                "Description": description[:4000] if description else "Invoice",
                "Quantity": 1.0,
                "UnitAmount": amount,
                "AccountCode": cfg.get("xero", {}).get("default_account_code", "200"),
                "TaxType": cfg.get("xero", {}).get("default_tax_type", "NONE"),
            }],
        }
        resp = requests.post(url, headers=self._headers(), json={"Invoices": [invoice]})
        resp.raise_for_status()
        created = resp.json()["Invoices"][0]
        invoice_id = created["InvoiceID"]
        log.info(f"  ✓ Xero invoice created: {invoice_num} (ID: {invoice_id}) — SUBMITTED")
        if pdf_bytes:
            self._attach_pdf(invoice_id, pdf_bytes, f"{invoice_num}.pdf")
        return created

    def _attach_pdf(self, invoice_id, pdf_bytes, filename):
        url = f"{self.BASE_URL}/Invoices/{invoice_id}/Attachments/{filename}"
        headers = self._headers()
        headers["Content-Type"] = "application/pdf"
        resp = requests.post(url, headers=headers, data=pdf_bytes)
        if resp.status_code in (200, 201):
            log.info(f"  ✓ PDF attached")
        elif resp.status_code == 409:
            log.info(f"  PDF already attached")
        else:
            log.warning(f"  PDF attach status {resp.status_code}: {resp.text[:200]}")

    def get_invoice_url(self, invoice_id):
        return f"https://go.xero.com/AccountsPayable/Edit/{invoice_id}"


# ─────────────────────────────────────────────────────────────────────
# Invoice PDF Parser
# ─────────────────────────────────────────────────────────────────────

class InvoiceParser:
    DEFAULT_PATTERNS = {
        "invoice_number": [
            r"(?:Invoice|Inv\.?|Invoice\s+#)\s*[:\s]*([A-Z0-9][-A-Z0-9/]+)",
            r"(?:Invoice|Inv\.?)\s*No\s*[.:]?\s*([A-Z0-9][-A-Z0-9/]+)",
            r"Tax\s+Invoice\s+#?\s*([A-Z0-9][-A-Z0-9/]+)",
        ],
        "total_amount": [
            r"(?:Total|Amount\s+Due|Balance\s+Due|Grand\s+Total)[^$]*?\$?([\d,]+,?\d{0,2})",
            r"(?:Total|Amount)[^:]*:\s*\$?([\d,]+,?\d{2})",
            r"TOTAL\s+(?:USD|AUD|NZD)?\s*\$?([\d,]+,?\d{2})",
        ],
        "due_date": [
            r"(?:Due\s+Date|Payment\s+Due|Due)[^:]*:\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
            r"(?:Due[^:]*)[^:]*:\s*(\d{1,2}\s+[A-Z][a-z]+\s+\d{4})",
        ],
    }

    def __init__(self, config):
        custom = config.get("pdf_patterns", {})
        self._patterns = dict(self.DEFAULT_PATTERNS)
        for key, patterns in custom.items():
            if key in self._patterns:
                self._patterns[key] = patterns + self._patterns[key]
            else:
                self._patterns[key] = patterns

    def parse(self, pdf_bytes, fallback_supplier=None):
        try:
            with open_pdf(BytesIO(pdf_bytes)) as pdf:
                text = "\n".join([p.extract_text() or "" for p in pdf.pages])
        except Exception as e:
            log.warning(f"PDF parse failed: {e}")
            return None

        if not text.strip():
            log.warning("No text in PDF (scanned image?)")
            return {
                "invoice_number": None,
                "total_amount": None,
                "due_date": None,
                "supplier_name": fallback_supplier,
                "raw_text": "",
            }

        result = {}
        for field, patterns in self._patterns.items():
            for pattern in patterns:
                m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                if m:
                    result[field] = m.group(1).strip()
                    break

        if "supplier_name" not in result and fallback_supplier:
            result["supplier_name"] = fallback_supplier

        if "total_amount" in result:
            try:
                result["total_amount"] = float(result["total_amount"].replace(",", ""))
            except ValueError:
                pass

        log.info(f"  Parsed: #{result.get('invoice_number','?')} | "
                 f"${result.get('total_amount',0):.2f} | "
                 f"due {result.get('due_date','?')} | "
                 f"supplier {result.get('supplier_name','?')}")
        return result


# ─────────────────────────────────────────────────────────────────────
# Approval Email via Logic App
# ─────────────────────────────────────────────────────────────────────

def send_approval_via_logic_app(supplier_name, invoice_num, amount, due_date,
                                 xero_url, approver_email):
    """Ask the Logic App to send the approval email (it has the mailbox connection)."""
    logic_app_url = cfg.get("logic_app_url", "")
    if not logic_app_url:
        log.warning("No logic_app_url configured — approval email not sent")
        return

    payload = {
        "action": "send_approval_email",
        "to": approver_email,
        "supplier_name": supplier_name,
        "invoice_number": invoice_num,
        "amount": amount,
        "due_date": due_date,
        "xero_invoice_url": xero_url,
    }
    try:
        resp = requests.post(logic_app_url, json=payload, timeout=30)
        log.info(f"  Approval email sent via Logic App: {resp.status_code}")
    except Exception as e:
        log.error(f"  Failed to send approval email: {e}")


# ─────────────────────────────────────────────────────────────────────
# Main Pipeline
# ─────────────────────────────────────────────────────────────────────

def process_invoice(data):
    """
    Process one invoice from the Logic App payload.
    Returns dict with {success, invoice_num, xero_url, error_message}
    """
    pdf_b64 = data.get("invoice_pdf_base64", "")
    if not pdf_b64:
        return {"success": False, "error": "Missing invoice_pdf_base64"}

    try:
        pdf_bytes = base64.b64decode(pdf_b64)
    except Exception as e:
        return {"success": False, "error": f"Invalid base64: {e}"}

    from_name = data.get("from_name", "Unknown Supplier")
    from_email = data.get("from_email", "")

    # Initialize clients
    xero = XeroClient(cfg)
    parser = InvoiceParser(cfg)
    approver = data.get("approver_email") or cfg.get("approver_email", "")

    # Parse PDF
    invoice_data = parser.parse(pdf_bytes, fallback_supplier=from_name)
    if not invoice_data:
        return {"success": False, "error": "Could not parse PDF"}

    supplier_name = invoice_data.get("supplier_name", from_name)
    invoice_num = invoice_data.get("invoice_number")
    amount = invoice_data.get("total_amount")
    due_date = invoice_data.get("due_date") or datetime.now().strftime("%Y-%m-%d")

    if not invoice_num or not amount:
        missing = []
        if not invoice_num: missing.append("invoice_number")
        if not amount: missing.append("total_amount")
        return {"success": False, "error": f"Could not extract: {', '.join(missing)}"}

    # Find or create Xero contact
    contact = xero.find_or_create_contact(supplier_name, from_email)

    # Create Xero bill
    description = f"Invoice {invoice_num} — {data.get('email_subject', '')}"
    xero_invoice = xero.create_bill(contact["ContactID"], invoice_num, amount,
                                     due_date, description, pdf_bytes)

    # Send approval email via Logic App
    xero_url = xero.get_invoice_url(xero_invoice["InvoiceID"])
    if approver:
        send_approval_via_logic_app(supplier_name, invoice_num, amount, due_date,
                                     xero_url, approver)

    return {
        "success": True,
        "invoice_num": invoice_num,
        "supplier": supplier_name,
        "amount": amount,
        "xero_url": xero_url,
    }


# ─────────────────────────────────────────────────────────────────────
# HTTP Server
# ─────────────────────────────────────────────────────────────────────

class InvoiceHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length)

        try:
            data = json.loads(body)
        except json.JSONDecodeError as e:
            self._respond(400, {"success": False, "error": f"Invalid JSON: {e}"})
            return

        log.info(f"Received invoice: {data.get('email_subject', '(no subject)')}")
        try:
            result = process_invoice(data)
            status = 200 if result.get("success") else 422
        except Exception as e:
            log.exception(f"Pipeline error: {e}")
            result = {"success": False, "error": str(e)}
            status = 500
        self._respond(status, result)

    def do_GET(self):
        if self.path == "/health":
            self._respond(200, {"status": "ok", "time": datetime.now().isoformat()})
        else:
            self._respond(404, {"error": "Not found. POST to / with invoice data."})

    def _respond(self, status, body):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(body).encode())

    def log_message(self, format, *args):
        log.info(f"HTTP {self.command} {self.path} — {args[0]} {args[1]}")


def main():
    log.info(f"Starting {APP_NAME} Invoice Server on port {PORT}")
    server = HTTPServer(("0.0.0.0", PORT), InvoiceHandler)
    log.info(f"Listening on http://0.0.0.0:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        log.info("Shutting down")
        server.shutdown()


if __name__ == "__main__":
    main()
