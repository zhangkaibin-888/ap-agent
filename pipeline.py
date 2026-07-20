#!/usr/bin/env python3
"""
Invoice-to-Xero Pipeline
=========================
Combined pipeline: reads invoice PDFs from a shared mailbox via Microsoft Graph,
parses them (regex and/or LLM), creates bills in Xero, sends approval/summary emails.

Can be run:
  - Directly (manual / cron):  python3 pipeline.py
  - Via cron_trigger.py:       delegates to Logic App which POSTs to server.py

Requires:
  - config.json with Graph + Xero credentials
  - setup_xero_auth.py run once for Xero tokens
  - LLM_API_KEY env var (if using LLM/hybrid parsing mode)
"""

import json
import base64
import logging
import os
import re
import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from io import BytesIO

import requests
from msal import ConfidentialClientApplication
from pdfplumber import open as open_pdf

# ── Config ──
CONFIG_PATH = Path(__file__).parent / "config.json"
if not CONFIG_PATH.exists():
    print(f"FATAL: config.json not found at {CONFIG_PATH}")
    sys.exit(1)

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
log = logging.getLogger(APP_NAME.lower().replace(" ", "-"))

# ── Processed Cache ──
PROCESSED_CACHE = Path("/var/log/invoice-xero-processed-cache.json")

def _load_processed_cache():
    try:
        with open(PROCESSED_CACHE) as f:
            return set(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()

def _save_processed_cache(cache):
    try:
        PROCESSED_CACHE.parent.mkdir(parents=True, exist_ok=True)
        with open(PROCESSED_CACHE, "w") as f:
            json.dump(sorted(cache), f)
    except Exception as e:
        log.warning(f"Could not save processed cache: {e}")

def _smart_lookback_hours():
    """Monday=72h (catch Friday invoices), Tue-Sun=24h."""
    return 72 if datetime.now().weekday() == 0 else 24


# ═════════════════════════════════════════════════════════════════════════
# PART 1: Microsoft Graph — Read Inbox & Send Email
# ═════════════════════════════════════════════════════════════════════════

class GraphClient:
    """Connect to Microsoft Graph API to read/send mail for a shared mailbox."""

    def __init__(self, cfg):
        g = cfg["graph"]
        self.tenant = g["tenant_id"]
        self.client_id = g["client_id"]
        self.secret = g["client_secret"]
        self.mailbox = g["shared_mailbox"]
        self._token = None

    def _get_token(self):
        if self._token and self._token.get("expires_at", 0) > datetime.now().timestamp():
            return self._token["access_token"]
        app = ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant}",
            client_credential=self.secret,
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            raise RuntimeError(f"Graph auth failed: {result.get('error_description', result)}")
        self._token = {
            "access_token": result["access_token"],
            "expires_at": datetime.now().timestamp() + result.get("expires_in", 3600) - 60,
        }
        return self._token["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self._get_token()}", "Content-Type": "application/json"}

    def fetch_recent_emails(self, hours=None):
        """Get emails from the last N hours with PDF/zip attachments.
        Defaults to smart lookback (72h Monday, 24h otherwise)."""
        hours = hours if hours is not None else _smart_lookback_hours()
        since = (datetime.now() - timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ")
        url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/messages"
        params = {
            "$filter": f"receivedDateTime ge {since}",
            "$top": 50,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,hasAttachments,isRead",
        }
        resp = requests.get(url, headers=self._headers(), params=params)
        resp.raise_for_status()
        raw_messages = resp.json().get("value", [])
        messages = [m for m in raw_messages if m.get("hasAttachments")]
        log.info(f"Found {len(messages)} emails with attachments (last {hours}h)")
        return self._resolve_attachments(messages)

    def fetch_unread_invoices(self):
        """Fetch unread email from the shared mailbox that have PDF attachments."""
        url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/messages"
        params = {
            "$filter": "isRead eq false AND hasAttachments eq true",
            "$top": 20,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,hasAttachments",
        }
        resp = requests.get(url, headers=self._headers(), params=params)
        resp.raise_for_status()
        messages = resp.json().get("value", [])
        log.info(f"Found {len(messages)} unread emails with attachments")
        return self._resolve_attachments(messages)

    def _resolve_attachments(self, messages):
        results = []
        for msg in messages:
            log.info(f"  Checking: {msg['subject']}")
            att_url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/messages/{msg['id']}/attachments"
            att_resp = requests.get(att_url, headers=self._headers())
            att_resp.raise_for_status()

            pdfs = []
            zips = []
            for att in att_resp.json().get("value", []):
                name = att.get("name", "")
                if "@odata.mediaContentType" in att:
                    dl_url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/messages/{msg['id']}/attachments/{att['id']}/$value"
                    dl_resp = requests.get(dl_url, headers=self._headers())
                    dl_resp.raise_for_status()
                    content = dl_resp.content
                else:
                    content = base64.b64decode(att.get("contentBytes", ""))

                if name.lower().endswith(".pdf"):
                    pdfs.append({"name": name, "bytes": content})
                elif name.lower().endswith(".zip"):
                    zips.append(content)

            # Extract PDFs from ZIP attachments
            for zdata in zips:
                try:
                    with zipfile.ZipFile(BytesIO(zdata)) as zf:
                        for zname in zf.namelist():
                            if zname.lower().endswith(".pdf"):
                                pdfs.append({"name": zname, "bytes": zf.read(zname), "_from_zip": True})
                except Exception as e:
                    log.warning(f"    ⚠ Could not extract ZIP: {e}")

            if not pdfs:
                log.info("    No PDFs — skipping")
                continue

            from_data = msg.get("from", {}).get("emailAddress", {})
            results.append({
                "id": msg["id"],
                "subject": msg["subject"],
                "from_name": from_data.get("name", ""),
                "from_email": from_data.get("address", ""),
                "received": msg.get("receivedDateTime", ""),
                "pdfs": pdfs,
            })
            log.info(f"    → {len(pdfs)} PDF(s)")
        return results

    def mark_as_read(self, msg_id):
        """Mark a message as read so we don't process it again."""
        try:
            resp = requests.patch(
                f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/messages/{msg_id}",
                headers=self._headers(),
                json={"isRead": True},
            )
            if resp.ok:
                log.info("    ✓ Marked read")
            else:
                log.warning(f"    ⚠ Could not mark as read ({resp.status_code}) — skipping")
        except Exception as e:
            log.warning(f"    ⚠ Could not mark as read: {e}")

    def send_approval_email(self, to_email, supplier_name, invoice_num, amount, due_date, xero_url):
        """Send individual approval request email from the shared mailbox."""
        subject = f"Invoice {invoice_num} from {supplier_name} — Ready for Approval"
        body = f"""<html><body style="font-family:Arial,sans-serif;padding:20px;">
<h2>Invoice Ready for Approval</h2>
<table style="border-collapse:collapse;width:100%;max-width:600px;">
<tr style="background:#f0f0f0;"><td style="padding:8px;border:1px solid #ddd;"><strong>Supplier</strong></td><td style="padding:8px;border:1px solid #ddd;">{supplier_name}</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;"><strong>Invoice #</strong></td><td style="padding:8px;border:1px solid #ddd;">{invoice_num}</td></tr>
<tr style="background:#f0f0f0;"><td style="padding:8px;border:1px solid #ddd;"><strong>Amount</strong></td><td style="padding:8px;border:1px solid #ddd;">${amount:.2f}</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;"><strong>Due Date</strong></td><td style="padding:8px;border:1px solid #ddd;">{due_date}</td></tr>
</table>
<br/>
<p>This invoice has been created in Xero and is awaiting your approval.</p>
<p><a href="{xero_url}" style="background:#13B5EA;color:white;padding:12px 24px;text-decoration:none;border-radius:4px;display:inline-block;">View Invoice in Xero →</a></p>
<br/>
<p><em>This is an automated message from {APP_NAME}.</em></p>
</body></html>"""
        email_payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "html", "content": body},
                "toRecipients": [{"emailAddress": {"address": to_email}}],
            },
            "saveToSentItems": True,
        }
        url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/sendMail"
        resp = requests.post(url, headers=self._headers(), json=email_payload)
        if resp.ok:
            log.info(f"  ✓ Approval email sent to {to_email}")
        else:
            log.warning(f"  ⚠ Could not send approval email ({resp.status_code})")

    def send_summary(self, invoices_created, issues=None):
        """Send a summary report email to all notification recipients."""
        issues = issues or []
        has_invoices = len(invoices_created) > 0
        has_issues = len(issues) > 0

        invoice_rows = ""
        for inv in invoices_created:
            invoice_rows += f"""
<tr style="background:#f9f9f9;">
  <td style="padding:8px;border:1px solid #ddd;">{inv['supplier']}</td>
  <td style="padding:8px;border:1px solid #ddd;">{inv['number']}</td>
  <td style="padding:8px;border:1px solid #ddd;text-align:right;">${inv['amount']:.2f}</td>
  <td style="padding:8px;border:1px solid #ddd;">{inv['due_date']}</td>
  <td style="padding:8px;border:1px solid #ddd;"><a href="{inv['url']}" style="color:#13B5EA;">View</a></td>
</tr>"""
        total = sum(inv['amount'] for inv in invoices_created)

        issue_rows = ""
        for iss in issues:
            issue_rows += f"""
<tr style="background:#fff3f3;">
  <td style="padding:8px;border:1px solid #e0b0b0;">{iss.get('supplier','?')}</td>
  <td style="padding:8px;border:1px solid #e0b0b0;">{iss.get('subject','')[:60]}</td>
  <td style="padding:8px;border:1px solid #e0b0b0;">{iss.get('reason','')}</td>
</tr>"""

        parts = []
        if has_invoices:
            parts.append(f"""<h2>✅ Created Bills</h2>
<table style="border-collapse:collapse;width:100%;max-width:700px;">
<tr style="background:#13B5EA;color:white;">
  <th style="padding:8px;border:1px solid #13B5EA;text-align:left;">Supplier</th>
  <th style="padding:8px;border:1px solid #13B5EA;text-align:left;">Invoice</th>
  <th style="padding:8px;border:1px solid #13B5EA;text-align:right;">Amount</th>
  <th style="padding:8px;border:1px solid #13B5EA;text-align:left;">Due</th>
  <th style="padding:8px;border:1px solid #13B5EA;">Link</th>
</tr>{invoice_rows}
<tr style="font-weight:bold;">
  <td colspan="2" style="padding:8px;border:1px solid #ddd;text-align:right;">Total</td>
  <td style="padding:8px;border:1px solid #ddd;text-align:right;">${total:,.2f}</td>
  <td colspan="2" style="padding:8px;border:1px solid #ddd;"></td>
</tr>
</table>""")

        if has_issues:
            parts.append(f"""<h2 style="color:#cc0000;">⚠ Items Needing Attention</h2>
<table style="border-collapse:collapse;width:100%;max-width:700px;">
<tr style="background:#cc0000;color:white;">
  <th style="padding:8px;border:1px solid #cc0000;text-align:left;">Supplier</th>
  <th style="padding:8px;border:1px solid #cc0000;text-align:left;">Email</th>
  <th style="padding:8px;border:1px solid #cc0000;text-align:left;">Issue</th>
</tr>{issue_rows}
</table>""")

        if not has_invoices and not has_issues:
            parts.append("<p>Nothing to report.</p>")

        body_html = f"""<html><body style="font-family:Arial,sans-serif;padding:20px;">
<h1>📬 {APP_NAME} Daily Invoice Report</h1>
<p>Run at {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
{"".join(parts)}
<br/><p><em>Automated from {APP_NAME}</em></p>
</body></html>"""

        if has_invoices and has_issues:
            subject = f"{APP_NAME} — {len(invoices_created)} new, {len(issues)} issue(s)"
        elif has_issues:
            subject = f"{APP_NAME} — ⚠ {len(issues)} issue(s) need attention"
        else:
            subject = f"{APP_NAME} — {len(invoices_created)} invoice(s) processed"

        recipients = cfg.get("notification_recipients", [])
        if not recipients:
            log.info("No notification_recipients configured — skipping summary email")
            return

        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "html", "content": body_html},
                "toRecipients": [{"emailAddress": {"address": r}} for r in recipients],
            },
            "saveToSentItems": True,
        }
        try:
            resp = requests.post(
                f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/sendMail",
                headers=self._headers(),
                json=payload,
            )
            if resp.ok:
                log.info(f"  ✓ Summary email sent to {len(recipients)} recipient(s)")
            else:
                log.warning(f"  ⚠ Could not send summary email ({resp.status_code})")
        except Exception as e:
            log.warning(f"  ⚠ Could not send summary email: {e}")


# ═════════════════════════════════════════════════════════════════════════
# PART 2: Invoice PDF Parser
# ═════════════════════════════════════════════════════════════════════════

class InvoiceParser:
    """Extract structured data from invoice PDFs using regex patterns and/or LLM.

    Supports three parsing modes (configurable via parsing_mode):
      - "regex":  fast, uses only regex patterns, no API cost
      - "llm":    uses an LLM to extract fields from the PDF text
      - "hybrid": tries regex first, falls back to LLM if fields are missing
    """

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
        "supplier_name": [
            r"^(?:From|Supplier|Vendor)[^:]*:\s*(.+)",
        ],
    }

    def __init__(self, cfg):
        # Merge custom patterns with defaults
        custom = cfg.get("pdf_patterns", {})
        self._patterns = dict(self.DEFAULT_PATTERNS)
        for key, patterns in custom.items():
            if key in self._patterns:
                self._patterns[key] = patterns + self._patterns[key]
            else:
                self._patterns[key] = patterns

        self.mode = cfg.get("parsing_mode", "hybrid")

        # LLM configuration
        llm_cfg = cfg.get("llm", {})
        self.llm_provider = llm_cfg.get("provider", "openai")
        self.llm_model = llm_cfg.get("model", "gpt-4o-mini")
        self.llm_temperature = llm_cfg.get("temperature", 0.0)
        self.llm_max_tokens = llm_cfg.get("max_tokens", 500)
        self.llm_api_key = os.environ.get(llm_cfg.get("api_key_env", "LLM_API_KEY"), "")

        if self.llm_provider == "deepseek":
            self._llm_base_url = "https://api.deepseek.com/v1"
            self._llm_model = self.llm_model or "deepseek-chat"
        elif self.llm_provider == "openai":
            self._llm_base_url = "https://api.openai.com/v1"
            self._llm_model = self.llm_model or "gpt-4o-mini"
        elif self.llm_provider == "anthropic":
            self._llm_base_url = "https://api.anthropic.com/v1"
            self._llm_model = self.llm_model or "claude-3-haiku-20240307"
        elif self.llm_provider == "custom":
            self._llm_base_url = llm_cfg.get("base_url", "https://api.openai.com/v1").rstrip("/")
            self._llm_model = self.llm_model or "gpt-4o-mini"
        else:
            raise ValueError(f"Unknown LLM provider: {self.llm_provider}. Use: openai, deepseek, anthropic, custom")

        if not self.llm_api_key and self.mode in ("llm", "hybrid"):
            log.warning(f"LLM_API_KEY not set — LLM parsing will fail. Set env var {llm_cfg.get('api_key_env', 'LLM_API_KEY')}")

    # ── Regex Parsing ──────────────────────────────────────────────────

    def _extract_text(self, pdf_bytes):
        try:
            with open_pdf(BytesIO(pdf_bytes)) as pdf:
                texts = [p.extract_text() or "" for p in pdf.pages]
        except Exception as e:
            log.warning(f"  PDF parse error: {e}")
            return None
        text = "\n".join(texts)
        if not text.strip():
            log.warning("  PDF has no extractable text (scanned image?)")
            return None
        return text

    def _parse_regex(self, text, fallback_supplier=None):
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
                del result["total_amount"]
        return result

    # ── LLM Parsing ────────────────────────────────────────────────────

    def _parse_llm(self, text, fallback_supplier=None):
        """Call configured LLM to extract invoice fields from text."""
        prompt = f"""Extract invoice fields from this invoice text. Return ONLY a valid JSON object with these fields:
- "invoice_number": the invoice or reference number (string, or null if not found)
- "total_amount": the total amount due as a number (float, or null if not found)
- "supplier_name": the supplier/vendor name (string, or null if not found)
- "due_date": the due date in DD Month YYYY format (string, or null if not found)

CRITICAL RULES:
- total_amount must be the final total owed, including tax/GST — NOT a subtotal, NOT a unit price, NOT a reference number
- If a number appears without a currency sign AND also appears as an invoice/reference number nearby, it's likely NOT the total amount
- supplier_name is the company sending the invoice (the FROM party), NOT the recipient
- due_date is when payment is due, NOT the invoice date

Invoice text:
---
{text[:8000]}
---

Respond with ONLY the JSON object, no explanation."""

        if self.llm_provider == "anthropic":
            return self._call_anthropic(prompt, fallback_supplier)
        else:
            return self._call_openai_compat(prompt, fallback_supplier)

    def _call_openai_compat(self, prompt, fallback_supplier):
        """Call any OpenAI-compatible chat completions API."""
        try:
            resp = requests.post(
                f"{self._llm_base_url}/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.llm_api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": self._llm_model,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": self.llm_temperature,
                    "max_tokens": self.llm_max_tokens,
                },
                timeout=30,
            )
            resp.raise_for_status()
            raw = resp.json()["choices"][0]["message"]["content"].strip()
            return self._parse_llm_response(raw, fallback_supplier)
        except Exception as e:
            log.warning(f"  ⚠ LLM extraction failed: {e}")
            return self._empty_llm_result(fallback_supplier, f"LLM error: {e}")

    def _call_anthropic(self, prompt, fallback_supplier):
        """Call Anthropic's Messages API."""
        try:
            resp = requests.post(
                f"{self._llm_base_url}/messages",
                headers={
                    "x-api-key": self.llm_api_key,
                    "anthropic-version": "2023-06-01",
                    "Content-Type": "application/json",
                },
                json={
                    "model": self._llm_model,
                    "max_tokens": self.llm_max_tokens,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": self.llm_temperature,
                },
                timeout=30,
            )
            resp.raise_for_status()
            content_blocks = resp.json().get("content", [])
            raw = "".join(block.get("text", "") for block in content_blocks if block.get("type") == "text")
            return self._parse_llm_response(raw.strip(), fallback_supplier)
        except Exception as e:
            log.warning(f"  ⚠ Anthropic LLM extraction failed: {e}")
            return self._empty_llm_result(fallback_supplier, f"LLM error: {e}")

    def _parse_llm_response(self, raw, fallback_supplier):
        # Strip markdown code fences if present
        if raw.startswith("```"):
            raw = raw.split("\n", 1)[-1]
            raw = raw.rsplit("```", 1)[0].strip()

        data = json.loads(raw)
        result = {
            "invoice_number": data.get("invoice_number"),
            "total_amount": data.get("total_amount"),
            "supplier_name": data.get("supplier_name") or fallback_supplier,
            "due_date": data.get("due_date"),
            "qa_suspicious": False,
            "qa_reason": "",
        }
        # Ensure total_amount is float
        if result.get("total_amount") is not None:
            try:
                result["total_amount"] = float(result["total_amount"])
            except (ValueError, TypeError):
                result["total_amount"] = None
        return result

    def _empty_llm_result(self, fallback_supplier, reason=""):
        return {
            "invoice_number": None,
            "total_amount": None,
            "supplier_name": fallback_supplier,
            "due_date": None,
            "qa_suspicious": False,
            "qa_reason": reason,
        }

    # ── Public Parse Method ────────────────────────────────────────────

    def parse(self, pdf_bytes, fallback_supplier=None):
        """Parse a PDF invoice. Returns dict or None if PDF can't be read."""
        text = self._extract_text(pdf_bytes)
        if text is None:
            return None

        if self.mode == "regex":
            result = self._parse_regex(text, fallback_supplier)
        elif self.mode == "llm":
            result = self._parse_llm(text, fallback_supplier)
        else:  # hybrid — try regex first, fall back to LLM
            result = self._parse_regex(text, fallback_supplier)
            has_invoice_num = bool(result.get("invoice_number"))
            has_amount = result.get("total_amount") is not None
            if not has_invoice_num or not has_amount:
                log.info(f"  Regex partial ({'inv#' if has_invoice_num else 'no inv#'}, "
                         f"{'amt' if has_amount else 'no amt'}) — trying LLM...")
                llm_result = self._parse_llm(text, fallback_supplier)
                # Merge: LLM values override regex values
                if llm_result.get("invoice_number"):
                    result["invoice_number"] = llm_result["invoice_number"]
                if llm_result.get("total_amount") is not None:
                    result["total_amount"] = llm_result["total_amount"]
                if llm_result.get("supplier_name"):
                    result["supplier_name"] = llm_result["supplier_name"]
                if llm_result.get("due_date"):
                    result["due_date"] = llm_result["due_date"]
                result["qa_suspicious"] = llm_result.get("qa_suspicious", False)
                result["qa_reason"] = llm_result.get("qa_reason", "")

        log.info(f"  Parsed: #{result.get('invoice_number','?')} | "
                 f"${result.get('total_amount', 0):.2f} | "
                 f"due {result.get('due_date','?')} | "
                 f"supplier {result.get('supplier_name','?')}")
        return result


# ═════════════════════════════════════════════════════════════════════════
# PART 3: Xero API Integration
# ═════════════════════════════════════════════════════════════════════════

class XeroClient:
    """Connect to Xero API to manage contacts and invoices."""

    BASE_URL = "https://api.xero.com/api.xro/2.0"

    def __init__(self, cfg):
        x = cfg["xero"]
        self.client_id = x["client_id"]
        self.client_secret = x["client_secret"]
        self.tenant_id = x["tenant_id"]
        self.account_code = x.get("default_account_code", "200")
        self.tax_type = x.get("default_tax_type", "NONE")
        self._tokens = None
        tp = Path(__file__).parent / "xero_tokens.json"
        if tp.exists():
            with open(tp) as f:
                self._tokens = json.load(f)
        self._ensure_token()

    def _ensure_token(self):
        """Refresh OAuth 2.0 token if needed.

        Supports both authorization_code (refresh_token) and
        client_credentials grant flows automatically.
        """
        if not self._tokens:
            raise RuntimeError("Xero not authenticated. Run setup_xero_auth.py first.")

        if self._tokens.get("expires_at", 0) > datetime.now().timestamp():
            return

        # Try refresh_token flow first (authorization_code grant)
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

        # Fall back to client_credentials grant
        log.info("Refresh token expired or absent — using client_credentials grant")
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

    def find_contact(self, name):
        """Search Xero for a contact by name."""
        resp = requests.get(
            f"{self.BASE_URL}/Contacts",
            headers=self._headers(),
            params={"where": f'Name=="{name}"'},
        )
        resp.raise_for_status()
        contacts = resp.json().get("Contacts", [])
        if contacts:
            log.info(f"  Found existing Xero contact: {contacts[0]['Name']} (ID: {contacts[0]['ContactID']})")
            return contacts[0]
        log.info(f"  Contact '{name}' not found in Xero")
        return None

    def create_contact(self, name, email=""):
        """Create a new supplier contact in Xero."""
        resp = requests.post(
            f"{self.BASE_URL}/Contacts",
            headers=self._headers(),
            json={"Contacts": [{"Name": name, "IsSupplier": True, "EmailAddress": email}]},
        )
        resp.raise_for_status()
        contact = resp.json()["Contacts"][0]
        log.info(f"  ✓ Created Xero contact: {contact['Name']} (ID: {contact['ContactID']})")
        return contact

    def find_invoice(self, inv_num):
        """Check if an invoice already exists in Xero (duplicate check)."""
        resp = requests.get(
            f"{self.BASE_URL}/Invoices",
            headers=self._headers(),
            params={"where": f'InvoiceNumber=="{inv_num}"'},
        )
        if resp.ok:
            invoices = resp.json().get("Invoices", [])
            return invoices[0] if invoices else None
        return None

    def create_bill(self, contact_id, inv_num, amount, due_date, description, pdf_bytes=None, account_code=None, customer=None):
        """Create a bill (accounts payable invoice) in Xero.
        Status: SUBMITTED — awaiting approval, does NOT release payment.
        Optionally set a customer tracking category per line item."""
        line_item = {
            "Description": (description or "Invoice")[:4000],
            "Quantity": 1.0,
            "UnitAmount": amount,
            "AccountCode": account_code or self.account_code,
            "TaxType": self.tax_type,
        }
        if customer:
            line_item["Tracking"] = [{"Name": "Customer", "Option": customer[:100]}]
        invoice = {
            "Type": "ACCPAY",
            "Contact": {"ContactID": contact_id},
            "InvoiceNumber": inv_num,
            "Date": datetime.now().strftime("%Y-%m-%d"),
            "DueDate": due_date,
            "Status": "SUBMITTED",
            "LineItems": [line_item],
        }
        resp = requests.post(
            f"{self.BASE_URL}/Invoices",
            headers=self._headers(),
            json={"Invoices": [invoice]},
        )
        resp.raise_for_status()
        created = resp.json()["Invoices"][0]
        inv_id = created["InvoiceID"]
        log.info(f"  ✓ Xero invoice created: {inv_num} (ID: {inv_id}) — SUBMITTED (approval pending)")

        if pdf_bytes:
            self._attach_pdf(inv_id, pdf_bytes, f"{inv_num}.pdf")
        return created

    def _attach_pdf(self, inv_id, pdf_bytes, filename):
        """Attach the original invoice PDF to the Xero invoice."""
        headers = self._headers()
        headers["Content-Type"] = "application/pdf"
        resp = requests.post(
            f"{self.BASE_URL}/Invoices/{inv_id}/Attachments/{filename}",
            headers=headers,
            data=pdf_bytes,
        )
        if resp.status_code in (200, 201):
            log.info("    ✓ PDF attached to Xero invoice")
        elif resp.status_code == 409:
            log.info("    PDF already attached (duplicate)")
        else:
            log.warning(f"    PDF attach returned {resp.status_code}: {resp.text[:200]}")

    @staticmethod
    def invoice_url(inv_id):
        return f"https://go.xero.com/AccountsPayable/Edit/{inv_id}"

    def create_bill_with_line_items(self, contact_id, inv_num, due_date, line_items, reference="", pdf_bytes=None):
        """Create a bill with multiple line items (e.g. hardware invoices with per-item rows).
        Each line_item dict: {Description, Quantity, UnitAmount, AccountCode, TaxType, [Tracking]}."""
        invoice = {
            "Type": "ACCPAY",
            "Contact": {"ContactID": contact_id},
            "InvoiceNumber": inv_num,
            "Date": datetime.now().strftime("%Y-%m-%d"),
            "DueDate": due_date,
            "Status": "SUBMITTED",
            "Reference": reference[:255] if reference else "",
            "LineItems": line_items,
        }
        resp = requests.post(
            f"{self.BASE_URL}/Invoices",
            headers=self._headers(),
            json={"Invoices": [invoice]},
        )
        resp.raise_for_status()
        created = resp.json()["Invoices"][0]
        inv_id = created["InvoiceID"]
        log.info(f"  ✓ Xero bill created: {inv_num} (ID: {inv_id}) — SUBMITTED ({len(line_items)} line items)")

        if pdf_bytes:
            self._attach_pdf(inv_id, pdf_bytes, f"{inv_num}.pdf")
        return created


# ═════════════════════════════════════════════════════════════════════════
# PART 4: Pipeline Logic
# ═════════════════════════════════════════════════════════════════════════

def process_email(graph, xero, parser, email):
    """Process one email with its PDF attachments. Returns (created, issues)."""
    log.info(f"\n{'='*60}")
    log.info(f"Email: {email['subject']}")
    log.info(f"From: {email['from_name']} <{email['from_email']}>")
    created = []
    issues = []

    for pdf in email["pdfs"]:
        log.info(f"  PDF: {pdf['name']} ({len(pdf['bytes']):,} bytes)")

        inv_data = parser.parse(pdf["bytes"], fallback_supplier=email["from_name"])
        if not inv_data:
            log.warning("  ⚠ Could not parse PDF — forwarding to manual review")
            issues.append({
                "supplier": email["from_name"],
                "subject": email["subject"],
                "reason": "Could not parse PDF (possibly scanned image)",
            })
            continue

        supplier = inv_data.get("supplier_name") or email["from_name"]
        inv_num = inv_data.get("invoice_number")
        amount = inv_data.get("total_amount")
        due_date = inv_data.get("due_date") or datetime.now().strftime("%Y-%m-%d")

        if not inv_num or not amount:
            log.warning(f"  ⚠ Missing fields — inv#={'yes' if inv_num else 'no'}, "
                        f"amount={'yes' if amount else 'no'}")
            issues.append({
                "supplier": supplier,
                "subject": email["subject"],
                "reason": f"Missing fields (inv#={'yes' if inv_num else 'no'}, "
                          f"amount={'yes' if amount else 'no'})",
            })
            continue

        log.info(f"  Parsed: #{inv_num} | ${amount:.2f} | due {due_date} | supplier: {supplier}")

        # QA check (only from LLM parsing)
        if inv_data.get("qa_suspicious"):
            reason = inv_data.get("qa_reason", "Unknown reason")
            log.warning(f"  ⛔ QA FAILED: {reason}")
            log.warning(f"  ⛔ Skipping — manual review required for #{inv_num} (${amount:.2f})")
            issues.append({
                "supplier": supplier,
                "subject": email["subject"],
                "reason": f"QA flagged: {reason} (parsed ${amount:.2f})",
            })
            continue

        # Duplicate check (local cache first, then Xero)
        processed_cache = _load_processed_cache()
        if inv_num in processed_cache:
            log.info(f"  ⏭ Invoice #{inv_num} already in local cache — skipping (no Xero call)")
            continue
        existing = xero.find_invoice(inv_num)
        if existing:
            log.info(f"  ⏭ Invoice #{inv_num} already exists in Xero "
                     f"(Status: {existing.get('Status', '?')}) — skipping")
            processed_cache.add(inv_num)
            _save_processed_cache(processed_cache)
            continue

        # Find or create Xero contact
        contact = xero.find_contact(supplier)
        if not contact:
            contact = xero.create_contact(supplier, email.get("from_email", ""))

        # Create bill in Xero
        description = f"Invoice {inv_num} — {email['subject']}"
        xero_inv = xero.create_bill(contact["ContactID"], inv_num, amount, due_date,
                                    description, pdf["bytes"])
        url = XeroClient.invoice_url(xero_inv["InvoiceID"])
        log.info(f"  ✓ Xero bill created: {inv_num} (SUBMITTED) — {url}")
        created.append({
            "supplier": supplier,
            "number": inv_num,
            "amount": amount,
            "due_date": due_date,
            "url": url,
        })
        # Save to cache so future lookback runs skip duplicate Xero API calls
        processed_cache.add(inv_num)
        _save_processed_cache(processed_cache)

    # Mark email as read
    graph.mark_as_read(email["id"])
    return created, issues


# ═════════════════════════════════════════════════════════════════════════
# PART 5: Main Entry Point
# ═════════════════════════════════════════════════════════════════════════

def main():
    log.info("=" * 60)
    log.info(f"{APP_NAME} Invoice-to-Xero Pipeline — Starting ({datetime.now().isoformat()})")
    log.info("=" * 60)

    graph = GraphClient(cfg)
    xero = XeroClient(cfg)
    parser = InvoiceParser(cfg)

    # Fetch emails — prefer unread, fall back to smart lookback
    emails = graph.fetch_unread_invoices()
    if not emails:
        log.info("No unread emails, checking recent (smart lookback)...")
        emails = graph.fetch_recent_emails()

    log.info(f"\nProcessing {len(emails)} invoice emails...")

    all_created = []
    all_issues = []
    for email in emails:
        try:
            created, issues = process_email(graph, xero, parser, email)
            all_created.extend(created)
            all_issues.extend(issues)

            # Send individual approval email if approver configured
            approver = cfg.get("approver_email", "")
            if approver and created:
                for inv in created:
                    graph.send_approval_email(
                        approver,
                        inv["supplier"],
                        inv["number"],
                        inv["amount"],
                        inv["due_date"],
                        inv["url"],
                    )
        except Exception as e:
            log.exception(f"Failed to process {email.get('subject','?')}: {e}")
            all_issues.append({
                "supplier": "?",
                "subject": email.get("subject", "?"),
                "reason": f"Error: {e}",
            })

    log.info(f"\n{'='*60}")
    log.info(f"Complete. Processed {len(emails)} emails, "
             f"created {len(all_created)} new bills.")
    if all_issues:
        log.info(f"  ⚠ {len(all_issues)} issue(s) required attention")
    log.info(f"{'='*60}")

    # Send summary notification
    if all_created or all_issues:
        graph.send_summary(all_created, all_issues)
    else:
        log.info("No new invoices or issues — skipping notification email.")


if __name__ == "__main__":
    main()
