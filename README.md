# Invoice-to-Xero Bridge

Automatically processes invoice PDFs emailed to a shared mailbox (via Microsoft Graph), extracts invoice data using regex and/or LLM, creates bills in Xero, and sends approval/summary notification emails.

The application name is fully configurable via `config.json` — all email subjects, log messages, report titles, and user-facing strings use the configured `app_name`.

---

## Architecture

```
                              ┌──────────────────────┐
  9 AM cron ─────────────────▶│  cron_trigger.py     │
  (or any scheduler)          │  HTTP POST to        │
                              │  Logic App           │
                              └──────────┬───────────┘
                                         │
                            Logic App reads shared mailbox,
                            finds unread invoice PDFs
                                         │
                                         │ POST invoice data
                                         ▼
┌──────────────────────────────────────────────────────────────────┐
│  server.py (HTTP server, PM2 process, port 8088)                │
│                                                                  │
│  POST /  ← receives: {invoice_pdf_base64, from_name,            │
│                        from_email, email_subject,                │
│                        approver_email}                           │
│                                                                  │
│  1. Parse PDF (regex or LLM)                                    │
│  2. Find or create supplier contact in Xero                      │
│  3. Create bill in Xero (Status: SUBMITTED — awaiting approval) │
│  4. Attach original PDF to Xero invoice                          │
│  5. Tell Logic App to send approval email                        │
└──────────────────────────────────────────────────────────────────┘

                        OR (direct pipeline mode)

┌──────────────────────────────────────────────────────────────────┐
│  pipeline.py (runs standalone or via cron)                       │
│                                                                  │
│  1. Connect to Microsoft Graph                                   │
│  2. Fetch unread invoice emails from shared mailbox              │
│  3. Parse PDF attachments (regex, LLM, or hybrid)                │
│  4. Create/update suppliers and bills in Xero                    │
│  5. Send approval email (to approver)                            │
│  6. Send summary report (to all notification recipients)         │
│  7. Mark emails as read                                          │
└──────────────────────────────────────────────────────────────────┘
```

## Files

All files live in the installation directory (e.g., `/path/to/invoice-xero-bridge/`):

| File | Purpose |
|------|---------|
| `pipeline.py` | **Combined pipeline** — reads mailbox, parses PDFs, creates Xero bills, sends notifications |
| `server.py` | **HTTP server** — receives invoice data from Logic App, processes it |
| `cron_trigger.py` | **Cron trigger** — calls your Logic App to check the mailbox |
| `setup_xero_auth.py` | **One-time Xero OAuth setup** — gets tokens |
| `config.json` | **Configuration** — all credentials and settings (fill this in) |
| `ecosystem.config.js` | **PM2 configuration** — for deploying server.py |
| `setup.py` | **Setup wizard** — interactive configuration |
| `requirements.txt` | **Python dependencies** |

## Prerequisites

- **Python 3.10+** (tested with 3.12)
- **pip packages**: `msal`, `requests`, `pdfplumber`
- **Azure App Registration** with Microsoft Graph permissions (Mail.Read, Mail.Send — Application permissions, admin consented)
- **Xero App** at [developer.xero.com](https://developer.xero.com) with redirect URI `http://localhost:8080/callback`
- **Logic App** (optional) — if using server.py workflow for approval emails
- **PM2** (optional) — for running server.py as a daemon

## Installation

```bash
# 1. Set up the directory
git clone <your-repo> /path/to/invoice-xero-bridge
# OR copy the files manually

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. Run the setup wizard
python3 setup.py

# 4. Or configure manually — edit config.json with your credentials
```

## Configuration Guide

There are three major configuration areas:

### 1. Microsoft Graph (Email)

Required for reading the shared mailbox and sending emails.

**Azure Portal steps:**
1. Go to [Azure Portal](https://portal.azure.com) → App Registrations → New Registration
2. Name it (e.g., "InvoiceXero Bridge")
3. Under **API Permissions** → Add permissions:
   - `Microsoft Graph` → **Application permissions**:
     - `Mail.Read` — read the shared mailbox
     - `Mail.Send` — send approval/summary emails
4. Click **Grant Admin Consent** (required for application permissions)
5. Under **Certificates & Secrets** → New Client Secret — copy the value

**config.json fields:**
```json
{
  "graph": {
    "tenant_id": "your-azure-tenant-id",
    "client_id": "your-app-client-id",
    "client_secret": "your-client-secret",
    "shared_mailbox": "accounts@yourcompany.com"
  }
}
```

### 2. LLM Provider (AI Invoice Parsing)

Used when `parsing_mode` is set to `"llm"` or `"hybrid"`. Extracts invoice fields (number, amount, supplier, due date) from PDF text.

**Supported providers:**

| Provider | Base URL | Model Examples |
|----------|----------|----------------|
| `openai` | `https://api.openai.com/v1` | `gpt-4o-mini`, `gpt-4o` |
| `deepseek` | `https://api.deepseek.com/v1` | `deepseek-chat`, `deepseek-reasoner` |
| `anthropic` | `https://api.anthropic.com/v1` | `claude-3-haiku-20240307`, `claude-3-sonnet-20240229` |
| `custom` | Your choice | Any |

**config.json fields:**
```json
{
  "llm": {
    "provider": "openai",
    "model": "gpt-4o-mini",
    "api_key_env": "LLM_API_KEY",
    "base_url": "https://api.openai.com/v1",
    "temperature": 0.0,
    "max_tokens": 500
  }
}
```

The API key is read from the environment variable specified in `api_key_env`. Set it:
```bash
export LLM_API_KEY="sk-..."
# Add to ~/.bashrc for persistence
```

### 3. Xero API

Required for creating contacts and bills in Xero.

**Xero Developer steps:**
1. Go to [Xero Developer](https://developer.xero.com) → My Apps → Create App
2. Choose "Integration" app type
3. Set redirect URI: `http://localhost:8080/callback`
4. Copy **Client ID** and **Client Secret**

**After configuring** `client_id` and `client_secret` in config.json, run:
```bash
python3 setup_xero_auth.py
```
This will:
1. Open a browser for Xero authorization
2. Exchange the code for access/refresh tokens
3. Save tokens to `xero_tokens.json`
4. Auto-discover your Xero organisation (tenant) and save its ID

**config.json fields:**
```json
{
  "xero": {
    "client_id": "your-xero-client-id",
    "client_secret": "your-xero-client-secret",
    "tenant_id": "auto-filled-during-oauth",
    "default_account_code": "200",
    "default_tax_type": "NONE"
  }
}
```

## Parsing Modes

The `parsing_mode` config field controls how invoice data is extracted:

| Mode | Description | Best For |
|------|-------------|----------|
| `regex` | Regex pattern matching only | Known suppliers, fast, no API cost |
| `llm` | AI-powered extraction via configured LLM | Complex/unusual invoice formats |
| `hybrid` | Try regex first, fall back to LLM on failure | General use — best of both |

When using `regex` or `hybrid`, add custom patterns in `pdf_patterns` for your regular suppliers.

## How to Run

### Run the Pipeline (direct invoice processing)

```bash
python3 pipeline.py
```

This fetches unread emails from the shared mailbox, processes any invoice PDFs found, creates Xero bills, and sends notifications.

### Run the HTTP Server (receives data from Logic App)

```bash
python3 server.py
```

Listens on the configured port (default 8088). The Logic App POSTs invoice data here.

### Deploy with PM2

```bash
# Install PM2 if needed
npm install -g pm2

# Start the server
pm2 start ecosystem.config.js
pm2 save
pm2 startup
```

### Set Up Cron for Daily Run

```bash
crontab -e
# Add:
# Daily pipeline run at 9 AM:
0 9 * * * cd /path/to/invoice-xero-bridge && python3 pipeline.py >> /var/log/invoice-xero-bridge.log 2>&1
```

Or use the Logic App workflow:
```bash
# Daily trigger at 9 AM:
0 9 * * * cd /path/to/invoice-xero-bridge && python3 cron_trigger.py >> /var/log/invoice-xero-bridge.log 2>&1
```

## How to Test

### 1. Test LLM Connection

```bash
# Set your API key
export LLM_API_KEY="sk-..."

# Quick LLM test via pipeline
python3 -c "
from pipeline import InvoiceParser
import json
with open('config.json') as f:
    cfg = json.load(f)
parser = InvoiceParser(cfg)
result = parser._parse_llm('Invoice #INV-1234\nTotal: \$1,500.00\nDue: 30/06/2025\nSupplier: Acme Corp')
print(json.dumps(result, indent=2))
"
```

### 2. Test Graph Connection

```bash
python3 -c "
from pipeline import GraphClient
import json
with open('config.json') as f:
    cfg = json.load(f)
g = GraphClient(cfg)
emails = g.fetch_recent_emails(hours=48)
print(f'Found {len(emails)} emails with PDFs in last 48h')
for e in emails:
    print(f'  {e[\"subject\"]} — {len(e[\"pdfs\"])} PDF(s)')
"
```

### 3. Test Full Pipeline (dry run)

```bash
# Dry run — will process what it finds
python3 pipeline.py
```

### 4. Test the HTTP Server

```bash
# Start server in one terminal
python3 server.py

# In another terminal, send a test invoice
python3 -c "
import requests, base64, json
# Create a minimal test PDF (replace with a real one)
with open('test_invoice.pdf', 'rb') as f:
    pdf_b64 = base64.b64encode(f.read()).decode()
resp = requests.post('http://localhost:8088/', json={
    'invoice_pdf_base64': pdf_b64,
    'invoice_pdf_filename': 'test_invoice.pdf',
    'email_subject': 'Invoice from Test Supplier',
    'from_name': 'Test Supplier Pty Ltd',
    'from_email': 'billing@testsupplier.com',
    'approver_email': 'approver@example.com',
})
print(json.dumps(resp.json(), indent=2))
"
```

## How to Add New PDF Patterns

If your regular suppliers have unique invoice formats, add custom regex patterns in `config.json` under `pdf_patterns`:

```json
{
  "pdf_patterns": {
    "invoice_number": [
      "YourCustomRegexHere"
    ],
    "total_amount": [
      "TOTAL\\s+Inc\\.\\s*\\$([\\d,]+\\.\\d{2})"
    ],
    "due_date": [],
    "supplier_name": []
  }
}
```

Patterns are tested in order (custom first, then built-in defaults). The first match wins.

**Pattern tips:**
- Use capture groups `()` to extract the value
- `re.IGNORECASE` and `re.MULTILINE` are enabled automatically
- Test regexes at [regex101.com](https://regex101.com/) with the Python flavor
- For multi-line patterns, use `re.DOTALL` style with `[^]*` instead of `.*`

## HTTP API

### `POST /`
Process an invoice from a Logic App.

**Request body:**
```json
{
  "invoice_pdf_base64": "<base64-encoded PDF>",
  "invoice_pdf_filename": "invoice_123.pdf",
  "email_subject": "Invoice from Supplier XYZ",
  "from_name": "Supplier XYZ Pty Ltd",
  "from_email": "billing@supplierxyz.com",
  "approver_email": "approver@example.com"
}
```

**Success response (200):**
```json
{
  "success": true,
  "invoice_num": "INV-12345",
  "supplier": "Supplier XYZ Pty Ltd",
  "amount": 1500.00,
  "xero_url": "https://go.xero.com/AccountsPayable/Edit/..."
}
```

**Error response (422/500):**
```json
{
  "success": false,
  "error": "Could not extract: invoice_number, total_amount"
}
```

### `GET /health`
Health check endpoint.

```json
{
  "status": "ok",
  "time": "2025-01-15T09:00:00"
}
```

## FAQ / Troubleshooting

### "Xero not authenticated. Run setup_xero_auth.py first."
Your Xero tokens are missing or expired. Run `python3 setup_xero_auth.py` to re-authenticate.

### "LLM extraction failed: 401 Unauthorized"
Your LLM API key is missing or invalid. Check:
- The environment variable defined in `config.json` under `llm.api_key_env` is set
- The key has correct permissions for the model you're using

### "Graph auth failed"
Your Azure app registration may have issues. Check:
- `tenant_id`, `client_id`, and `client_secret` are correct
- Admin consent has been granted for Mail.Read and Mail.Send
- The application permissions are Mail.Read and Mail.Send (not Delegated)

### "Can't access mailbox"
The service principal doesn't have access. Ensure:
- Admin consent was granted for the API permissions
- The shared mailbox exists and the service principal has been granted access via Exchange Online

### "Could not parse PDF (possibly scanned image)"
The invoice is a scanned image without extractable text. Options:
- Enable OCR by installing `pytesseract` and `tesseract-ocr`
- Switch to `"llm"` or `"hybrid"` parsing mode (some LLMs handle image-based PDFs better)
- Process the invoice manually

### "Invoice #XXX already exists in Xero"
Duplicate detection is working. The invoice number already exists in Xero and will be skipped.

### How do I change the application name?
Edit `app_name` in `config.json`. All log messages, email subjects, and report titles will use the new name.

### Can I use this with multiple companies?
Each deployment handles one Xero organisation. For multiple companies, clone the directory and create separate configs.

### How do I update?
```bash
# Pull latest code
git pull

# Reinstall dependencies if changed
pip install -r requirements.txt

# Restart PM2
pm2 restart ecosystem.config.js
```
