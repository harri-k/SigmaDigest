# Sigma Digest — Outlook Daily Digest (Azure Functions, Python)

A serverless **Azure Function (Timer Trigger)** that fetches the last 24 hours of Outlook emails for a target user, ranks them by importance, and emails a **Daily Digest** back to that user. Uses **Microsoft Graph** with **application permissions** (client credentials).

## Architecture
- **Timer Trigger** (default: 13:30 UTC daily ≈ 7:30 AM US Central in winter)
- **Graph API**: `/users/{TARGET_USER}/messages` and `/users/{TARGET_USER}/sendMail`
- **MSAL** client credentials (`.default` scope) — requires admin consent
- **Scoring** heuristics: importance, keywords, recency, attachments, sender hints

## Quick Start

1. **Create App Registration** (in your dev tenant):
   - API permissions (Application):
     - `Mail.Read` or `Mail.ReadBasic.All`
     - `Mail.Send`
   - Grant **Admin consent**.
   - Create a **Client Secret** and note: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`.

2. **Set Azure Function App Settings** (Configuration):
   - `TENANT_ID`: your tenant GUID
   - `CLIENT_ID`: app registration client id
   - `CLIENT_SECRET`: client secret value
   - `TARGET_USER`: the mailbox to summarize, e.g. `you@yourtenant.onmicrosoft.com`
   - Optional:
     - `TIMEZONE`: default `America/Chicago`
     - `MAX_ITEMS`: default `8`

3. **Deploy**
   - Zip deploy via Azure or GitHub Actions.
   - Runtime: **Python 3.11**, Plan: **Consumption**, OS: **Linux**.

4. **Test locally** (optional):
   - Copy `local.settings.json.example` to `local.settings.json` and fill values.
   - Run the function host:
     ```bash
     func start
     ```

## Repo Layout
```
.
├── host.json
├── requirements.txt
├── local.settings.json.example
└── SigmaDigestTimer
    ├── __init__.py
    └── function.json
```

## Notes
- This sample uses **application permissions**; your app acts as itself and accesses the target mailbox via Graph. Ensure compliance and tenant consent.
- If you prefer **delegated user auth** (per-user consent + refresh tokens), adapt to Auth Code Flow and store tokens in Key Vault. (I can provide that template if needed.)

## Security
- Store secrets in **Azure Key Vault**; reference them in App Settings.
- Limit app permissions to the minimum and restrict mailboxes via Graph App Access Policy if desired.

---

Generated on 2025-10-20T00:23:04.057481Z.
