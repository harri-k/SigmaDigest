import logging
import os
import json
import pytz
import datetime as dt
from dateutil import tz
import requests
import msal
import azure.functions as func

GRAPH = "https://graph.microsoft.com/v1.0"

def _get_env(name, default=None, required=False):
    val = os.getenv(name, default)
    if required and not val:
        raise RuntimeError(f"Missing required setting: {name}")
    return val

def _get_token():
    tenant_id = _get_env("TENANT_ID", required=True)
    client_id = _get_env("CLIENT_ID", required=True)
    client_secret = _get_env("CLIENT_SECRET", required=True)
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(client_id=client_id, client_credential=client_secret, authority=authority)
    # Application permissions (.default) require admin-consented Graph app permissions (Mail.Read, Mail.Send)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result}")
    return result["access_token"]

KW = {"urgent","invoice","payment","due","meeting","slides","assignment","grade","interview","offer"}
PRIORITY_HINTS = {"washu.edu","registrar","billing","advisor","prof","dean","financial"}

def _score(msg):
    s = 0
    imp = (msg.get("importance") or "normal").lower()
    s += 3 if imp == "high" else (-1 if imp == "low" else 0)
    subj = (msg.get("subject") or "").lower()
    prev = (msg.get("bodyPreview") or "").lower()
    if any(k in subj or k in prev for k in KW): s += 2
    if msg.get("hasAttachments"): s += 1
    sender_addr = (((msg.get("from") or {}).get("emailAddress")) or {}).get("address","").lower()
    if any(h in sender_addr for h in PRIORITY_HINTS): s += 2
    # recency
    now = dt.datetime.now(dt.timezone.utc)
    rdt = dt.datetime.fromisoformat(msg["receivedDateTime"].replace("Z","+00:00"))
    hours = (now - rdt).total_seconds()/3600
    s += 2 if hours <= 6 else (1 if hours <= 24 else 0)
    # directness - if only cc, slightly less important (Graph returns toRecipients/ccRecipients only on expand - omitted here for simplicity)
    return s

def _list_recent_messages(token, user, since_iso):
    h = {"Authorization": f"Bearer {token}"}
    # Filter by receivedDateTime; select only needed fields
    # Note: receivedDateTime filter requires proper datetime format
    url = (f"{GRAPH}/users/{user}/messages"
           f"?$select=id,subject,receivedDateTime,importance,webLink,hasAttachments,from,bodyPreview,conversationId"
           f"&$orderby=receivedDateTime desc&$top=100"
    )
    msgs = []
    r = requests.get(url, headers=h)
    r.raise_for_status()
    data = r.json()
    for m in data.get("value", []):
        if m.get("receivedDateTime", "") >= since_iso:
            msgs.append(m)
    return msgs

def _html_digest(items, date_str):
    rows = []
    for m in items:
        sender = (((m.get("from") or {}).get("emailAddress")) or {}).get("name","Unknown")
        sub = m.get("subject","(no subject)")
        link = m.get("webLink","#")
        prev = (m.get("bodyPreview","") or "").replace("<","&lt;").replace(">","&gt;")[:240]
        rows.append(f"""
        <tr>
          <td style="padding:12px;border-bottom:1px solid #eaeaea">
            <div style="font-weight:600">{sender}</div>
            <div style="color:#111;margin:2px 0 6px">{sub}</div>
            <div style="color:#555;font-size:13px;margin:4px 0 8px">{prev}</div>
            <a href="{link}" style="font-size:13px;text-decoration:none">Open in Outlook →</a>
          </td>
        </tr>
        """)
    html = f"""
    <div style="font-family:Inter,Arial,sans-serif;max-width:720px;margin:0 auto">
      <h2 style="margin:0 0 12px">Daily Digest — {date_str}</h2>
      <table style="width:100%;border-collapse:collapse">{''.join(rows) if rows else '<tr><td>No important emails in the last 24h.</td></tr>'}</table>
      <p style="color:#777;font-size:12px;margin-top:12px">Top items from the last 24 hours.</p>
    </div>
    """
    return html

def _send_mail(token, user, subject, html):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": user}}]
        },
        "saveToSentItems": True
    }
    r = requests.post(f"{GRAPH}/users/{user}/sendMail", headers=h, data=json.dumps(body))
    if r.status_code not in (200, 202, 202):
        logging.error("sendMail failed: %s %s", r.status_code, r.text)

def main(mytimer: func.TimerRequest) -> None:
    logging.info("SigmaDigestTimer triggered")
    try:
        token = _get_token()
        user = _get_env("TARGET_USER", required=True)
        tzname = _get_env("TIMEZONE", "America/Chicago")
        tzinfo = pytz.timezone(tzname)
        now_local = dt.datetime.now(tzinfo)
        since = (now_local - dt.timedelta(days=1)).astimezone(dt.timezone.utc).isoformat()
        msgs = _list_recent_messages(token, user, since)
        scored = sorted(msgs, key=_score, reverse=True)
        max_items = int(_get_env("MAX_ITEMS", "8"))
        top = scored[:max_items]
        html = _html_digest(top, now_local.strftime("%b %d, %Y"))
        _send_mail(token, user, "Your Daily Digest", html)
        logging.info("Digest sent with %d items", len(top))
    except Exception as e:
        logging.exception("Digest failed: %s", e)
