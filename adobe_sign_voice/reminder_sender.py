"""
reminder_sender.py — Sends chaser emails at 7, 14 and 30 days after initial send.

Emails are sent via Microsoft Graph API using the shared PSTN mailbox,
following the same pattern as graph_mailbox_check.py.

The signing URL for each partner is retrieved from Adobe Sign and embedded
in the chaser email to make it easy for the partner to complete the form.
"""
import json
import logging
from datetime import datetime, timezone

import msal
import requests

from adobe_sign_client import AdobeSignClient, AdobeSignError
from config import (
    GRAPH_CLIENT_ID,
    GRAPH_CLIENT_SECRET,
    GRAPH_SENDER_MAILBOX,
    GRAPH_TENANT_ID,
    LOGS_DIR,
    REMINDER_DAYS,
    STATE_FILE,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOGS_DIR / "reminder_sender.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

_GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


# ── Graph auth ────────────────────────────────────────────────────────────────

def _get_graph_token() -> str:
    app = msal.ConfidentialClientApplication(
        GRAPH_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
        client_credential=GRAPH_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=_GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(
            f"Graph API authentication failed: {result.get('error_description')}"
        )
    return result["access_token"]


# ── Email builder ─────────────────────────────────────────────────────────────

def _build_chaser_email(partner_name: str, day: int, signing_url: str) -> dict:
    """Return a Graph API message payload for a chaser email."""
    if signing_url:
        action_block = (
            f'<p style="margin:20px 0;">'
            f'<a href="{signing_url}" style="background:#0066cc;color:#fff;padding:10px 20px;'
            f'text-decoration:none;border-radius:4px;font-weight:bold;">Complete Your Form</a>'
            f"</p>"
        )
    else:
        action_block = (
            "<p>Please check your inbox for the original Adobe Sign email and "
            "click the link to complete your form.</p>"
        )

    if day == 7:
        urgency = "a gentle reminder"
        deadline_note = "Please aim to complete this within the next week."
    elif day == 14:
        urgency = "a second reminder"
        deadline_note = "It's important that you complete this as soon as possible to avoid delays to your migration planning."
    else:
        urgency = "a final reminder"
        deadline_note = (
            "<strong>This is your final reminder.</strong> If you do not respond, "
            "your account may be included in the default Responsible Wholesaler migration plan. "
            "Please contact your BDM if you need assistance."
        )

    html_body = f"""
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
<p>Dear {partner_name},</p>

<p>This is {urgency} regarding your <strong>PSTN Migration Opt-Out / Opt-In form</strong>
that was sent to you via Adobe Sign.</p>

<p>We have not yet received a completed response from you.</p>

<h3>What you need to do</h3>
<ol>
  <li>Click the button below (or use the link in your original Adobe Sign email)</li>
  <li>Review your product mapping</li>
  <li>Select your migration preference (Opt-Out, Partial Opt-Out, or Opt-In)</li>
  <li>Complete all required fields and e-sign the form</li>
</ol>

{action_block}

<p>{deadline_note}</p>

<h3>Migration options reminder</h3>
<ul>
  <li><strong>Full Opt-Out</strong> — manage all migrations independently</li>
  <li><strong>Partial Opt-Out</strong> — manage only specific lines independently</li>
  <li><strong>Full Opt-In</strong> — confirm end-user notification date so we can allocate your migration window</li>
</ul>

<p>If you believe you have already completed this form, please disregard this message.</p>

<p>If you have any questions, please contact your Gamma BDM.</p>

<p>Kind regards,<br>
<strong>Gamma PSTN Migration Team</strong></p>

<hr style="margin-top:30px;border:none;border-top:1px solid #ccc;">
<p style="font-size:11px;color:#888;">
This is an automated message. Please do not reply directly to this email.
</p>
</body></html>
"""
    return {
        "message": {
            "subject": f"Action Required: PSTN Migration Opt-Out Form — Reminder {day} Days",
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [
                {"emailAddress": {"address": ""}}  # filled in by caller
            ],
        }
    }


# ── State helpers ─────────────────────────────────────────────────────────────

def _load_state() -> list[dict]:
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    return []


def _save_state(state: list[dict]) -> None:
    with open(STATE_FILE, "w", encoding="utf-8") as fh:
        json.dump(state, fh, indent=2, ensure_ascii=False)


def _days_since(iso_date: str) -> float:
    sent = datetime.fromisoformat(iso_date)
    if sent.tzinfo is None:
        sent = sent.replace(tzinfo=timezone.utc)
    delta = datetime.now(timezone.utc) - sent
    return delta.total_seconds() / 86400


# ── Send a single chaser ──────────────────────────────────────────────────────

def _send_chaser(token: str, record: dict, day: int, signing_url: str) -> bool:
    """Send one chaser email via Graph. Returns True on success."""
    payload = _build_chaser_email(record["partner_name"], day, signing_url)
    payload["message"]["toRecipients"][0]["emailAddress"]["address"] = record["partner_email"]

    endpoint = f"https://graph.microsoft.com/v1.0/users/{GRAPH_SENDER_MAILBOX}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    response = requests.post(endpoint, headers=headers, json=payload, timeout=30)
    if response.status_code == 202:
        logger.info(
            "✅  Day-%d chaser sent to %s <%s>",
            day,
            record["partner_name"],
            record["partner_email"],
        )
        return True
    logger.error(
        "❌  Failed to send day-%d chaser to %s: HTTP %s — %s",
        day,
        record["partner_name"],
        response.status_code,
        response.text,
    )
    return False


# ── Main ──────────────────────────────────────────────────────────────────────

def run_reminders() -> dict:
    """
    Check all pending agreements and send any due reminder emails.

    Returns:
        Summary dict: {reminders_sent, skipped, errors}
    """
    logger.info("=" * 60)
    logger.info("Reminder sender started  |  schedule: day %s", REMINDER_DAYS)
    logger.info("=" * 60)

    state = _load_state()
    pending = [r for r in state if r.get("status") == "OUT_FOR_SIGNATURE"]

    if not pending:
        logger.info("No pending agreements — nothing to remind.")
        return {"reminders_sent": 0, "skipped": 0, "errors": 0}

    token = _get_graph_token()
    sign_client = AdobeSignClient()

    reminders_sent = 0
    skipped = 0
    errors = 0
    state_dirty = False

    for record in pending:
        days_elapsed = _days_since(record["sent_date"])

        # Determine which reminder thresholds have been crossed but not yet sent
        for day in sorted(REMINDER_DAYS):
            flag_key = f"reminder_{day}_sent"
            if record.get(flag_key):
                continue  # already sent
            if days_elapsed < day:
                continue  # not due yet

            # Retrieve the signing URL to embed in the email
            signing_url = ""
            try:
                signing_url = sign_client.get_signing_url(record["agreement_id"]) or ""
            except AdobeSignError as exc:
                logger.debug("Could not retrieve signing URL: %s", exc)

            if _send_chaser(token, record, day, signing_url):
                record[flag_key] = True
                state_dirty = True
                reminders_sent += 1
            else:
                errors += 1

    if state_dirty:
        _save_state(state)

    summary = {"reminders_sent": reminders_sent, "skipped": skipped, "errors": errors}
    logger.info(
        "Reminder run complete — sent: %d | errors: %d",
        reminders_sent,
        errors,
    )
    return summary


if __name__ == "__main__":
    run_reminders()
