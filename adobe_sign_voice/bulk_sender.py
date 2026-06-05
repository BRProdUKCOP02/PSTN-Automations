"""
bulk_sender.py — Reads master data and sends individual Adobe Sign agreements.

State is persisted in input/agreement_state.json so:
  • Partners already sent are never re-sent
  • The file is the single source of truth for monitor/reminder modules

State schema per record:
    {
        "agreement_id":      "<adobe-sign-id>",
        "partner_name":      "Acme Ltd",
        "partner_email":     "shared@acme.co.uk",
        "account_ref":       "ACC123",
        "sent_date":         "2026-04-28T10:00:00",
        "status":            "OUT_FOR_SIGNATURE",
        "reminder_{n}_sent": false  (one key per day in REMINDER_DAYS),
        "completed_date":    null,
        "processed":         false
    }
"""
import json
import logging
import time
from datetime import datetime, timezone

from adobe_sign_client import AdobeSignClient, AdobeSignError
from config import (
    ADOBE_SIGN_AGREEMENT_NAME_PREFIX,
    ADOBE_SIGN_LIBRARY_DOC_ID,
    BATCH_SIZE,
    LOGS_DIR,
    REMINDER_DAYS,
    SEND_DELAY_SECONDS,
    STATE_FILE,
)
from sharepoint_reader import build_merge_fields, load_master_data, try_update_excel_after_send

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOGS_DIR / "bulk_sender.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)


# ── State helpers ─────────────────────────────────────────────────────────────

def _load_state() -> list[dict]:
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    return []


def _save_state(state: list[dict]) -> None:
    with open(STATE_FILE, "w", encoding="utf-8") as fh:
        json.dump(state, fh, indent=2, ensure_ascii=False)
    logger.debug("State saved (%d records).", len(state))


def _already_sent(partner: dict) -> bool:
    """Returns True if date_sent is already filled in the Excel row."""
    return bool(partner.get("date_sent", "").strip())


# ── Main ──────────────────────────────────────────────────────────────────────

def run_bulk_send(dry_run: bool = False) -> dict:
    """
    Send Adobe Sign agreements for all partners not yet in the state file.

    Args:
        dry_run: If True, log what would be sent but do not call the API.

    Returns:
        Summary dict with counts.
    """
    logger.info("=" * 60)
    logger.info("Bulk sender started%s", " [DRY RUN]" if dry_run else "")
    logger.info("=" * 60)

    partners = load_master_data()
    state = _load_state()

    # Identify partners that haven't been sent yet (date_sent blank in Excel)
    pending = [p for p in partners if not _already_sent(p)]
    logger.info(
        "%d total partners | %d already sent | %d pending",
        len(partners),
        len(partners) - len(pending),
        len(pending),
    )

    if not pending:
        logger.info("Nothing to send. All partners have existing agreements.")
        return {"sent": 0, "failed": 0, "skipped": len(partners)}

    # Apply optional batch size cap
    batch = pending if BATCH_SIZE == 0 else pending[:BATCH_SIZE]
    if BATCH_SIZE > 0:
        logger.info("Batch size capped at %d.", BATCH_SIZE)

    client = AdobeSignClient()
    sent = 0
    failed = 0
    failed_details: list[str] = []

    for partner in batch:
        name = partner["partner_name"]
        email = partner["partner_email"]
        excel_row = partner.get("_excel_row")

        agreement_name = f"{ADOBE_SIGN_AGREEMENT_NAME_PREFIX} — {name}"
        merge_fields = build_merge_fields(partner)

        if dry_run:
            logger.info("[DRY RUN] Would send to: %s <%s>", name, email)
            sent += 1
            continue

        try:
            agreement_id = client.create_agreement(
                library_document_id=ADOBE_SIGN_LIBRARY_DOC_ID,
                recipient_email=email,
                recipient_name=name,
                agreement_name=agreement_name,
                merge_fields=merge_fields,
                message=(
                    "Please review your PSTN migration options and complete "
                    "this form at your earliest convenience."
                ),
            )
            state.append(
                {
                    "agreement_id": agreement_id,
                    "partner_name": name,
                    "partner_email": email,
                    "sent_date": datetime.now(timezone.utc).isoformat(),
                    "status": "OUT_FOR_SIGNATURE",
                    **{f"reminder_{d}_sent": False for d in REMINDER_DAYS},
                    "completed_date": None,
                    "processed": False,
                }
            )
            # Persist after every successful send — avoids data loss mid-batch
            _save_state(state)
            # Write agreement_id and date_sent back to the Excel file
            if excel_row:
                try_update_excel_after_send(excel_row, agreement_id)
            logger.info("✅  Sent to %-40s  agreement: %s", f"{name} <{email}>", agreement_id)
            sent += 1

        except AdobeSignError as exc:
            logger.error("❌  Failed for %s <%s>: %s", name, email, exc)
            failed += 1
            failed_details.append(f"{name}: {exc}")

        time.sleep(SEND_DELAY_SECONDS)

    summary = {"sent": sent, "failed": failed, "skipped": len(partners) - len(pending)}
    logger.info(
        "Bulk send complete — sent: %d | failed: %d | skipped (already sent): %d",
        sent,
        failed,
        summary["skipped"],
    )
    if failed_details:
        logger.warning("Failed partners:\n%s", "\n".join(f"  • {d}" for d in failed_details))

    return summary


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Send PSTN opt-out Adobe Sign agreements in bulk.")
    parser.add_argument("--dry-run", action="store_true", help="Log actions without calling the API.")
    args = parser.parse_args()
    run_bulk_send(dry_run=args.dry_run)
