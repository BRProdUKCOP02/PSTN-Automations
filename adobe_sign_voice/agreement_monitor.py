"""
agreement_monitor.py — Polls Adobe Sign for status changes on all pending agreements.

For each agreement in state with status OUT_FOR_SIGNATURE:
  - Fetches current status from Adobe Sign
  - If SIGNED / COMPLETED  → updates state and triggers response_processor
  - If CANCELLED / DECLINED → updates state and logs for manual review
  - Otherwise               → no change (still awaiting signature)
"""
import json
import logging
from datetime import datetime, timezone

from adobe_sign_client import AdobeSignClient, AdobeSignError
from config import LOGS_DIR, STATE_FILE

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOGS_DIR / "agreement_monitor.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

# Statuses that mean the agreement is done (signed/approved)
COMPLETED_STATUSES = {"SIGNED", "APPROVED", "ACCEPTED", "FORM_FILLED", "DELIVERED"}
# Statuses that mean the agreement ended without completion
CANCELLED_STATUSES = {"CANCELLED", "DECLINED", "EXPIRED", "REJECTED"}


def _load_state() -> list[dict]:
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    return []


def _save_state(state: list[dict]) -> None:
    with open(STATE_FILE, "w", encoding="utf-8") as fh:
        json.dump(state, fh, indent=2, ensure_ascii=False)


def run_monitor() -> dict:
    """
    Poll Adobe Sign and update state for all pending agreements.

    Returns:
        Summary dict: {completed, cancelled, still_pending, errors}
    """
    logger.info("=" * 60)
    logger.info("Agreement monitor started")
    logger.info("=" * 60)

    state = _load_state()
    pending = [r for r in state if r.get("status") == "OUT_FOR_SIGNATURE"]

    logger.info("Checking %d pending agreement(s).", len(pending))

    if not pending:
        logger.info("No pending agreements to check.")
        return {"completed": 0, "cancelled": 0, "still_pending": 0, "errors": 0}

    client = AdobeSignClient()
    completed = 0
    cancelled = 0
    still_pending = 0
    errors = 0

    for record in pending:
        agreement_id = record["agreement_id"]
        name = record["partner_name"]

        try:
            agreement = client.get_agreement(agreement_id)
            status = agreement.get("status", "UNKNOWN")

            if status in COMPLETED_STATUSES:
                record["status"] = status
                record["completed_date"] = datetime.now(timezone.utc).isoformat()
                logger.info("✅  COMPLETED  — %s  (%s)", name, agreement_id)
                completed += 1

                # Trigger response processing immediately
                try:
                    from response_processor import process_signed_agreement
                    process_signed_agreement(agreement_id, name)
                except Exception as proc_exc:
                    logger.error(
                        "⚠️   Response processing failed for %s: %s", name, proc_exc
                    )

            elif status in CANCELLED_STATUSES:
                record["status"] = status
                record["completed_date"] = datetime.now(timezone.utc).isoformat()
                logger.warning(
                    "⚠️   %s — %s  (%s) — manual review required", status, name, agreement_id
                )
                cancelled += 1

            else:
                logger.debug("Pending (%s) — %s  (%s)", status, name, agreement_id)
                still_pending += 1

        except AdobeSignError as exc:
            logger.error("❌  API error checking %s (%s): %s", name, agreement_id, exc)
            errors += 1

    _save_state(state)

    summary = {
        "completed": completed,
        "cancelled": cancelled,
        "still_pending": still_pending,
        "errors": errors,
    }
    logger.info(
        "Monitor complete — completed: %d | cancelled/declined: %d | still pending: %d | errors: %d",
        completed,
        cancelled,
        still_pending,
        errors,
    )
    return summary


if __name__ == "__main__":
    run_monitor()
