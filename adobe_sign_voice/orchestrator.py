"""
orchestrator.py — Main entry point for the PSTN Adobe Sign automation.

Modes (pass as CLI flag):
    --send            Read master data and send agreements to all new partners
    --send --dry-run  As above but log only — no API calls made
    --monitor         Poll agreement statuses and process any newly signed agreements
    --reminders       Check and send due chaser emails (7 / 14 / 30 day)
    --widget-monitor  Poll the Adobe Sign widget for new CP submissions (webform process)
    --status          Print a summary of the current state file

Default (no flags):  run --monitor then --reminders (intended for scheduled task)
Webform process:     run --widget-monitor only (no send or reminders needed)

Scheduled task example (Windows Task Scheduler):
    Program:   C:/Users/Public/RPA/code/.venv/Scripts/python.exe
    Arguments: "C:/Users/Public/RPA/code/PSTN Migration/adobe_sign/orchestrator.py"
    Start in:  C:/Users/Public/RPA/code/PSTN Migration/adobe_sign
"""
import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

# ── Logging: write to logs/ and stdout ───────────────────────────────────────
_LOGS_DIR = Path(__file__).parent / "logs"
_LOGS_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)-25s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            _LOGS_DIR / f"orchestrator_{datetime.now(timezone.utc).strftime('%Y%m%d')}.log",
            encoding="utf-8",
        ),
    ],
)
logger = logging.getLogger("orchestrator")


# ── Actions ───────────────────────────────────────────────────────────────────

def action_send(dry_run: bool = False) -> None:
    from bulk_sender import run_bulk_send

    summary = run_bulk_send(dry_run=dry_run)
    logger.info(
        "SEND complete — sent: %d | failed: %d | skipped: %d",
        summary["sent"],
        summary["failed"],
        summary["skipped"],
    )


def action_monitor() -> None:
    from agreement_monitor import run_monitor

    summary = run_monitor()
    logger.info(
        "MONITOR complete — completed: %d | cancelled: %d | pending: %d | errors: %d",
        summary["completed"],
        summary["cancelled"],
        summary["still_pending"],
        summary["errors"],
    )


def action_reminders() -> None:
    from reminder_sender import run_reminders

    summary = run_reminders()
    logger.info(
        "REMINDERS complete — sent: %d | errors: %d",
        summary["reminders_sent"],
        summary["errors"],
    )


def action_widget_monitor() -> None:
    from widget_monitor import run_widget_monitor

    summary = run_widget_monitor()
    logger.info(
        "WIDGET MONITOR complete — processed: %d | skipped: %d | errors: %d",
        summary["processed"],
        summary["skipped"],
        summary["errors"],
    )


def action_status() -> None:
    """Print a human-readable summary of the state file."""
    state_file = Path(__file__).parent / "input" / "agreement_state.json"
    if not state_file.exists():
        print("No state file found — no agreements have been sent yet.")
        return

    with open(state_file, encoding="utf-8") as fh:
        state = json.load(fh)

    total = len(state)
    by_status: dict[str, int] = {}
    reminders: dict[str, int] = {"7": 0, "14": 0, "30": 0}
    processed = 0

    for r in state:
        s = r.get("status", "UNKNOWN")
        by_status[s] = by_status.get(s, 0) + 1
        if r.get("reminder_7_sent"):
            reminders["7"] += 1
        if r.get("reminder_14_sent"):
            reminders["14"] += 1
        if r.get("reminder_30_sent"):
            reminders["30"] += 1
        if r.get("processed"):
            processed += 1

    print("=" * 60)
    print("  Adobe Sign Agreement Status Summary")
    print("=" * 60)
    print(f"  Total agreements tracked : {total}")
    for status, count in sorted(by_status.items()):
        print(f"  {status:<30}: {count}")
    print(f"\n  Reminders sent (day 7)   : {reminders['7']}")
    print(f"  Reminders sent (day 14)  : {reminders['14']}")
    print(f"  Reminders sent (day 30)  : {reminders['30']}")
    print(f"\n  Fully processed          : {processed}")
    print("=" * 60)

    # List unprocessed completed agreements
    unprocessed_complete = [
        r for r in state
        if r.get("status") not in ("OUT_FOR_SIGNATURE",) and not r.get("processed")
    ]
    if unprocessed_complete:
        print(f"\n  ⚠️   {len(unprocessed_complete)} completed agreement(s) not yet processed:")
        for r in unprocessed_complete:
            print(f"       • {r['partner_name']} — {r['status']} — {r['agreement_id']}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="PSTN Migration — Adobe Sign Automation Orchestrator"
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Read master data and send agreements to all new partners.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Used with --send: log actions without making API calls.",
    )
    parser.add_argument(
        "--monitor",
        action="store_true",
        help="Poll agreement statuses and process any newly signed agreements.",
    )
    parser.add_argument(
        "--reminders",
        action="store_true",
        help="Check and send due chaser emails.",
    )
    parser.add_argument(
        "--widget-monitor",
        action="store_true",
        help="Poll the Adobe Sign widget for new CP submissions (webform process).",
    )
    parser.add_argument(
        "--status",
        action="store_true",
        help="Print a summary of the current state file.",
    )
    args = parser.parse_args()

    # Default (no flags) → monitor + reminders
    run_default = not any([args.send, args.monitor, args.reminders, args.status, args.widget_monitor])

    logger.info("PSTN Adobe Sign Orchestrator started — %s", datetime.now(timezone.utc).isoformat())

    if args.status:
        action_status()
        return

    if args.send:
        action_send(dry_run=args.dry_run)

    if args.monitor or run_default:
        action_monitor()

    if args.reminders or run_default:
        action_reminders()

    if args.widget_monitor:
        action_widget_monitor()

    logger.info("Orchestrator finished.")


if __name__ == "__main__":
    main()
