"""
process_single_agreement.py - Manually process a specific agreement by ID

Usage:
    python adobe_sign_voice/process_single_agreement.py AGREEMENT_ID
    
Example:
    python adobe_sign_voice/process_single_agreement.py CBJCHBCAABAA4EVNcW4jvHO92KGByNMvJVazKSJaS2hW
"""
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

from adobe_sign_client import AdobeSignClient
from response_processor import process_signed_agreement
from widget_monitor import (
    _WIDGET_STATE_FILE,
    _get_signer_email,
    _get_signer_name,
    _load_widget_state,
    _save_widget_state,
)
from sharepoint_reader import try_write_to_audit_tab
from attachment_validator import flush_master_writes
from config import UPDATE_MASTER_DATA

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
logger = logging.getLogger(__name__)


def process_agreement_by_id(agreement_id: str) -> None:
    """Manually process a specific agreement that wasn't picked up by the widget monitor."""
    
    logger.info("=" * 70)
    logger.info("Processing agreement: %s", agreement_id)
    logger.info("=" * 70)
    
    client = AdobeSignClient()
    widget_state = _load_widget_state()
    
    # Check if already processed
    if agreement_id in widget_state:
        logger.warning("⚠️  Agreement %s has already been processed!", agreement_id)
        logger.info("Processed at: %s", widget_state[agreement_id].get("processed_at"))
        logger.info("Signer: %s <%s>", 
                   widget_state[agreement_id].get("signer_name"),
                   widget_state[agreement_id].get("signer_email"))
        
        response = input("\nProcess anyway? (y/N): ")
        if response.lower() != 'y':
            logger.info("Cancelled.")
            return
    
    # Fetch agreement details
    try:
        logger.info("Fetching agreement details from Adobe Sign...")
        agreement_detail = client.get_agreement(agreement_id)
    except Exception as exc:
        logger.error("❌ Failed to fetch agreement: %s", exc)
        return
    
    status = agreement_detail.get("status", "")
    signer_name = _get_signer_name(agreement_detail)
    signer_email = _get_signer_email(agreement_detail)
    
    logger.info("Agreement status: %s", status)
    logger.info("Signer: %s <%s>", signer_name, signer_email)
    
    # Process the agreement
    try:
        logger.info("Processing agreement...")
        result = process_signed_agreement(agreement_id, signer_name)
    except Exception as exc:
        logger.error("❌ Processing failed: %s", exc, exc_info=True)
        return
    
    # Build the audit record
    audit_data = {
        "signer_name": result.get("reseller_name") or signer_name,
        "signer_email": signer_email,
        "account_ref": result.get("account_number", ""),
        "migration_status": status,
        "opt_out_full": result.get("opt_out_full", False),
        "rwp_remain": result.get("rwp_remain", False),
        "confirm_amendments": result.get("confirm_amendments", False),
        "processed_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
    }
    
    # Write to Audit tab
    logger.info("Writing to Audit tab...")
    try_write_to_audit_tab(audit_data, agreement_id)
    
    # Save to widget state
    widget_state[agreement_id] = {
        "signer_name": signer_name,
        "signer_email": signer_email,
        "status": status,
        "opt_out_full": result.get("opt_out_full", False),
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "errors": result.get("errors", []),
        "manually_processed": True,
    }
    _save_widget_state(widget_state)
    logger.info("✅ Widget state updated")
    
    # Flush master writes
    if UPDATE_MASTER_DATA:
        logger.info("Flushing master Excel writes...")
        try:
            flush_master_writes()
        except Exception as exc:
            logger.error("flush_master_writes raised unexpectedly: %s", exc, exc_info=True)
    
    logger.info("=" * 70)
    logger.info("✅ Agreement processed successfully!")
    logger.info("=" * 70)
    logger.info("Reseller: %s", result.get("reseller_name", signer_name))
    logger.info("Account: %s", result.get("account_number", "N/A"))
    logger.info("Full Opt-Out: %s", "Yes" if result.get("opt_out_full") else "No")
    
    if result.get("is_partial_opt_out"):
        logger.info("⚠️  Partial opt-out detected")
    if result.get("has_data_quality_issues"):
        logger.warning("⚠️  Data quality issues detected - check the report")
    

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python adobe_sign_voice/process_single_agreement.py AGREEMENT_ID")
        print("\nExample:")
        print("  python adobe_sign_voice/process_single_agreement.py CBJCHBCAABAA4EVNcW4jvHO92KGByNMvJVazKSJaS2hW")
        sys.exit(1)
    
    agreement_id = sys.argv[1].strip()
    process_agreement_by_id(agreement_id)
