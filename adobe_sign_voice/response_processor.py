"""
response_processor.py — Processes a completed (signed) Adobe Sign agreement.

For each signed agreement this module:
  1. Downloads form data (GET /agreements/{id}/formData) → JSON
  2. Downloads any file attachments (GET /agreements/{id}/documents)
     — identifies Excel attachments uploaded by the signer (product mapping)
     — saves raw bytes to output/{partner_name}_attachment_{timestamp}.xlsx
     — parses the Excel into a pandas DataFrame and saves as CSV
  3. Downloads the combined signed PDF
  4. Saves a summary JSON to output/{partner_name}_form_data_{timestamp}.json
  5. Marks the record as processed=True in state
"""
import base64
import json
import logging
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import msal
import pandas as pd
import requests

from adobe_sign_client import AdobeSignClient, AdobeSignError
from attachment_validator import AttachmentValidationResult, validate_attachment, queue_full_opt_out_master_updates
from config import (
    CIRCUIT_MASTER_PATH,
    CONVERT_ATTACHMENT_TO_CSV,
    DATA_ALERT_EMAIL,
    FIELD_ACCOUNT_NUMBER,
    FIELD_CONFIRM_AMENDMENTS,
    FIELD_CONFIRM_RWP_REMAIN,
    FIELD_OPT_OUT_FULL,
    FIELD_RESELLER_NAME,
    GRAPH_CLIENT_ID,
    GRAPH_CLIENT_SECRET,
    GRAPH_SENDER_MAILBOX,
    GRAPH_TENANT_ID,
    LOGS_DIR,
    OPT_OUT_FORWARD_EMAIL,
    OUTPUT_DIR,
    PA_ATTACHMENT_DROP_FOLDER,
    PARTIAL_OPT_OUT_EMAIL,
    PS_SUMMARY_EMAIL,
    STATE_FILE,
)
from sharepoint_reader import try_update_excel_after_completion

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOGS_DIR / "response_processor.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)


def _parse_email_list(email_config: str) -> list[dict]:
    """Convert comma-separated emails to Graph API recipient format.
    
    Args:
        email_config: Single email or comma-separated list of emails
        
    Returns:
        List of recipient dicts in Graph API format: [{"emailAddress": {"address": "..."}}, ...]
    """
    if not email_config:
        return []
    return [
        {"emailAddress": {"address": addr.strip()}}
        for addr in email_config.split(",")
        if addr.strip()
    ]


def _safe_filename(name: str) -> str:
    """Strip characters unsafe for filenames."""
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def _load_state() -> list[dict]:
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    return []


def _save_state(state: list[dict]) -> None:
    with open(STATE_FILE, "w", encoding="utf-8") as fh:
        json.dump(state, fh, indent=2, ensure_ascii=False)


def _timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")


def _is_excel(doc: dict) -> bool:
    """Return True if the document looks like an Excel attachment."""
    name = doc.get("name", "").lower()
    mime = doc.get("mimeType", "").lower()
    return (
        name.endswith((".xlsx", ".xls", ".xlsm"))
        or "spreadsheet" in mime
        or "excel" in mime
        or "officedocument.spreadsheetml" in mime
    )


def _is_ticked(value: str) -> bool:
    """Return True if an Adobe Sign checkbox value represents a ticked state."""
    return str(value).strip().lower() in {"on", "yes", "true", "checked", "x", "1"}


def _get_graph_token() -> str:
    """Acquire a Graph API access token via client credentials."""
    app = msal.ConfidentialClientApplication(
        GRAPH_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
        client_credential=GRAPH_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(
            f"Graph API authentication failed: {result.get('error_description')}"
        )
    return result["access_token"]


def _forward_opt_out(
    token: str,
    reseller_name: str,
    account_number: str,
    agreement_id: str,
    signed_pdf_path: Optional[str],
    attachment_paths: list,
) -> None:
    """Send a notification email with signed PDF and asset register to the opt-out inbox."""

    graph_attachments = []

    if signed_pdf_path and Path(signed_pdf_path).exists():
        data = Path(signed_pdf_path).read_bytes()
        graph_attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": Path(signed_pdf_path).name,
            "contentType": "application/pdf",
            "contentBytes": base64.b64encode(data).decode(),
        })

    for att_path in attachment_paths:
        p = Path(att_path)
        if not p.exists():
            continue
        if p.suffix == ".xlsx":
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif p.suffix == ".csv":
            mime = "text/csv"
        else:
            mime = "application/octet-stream"
        data = p.read_bytes()
        graph_attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": p.name,
            "contentType": mime,
            "contentBytes": base64.b64encode(data).decode(),
        })

    has_register = len(graph_attachments) > 1
    html_body = (
        '<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333;">'
        "<p>A partner has selected <strong>Full Opt-Out</strong> on their "
        "PSTN Migration Acknowledgment Form.</p>"
        '<table style="border-collapse:collapse;margin:16px 0;">'
        f'<tr><td style="padding:4px 16px 4px 0;font-weight:bold;">Reseller Name</td>'
        f"<td>{reseller_name}</td></tr>"
        f'<tr><td style="padding:4px 16px 4px 0;font-weight:bold;">Gamma Account Number</td>'
        f"<td>{account_number}</td></tr>"
        f'<tr><td style="padding:4px 16px 4px 0;font-weight:bold;">Agreement ID</td>'
        f"<td>{agreement_id}</td></tr>"
        "</table>"
        f'<p>The signed form{" and asset register" if has_register else ""} '
        f'{"are" if has_register else "is"} attached.</p>'
        '<hr style="margin-top:30px;border:none;border-top:1px solid #ccc;">'
        '<p style="font-size:11px;color:#888;">This is an automated notification '
        "from the PSTN Migration Adobe Sign automation.</p>"
        "</body></html>"
    )

    payload = {
        "message": {
            "subject": f"Full Opt-Out Received — {reseller_name}",
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": _parse_email_list(OPT_OUT_FORWARD_EMAIL),
            "attachments": graph_attachments,
        }
    }

    endpoint = f"https://graph.microsoft.com/v1.0/users/{GRAPH_SENDER_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.post(endpoint, headers=headers, json=payload, timeout=60)
    if response.status_code == 202:
        logger.info(
            "✅  Opt-out forward sent to %s for %s (%d attachment(s))",
            OPT_OUT_FORWARD_EMAIL, reseller_name, len(graph_attachments),
        )
    else:
        logger.error(
            "❌  Failed to forward opt-out for %s: HTTP %s — %s",
            reseller_name, response.status_code, response.text,
        )


def _build_alert_recipients() -> list[dict]:
    """Return deduplicated list of Graph API recipient dicts for CTL data quality alerts.

    Only uses DATA_ALERT_EMAIL — other email lists (partial opt-out, opt-out forward,
    PS summary) have their own dedicated send functions.
    """
    seen: set[str] = set()
    recipients = []
    if DATA_ALERT_EMAIL:
        for addr in DATA_ALERT_EMAIL.split(","):
            addr = addr.strip()
            if addr and addr not in seen:
                seen.add(addr)
                recipients.append({"emailAddress": {"address": addr}})
    return recipients


def _send_data_alert_email(
    token: str,
    agreement_id: str,
    reseller_name: str,
    validation_results: list[AttachmentValidationResult],
    report_path: Optional[str],
) -> None:
    """
    Send a data quality alert to all configured alert email addresses.

    Triggered when any attachment has critical issues, changes, or PDF human-review flags.
    Attaches the data quality Excel report.
    """
    recipients = _build_alert_recipients()
    if not recipients:
        logger.warning("No alert email recipients configured — data quality alert not sent.")
        return

    # Determine banner severity — only real data issues raise CTL-critical
    # (pdf_human_review_required alone is not a data quality failure)
    any_critical = any(
        v.critical_count > 0 or (bool(v.missing_columns) and not v.pdf_human_review_required)
        for v in validation_results
    )
    any_pdf_review = any(v.pdf_human_review_required for v in validation_results)

    if any_critical:
        banner_colour = "#CC0000"
        banner_bg = "#FFCCCC"
        severity_label = "DATA QUALITY ISSUES — IMMEDIATE ACTION REQUIRED"
    else:
        banner_colour = "#B25900"
        banner_bg = "#FFE5B4"
        severity_label = "DATA CHANGES DETECTED — REVIEW REQUIRED"

    # Build critical issues rows
    critical_rows_html = ""
    for val in validation_results:
        for rr in val.row_results:
            if not rr.is_critical:
                continue
            for issue_type, detail in [
                ("Circuit Not Found", f"circuit_id '{rr.circuit_id}' not in master. POSSIBLE WRONG-RECORD SUBMISSION.") if not rr.found_in_master else (None, None),
            ] + [
                ("Immutable Mismatch", f"Field '{col}': expected '{d['expected']}', got '{d['actual']}'. May indicate wrong customer record.")
                for col, d in rr.immutable_mismatches.items()
            ] + [
                ("Missing Required Field", f"Field '{col}' is blank. " + ("Customer contact email is missing." if col == "email address" else "Required for safe processing."))
                for col in rr.missing_required
            ] + [
                ("Invalid Email — FIX NEEDED", f"Email '{e}' is not a valid email address. The master has NOT been updated. Fix the email and re-submit.")
                for e in rr.invalid_emails
            ]:
                if issue_type is None:
                    continue
                critical_rows_html += (
                    f"<tr style='background:#FFCCCC;'>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{rr.row_index + 1}</td>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{rr.circuit_id}</td>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc;font-weight:bold'>{issue_type}</td>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{detail}</td>"
                    "</tr>"
                )

    # Build a dedicated "fix needed" section for invalid emails
    has_invalid_email_fix = any(
        bool(rr.invalid_emails)
        for val in validation_results
        for rr in val.row_results
    )
    invalid_email_rows_html = ""
    if has_invalid_email_fix:
        for val in validation_results:
            for rr in val.row_results:
                for bad_email in rr.invalid_emails:
                    invalid_email_rows_html += (
                        f"<tr style='background:#FFCCCC;'>"
                        f"<td style='padding:4px 8px;border:1px solid #ccc'>{rr.row_index + 1}</td>"
                        f"<td style='padding:4px 8px;border:1px solid #ccc'>{rr.circuit_id}</td>"
                        f"<td style='padding:4px 8px;border:1px solid #ccc;font-weight:bold;color:#CC0000'>{bad_email}</td>"
                        f"<td style='padding:4px 8px;border:1px solid #ccc'>Not a valid email address. "
                        f"The master has <strong>NOT</strong> been updated for this circuit. "
                        f"Correct the email address and re-submit.</td>"
                        "</tr>"
                    )

    # Build changed fields rows
    changed_rows_html = ""
    for val in validation_results:
        for rr in val.row_results:
            for col, change in rr.mutable_changes.items():
                changed_rows_html += (
                    f"<tr style='background:#FFE5B4;'>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{rr.row_index + 1}</td>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{rr.circuit_id}</td>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{col}</td>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{change['old']}</td>"
                    f"<td style='padding:4px 8px;border:1px solid #ccc'>{change['new']}</td>"
                    "</tr>"
                )

    pdf_notice = ""
    if any_pdf_review:
        pdf_notice = (
            "<div style='background:#FFF3CD;border:2px solid #FFA500;padding:12px;margin:12px 0;border-radius:4px;'>"
            "<strong>⚠ PDF SUBMISSION — HUMAN REVIEW REQUIRED</strong><br>"
            "One or more attachments were submitted as PDF files. The data has been extracted "
            "and saved as a CSV file for your review. It has <strong>NOT</strong> been validated "
            "against the circuit master. You must review and confirm the data before it is used "
            "in any downstream process or migration order."
            "</div>"
        )

    html_body = (
        '<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333;">'
        f"<div style='background:{banner_bg};border-left:6px solid {banner_colour};padding:12px 16px;margin-bottom:16px;'>"
        f"<strong style='color:{banner_colour};font-size:16px;'>{severity_label}</strong><br>"
        f"<span>Reseller: <strong>{reseller_name}</strong> &nbsp;|&nbsp; Agreement: <strong>{agreement_id}</strong></span>"
        "</div>"
        "<p>A data quality issue has been identified in a CP-submitted attachment for the "
        "PSTN Voice Migration opt-out process. Please review the attached data quality report "
        "and take appropriate action.</p>"
        + pdf_notice
        + (
            "<h3 style='color:#CC0000;'>Critical Issues</h3>"
            '<table style="border-collapse:collapse;width:100%;margin-bottom:16px;">'
            "<tr style='background:#f0f0f0;'>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Row</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Circuit ID</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Issue Type</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Detail</th>"
            "</tr>"
            f"{critical_rows_html or '<tr><td colspan=4 style=padding:4px>None</td></tr>'}"
            "</table>"
            if critical_rows_html else ""
        )
        + (
            "<h3 style='color:#B25900;'>Changed Fields (Mutable — CP Is Permitted To Change)</h3>"
            '<table style="border-collapse:collapse;width:100%;margin-bottom:16px;">'
            "<tr style='background:#f0f0f0;'>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Row</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Circuit ID</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Field</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Old Value</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>New Value</th>"
            "</tr>"
            f"{changed_rows_html}"
            "</table>"
            if changed_rows_html else ""
        )
        + (
            "<div style='background:#FFCCCC;border:2px solid #CC0000;padding:12px;margin:12px 0;border-radius:4px;'>"
            "<strong style='color:#CC0000;font-size:15px;'>⚠ EMAIL ADDRESS — FIX NEEDED</strong><br>"
            "<p style='margin:6px 0;'>The following circuits have invalid email addresses. "
            "The master has <strong>NOT</strong> been updated for these rows. "
            "The email address must be corrected and the form re-submitted before the master can be updated.</p>"
            "</div>"
            '<table style="border-collapse:collapse;width:100%;margin-bottom:16px;">'
            "<tr style='background:#f0f0f0;'>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Row</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Circuit ID</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Invalid Email</th>"
            "<th style='padding:4px 8px;border:1px solid #ccc;text-align:left'>Action Required</th>"
            "</tr>"
            f"{invalid_email_rows_html}"
            "</table>"
            if has_invalid_email_fix else ""
        )
        + '<p>The full data quality report is attached.</p>'
        '<hr style="margin-top:30px;border:none;border-top:1px solid #ccc;">'
        '<p style="font-size:11px;color:#888;">Automated data quality alert — PSTN Migration Adobe Sign Widget Monitor</p>'
        "</body></html>"
    )

    graph_attachments = []
    if report_path and Path(report_path).exists():
        data = Path(report_path).read_bytes()
        graph_attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": Path(report_path).name,
            "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "contentBytes": base64.b64encode(data).decode(),
        })

    fix_needed_prefix = "⚠ FIX NEEDED — INVALID EMAIL | " if has_invalid_email_fix else ""
    payload = {
        "message": {
            "subject": f"{fix_needed_prefix}⚠ DATA QUALITY ALERT — {reseller_name} — {agreement_id}",
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": recipients,
            "attachments": graph_attachments,
        }
    }

    endpoint = f"https://graph.microsoft.com/v1.0/users/{GRAPH_SENDER_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.post(endpoint, headers=headers, json=payload, timeout=60)
    if response.status_code == 202:
        logger.info(
            "✅  Data quality alert sent to %d recipient(s) for %s",
            len(recipients), reseller_name,
        )
    else:
        logger.error(
            "❌  Failed to send data quality alert for %s: HTTP %s — %s",
            reseller_name, response.status_code, response.text,
        )


def _send_pdf_review_email(
    token: str,
    agreement_id: str,
    reseller_name: str,
    validation_results: list[AttachmentValidationResult],
    report_path: Optional[str],
) -> None:
    """
    Send a PDF submission notification — lower urgency than the CTL data quality alert.

    Sent when a CP attachment arrived as PDF and could not be automatically validated
    against the circuit master. This is a process gap notification, not a data failure.
    Asks the data controller to contact the CP to re-submit as Excel.
    """
    recipients = _build_alert_recipients()
    if not recipients:
        logger.warning("No alert email recipients configured — PDF review email not sent.")
        return

    pdf_errors = []
    for v in validation_results:
        if v.pdf_human_review_required:
            for err in v.errors:
                if err:
                    pdf_errors.append(err)

    error_rows_html = "".join(
        f"<li style='margin:4px 0;'>{e}</li>" for e in pdf_errors
    )

    html_body = (
        '<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333;">'
        "<div style='background:#FFF3CD;border-left:6px solid #FFA500;padding:12px 16px;margin-bottom:16px;'>"
        "<strong style='color:#856404;font-size:16px;'>📎 PDF SUBMISSION — MANUAL REVIEW REQUIRED</strong><br>"
        f"<span>Reseller: <strong>{reseller_name}</strong> &nbsp;|&nbsp; Agreement: <strong>{agreement_id}</strong></span>"
        "</div>"
        "<p>A submission has been received from the above reseller that included a <strong>PDF attachment</strong>. "
        "The PDF could not be automatically validated against the circuit master.</p>"
        "<p><strong>This is not a data quality failure</strong> — no incorrect data has been detected. "
        "However, the attachment could not be processed automatically and requires manual review.</p>"
        "<h3 style='color:#856404;'>Processing Notes</h3>"
        f"<ul style='background:#f9f9f9;padding:12px 12px 12px 28px;border:1px solid #ddd;border-radius:4px;'>{error_rows_html or '<li>No detail available.</li>'}</ul>"
        "<h3>Action Required</h3>"
        "<ol>"
        "<li>Review the data quality report attached to this email.</li>"
        "<li>Contact the reseller and request they re-submit the circuit data as an <strong>Excel (.xlsx) file</strong> "
        "by email or via the webform.</li>"
        "<li>Once the Excel is received, validate it manually against the circuit master before processing.</li>"
        "</ol>"
        + (
            '<p style="color:#888;font-size:12px;">If a CSV was extracted from the PDF, it is included in the data quality report for reference only. '
            "It has NOT been validated against the master and must NOT be used in downstream processing without human review.</p>"
        )
        + '<hr style="margin-top:30px;border:none;border-top:1px solid #ccc;">'
        '<p style="font-size:11px;color:#888;">Automated PDF submission notification — PSTN Migration Adobe Sign Widget Monitor</p>'
        "</body></html>"
    )

    graph_attachments = []

    # Attach the data quality report xlsx
    if report_path and Path(report_path).exists():
        data = Path(report_path).read_bytes()
        graph_attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": Path(report_path).name,
            "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "contentBytes": base64.b64encode(data).decode(),
        })

    # Attach the source PDF(s) and any extracted CSV so the data controller
    # has the original document to open/forward without needing to find it on disk
    for v in validation_results:
        if not v.pdf_human_review_required:
            continue
        src = Path(v.attachment_path) if v.attachment_path else None
        if src and src.exists():
            graph_attachments.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": src.name,
                "contentType": "application/pdf",
                "contentBytes": base64.b64encode(src.read_bytes()).decode(),
            })
        # If pdfminer/OCR extracted the table, an XLSX will exist alongside the PDF
        if src:
            xlsx_path = src.with_suffix(".xlsx")
            if xlsx_path.exists():
                graph_attachments.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": xlsx_path.name,
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": base64.b64encode(xlsx_path.read_bytes()).decode(),
                })

    payload = {
        "message": {
            "subject": f"📎 PDF Submission — Manual Review Required — {reseller_name}",
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": recipients,
            "attachments": graph_attachments,
        }
    }

    endpoint = f"https://graph.microsoft.com/v1.0/users/{GRAPH_SENDER_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.post(endpoint, headers=headers, json=payload, timeout=60)
    if response.status_code == 202:
        logger.info(
            "✅  PDF review notification sent to %d recipient(s) for %s",
            len(recipients), reseller_name,
        )
    else:
        logger.error(
            "❌  Failed to send PDF review notification for %s: HTTP %s — %s",
            reseller_name, response.status_code, response.text,
        )


def _send_partial_opt_out_email(
    token: str,
    agreement_id: str,
    reseller_name: str,
    account_number: str,
    partial_circuits: list[str],
    total_rows: int,
    signed_pdf_path: Optional[str],
    attachment_paths: list,
) -> None:
    """
    Send a partial opt-out notification to PARTIAL_OPT_OUT_EMAIL.

    Triggered when one or more circuits in the attachment have
    'include in migration y/n' set to N/n/NO/no.
    """
    if not PARTIAL_OPT_OUT_EMAIL:
        logger.warning("PARTIAL_OPT_OUT_EMAIL not configured — partial opt-out email not sent.")
        return

    count = len(partial_circuits)
    pct = f"{(count / total_rows * 100):.1f}" if total_rows > 0 else "0"

    circuit_rows_html = "".join(
        f"<tr><td style='padding:4px 12px 4px 0;border-bottom:1px solid #eee'>{cid}</td></tr>"
        for cid in partial_circuits
    )

    html_body = (
        '<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333;">'
        "<div style='background:#FFF3CD;border-left:6px solid #FFA500;padding:12px 16px;margin-bottom:16px;'>"
        "<strong style='font-size:16px;'>Partial Opt-Out Received</strong><br>"
        f"<span>Reseller: <strong>{reseller_name}</strong> &nbsp;|&nbsp; Agreement: <strong>{agreement_id}</strong></span>"
        "</div>"
        "<p>A Communication Provider has submitted a mapping file containing circuits that "
        "are <strong>opted out of migration</strong> (include in migration = N/No).</p>"
        '<table style="border-collapse:collapse;margin:12px 0;">'
        f"<tr><td style='padding:4px 16px 4px 0;font-weight:bold'>Reseller Name</td><td>{reseller_name}</td></tr>"
        f"<tr><td style='padding:4px 16px 4px 0;font-weight:bold'>Gamma Account Number</td><td>{account_number}</td></tr>"
        f"<tr><td style='padding:4px 16px 4px 0;font-weight:bold'>Agreement ID</td><td>{agreement_id}</td></tr>"
        f"<tr><td style='padding:4px 16px 4px 0;font-weight:bold'>Opted-Out Circuits</td><td><strong>{count}</strong> of {total_rows} ({pct}%)</td></tr>"
        "</table>"
        "<h3 style='margin-bottom:6px;'>Opted-Out Circuit IDs</h3>"
        '<table style="border-collapse:collapse;">'
        f"{circuit_rows_html}"
        "</table>"
        "<p>The signed form and asset register are attached for reference.</p>"
        '<hr style="margin-top:30px;border:none;border-top:1px solid #ccc;">'
        '<p style="font-size:11px;color:#888;">Automated notification — PSTN Migration Adobe Sign Widget Monitor</p>'
        "</body></html>"
    )

    graph_attachments = []
    if signed_pdf_path and Path(signed_pdf_path).exists():
        data = Path(signed_pdf_path).read_bytes()
        graph_attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": Path(signed_pdf_path).name,
            "contentType": "application/pdf",
            "contentBytes": base64.b64encode(data).decode(),
        })
    for att_path in attachment_paths:
        p = Path(att_path)
        if not p.exists():
            continue
        mime_map = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".csv": "text/csv"}
        mime = mime_map.get(p.suffix, "application/octet-stream")
        graph_attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": p.name,
            "contentType": mime,
            "contentBytes": base64.b64encode(p.read_bytes()).decode(),
        })

    payload = {
        "message": {
            "subject": f"Partial Opt-Out Received — {reseller_name}",
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": _parse_email_list(PARTIAL_OPT_OUT_EMAIL),
            "attachments": graph_attachments,
        }
    }

    endpoint = f"https://graph.microsoft.com/v1.0/users/{GRAPH_SENDER_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.post(endpoint, headers=headers, json=payload, timeout=60)
    if response.status_code == 202:
        logger.info(
            "✅  Partial opt-out email sent to %s for %s (%d circuit(s))",
            PARTIAL_OPT_OUT_EMAIL, reseller_name, count,
        )
    else:
        logger.error(
            "❌  Failed to send partial opt-out email for %s: HTTP %s — %s",
            reseller_name, response.status_code, response.text,
        )


def _find_pa_original(agreement_id: str, doc_name: str) -> Optional[Path]:
    """
    Look for an original attachment file dropped by Power Automate.

    Power Automate flow should save files to:
        {PA_ATTACHMENT_DROP_FOLDER}\\{agreement_id}\\{original_filename}

    Returns the Path if a matching file is found, otherwise None.
    Falls back to scanning the agreement subfolder for any Excel file when
    the exact filename is not present (handles PA renaming edge cases).
    """
    if not PA_ATTACHMENT_DROP_FOLDER:
        return None

    agreement_dir = Path(PA_ATTACHMENT_DROP_FOLDER) / agreement_id
    if not agreement_dir.is_dir():
        return None

    # Exact filename match first
    if doc_name:
        exact = agreement_dir / doc_name
        if exact.exists():
            return exact

    # Fallback: any Excel file in the agreement subfolder
    excel_exts = {".xlsx", ".xls", ".xlsm"}
    candidates = [f for f in agreement_dir.iterdir() if f.suffix.lower() in excel_exts]
    if candidates:
        # If more than one, prefer the one closest to the doc_name
        if len(candidates) == 1:
            return candidates[0]
        # Multiple — try stem match
        stem = Path(doc_name).stem.lower() if doc_name else ""
        for c in candidates:
            if c.stem.lower() == stem:
                return c
        # Give up on guessing — don't silently use the wrong file
        logger.warning(
            "PA drop folder %s has %d Excel files but none match '%s' — "
            "falling back to Adobe Sign download.",
            agreement_dir, len(candidates), doc_name,
        )
        return None

    return None


def process_signed_agreement(agreement_id: str, partner_name: str) -> dict:
    """
    Extract and save all data from a completed agreement.

    Args:
        agreement_id: Adobe Sign agreement ID
        partner_name:  Human-readable name (used for output filenames)

    Returns:
        Summary dict describing what was saved.
    """
    logger.info("Processing completed agreement: %s  (%s)", partner_name, agreement_id)

    client = AdobeSignClient()
    safe_name = _safe_filename(partner_name)
    ts = _timestamp()
    output: dict = {
        "agreement_id": agreement_id,
        "partner_name": partner_name,
        "reseller_name": "",
        "account_number": "",
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "form_data_file": None,
        "attachment_files": [],
        "attachment_ids": [],
        "signed_pdf_file": None,
        "opt_out_full": False,
        "rwp_remain": False,
        "confirm_amendments": False,
        "validation_results": [],
        "is_partial_opt_out": False,
        "has_data_quality_issues": False,
        "data_quality_report_file": None,
        "errors": [],
    }

    # 1. Form field data ───────────────────────────────────────────────────────
    try:
        form_rows = client.get_form_data(agreement_id)
        if form_rows:
            form_json_path = OUTPUT_DIR / f"{safe_name}_form_data_{ts}.json"
            with open(form_json_path, "w", encoding="utf-8") as fh:
                json.dump(form_rows, fh, indent=2, ensure_ascii=False)
            output["form_data_file"] = str(form_json_path)
            logger.info("Form data saved → %s", form_json_path.name)

            # Extract and log checkbox selections
            if form_rows:
                row = form_rows[-1]  # last row = the signer's submission
                output["opt_out_full"]       = _is_ticked(row.get(FIELD_OPT_OUT_FULL, ""))
                output["rwp_remain"]         = _is_ticked(row.get(FIELD_CONFIRM_RWP_REMAIN, ""))
                output["confirm_amendments"] = _is_ticked(row.get(FIELD_CONFIRM_AMENDMENTS, ""))
                output["reseller_name"]      = row.get(FIELD_RESELLER_NAME, "").strip()
                output["account_number"]     = row.get(FIELD_ACCOUNT_NUMBER, "").strip()
                logger.info(
                    "  %-30s: %s", FIELD_OPT_OUT_FULL,
                    "TICKED" if output["opt_out_full"] else "not ticked",
                )
                logger.info(
                    "  %-30s: %s", FIELD_CONFIRM_RWP_REMAIN,
                    "TICKED" if output["rwp_remain"] else "not ticked",
                )
                logger.info(
                    "  %-30s: %s", FIELD_CONFIRM_AMENDMENTS,
                    "TICKED" if output["confirm_amendments"] else "not ticked",
                )
        else:
            logger.warning("No form data returned for %s", agreement_id)
    except AdobeSignError as exc:
        msg = f"Could not retrieve form data: {exc}"
        logger.error(msg)
        output["errors"].append(msg)

    # 2. Signer-uploaded attachments ──────────────────────────────────────────
    # Primary path: combinedDocument?attachSupportingDocuments=true returns a
    # ZIP containing the original uploaded files (real .xlsx, not Adobe's
    # PDF-converted wrapper). If this succeeds and yields files, we use those
    # and skip the per-document download entirely.
    # Fallback: if the ZIP is empty (no supporting docs) or the call fails,
    # fall through to the standard get_documents() per-document loop which
    # handles the PA drop folder (Option A) and Adobe Sign download (Option B).

    _zip_attachments: dict[str, bytes] = {}
    try:
        _zip_attachments = client.download_original_attachments_zip(agreement_id)
        if _zip_attachments:
            logger.info(
                "ZIP attachment package: %d original file(s) retrieved for %s "
                "(original format preserved — bypassing Adobe PDF conversion).",
                len(_zip_attachments), agreement_id,
            )
    except AdobeSignError as zip_exc:
        logger.warning(
            "ZIP attachment download failed for %s (%s) — falling back to "
            "per-document download.", agreement_id, zip_exc,
        )

    if _zip_attachments:
        # ── ZIP path: original files ───────────────────────────────────────
        for orig_filename, raw_bytes in _zip_attachments.items():
            orig_path = Path(orig_filename)
            orig_ext  = orig_path.suffix.lower()
            safe_stem = re.sub(r"[^\w]", "_", orig_path.stem).lower()

            att_path = OUTPUT_DIR / f"{safe_name}_{safe_stem}_{ts}{orig_ext}"
            att_path.write_bytes(raw_bytes)
            output["attachment_files"].append(str(att_path))
            output["attachment_ids"].append("zip_source")
            logger.info(
                "Attachment saved (original) → %s  (%d bytes)",
                att_path.name, len(raw_bytes),
            )

            try:
                if att_path.suffix.lower() in (".xlsx", ".xls", ".xlsm", ".pdf") and CIRCUIT_MASTER_PATH:
                    val_result = validate_attachment(
                        att_path=att_path,
                        agreement_id=agreement_id,
                        reseller_name=output["reseller_name"] or partner_name,
                        convert_to_csv=CONVERT_ATTACHMENT_TO_CSV,
                    )
                    output["validation_results"].append(val_result)
                    xlsx_path = att_path.with_suffix(".xlsx")
                    if xlsx_path.exists() and str(xlsx_path) not in output["attachment_files"]:
                        output["attachment_files"].append(str(xlsx_path))
                elif att_path.suffix.lower() in (".xlsx", ".xls", ".xlsm") and CONVERT_ATTACHMENT_TO_CSV:
                    df = pd.read_excel(att_path, dtype=str)
                    csv_path = OUTPUT_DIR / f"{safe_name}_{safe_stem}_{ts}.csv"
                    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                    output["attachment_files"].append(str(csv_path))
                    logger.info(
                        "Attachment parsed → %s  (%d rows, %d cols)",
                        csv_path.name, len(df), len(df.columns),
                    )
            except Exception as val_exc:
                msg = f"Error processing ZIP attachment '{orig_filename}': {val_exc}"
                logger.error(msg)
                output["errors"].append(msg)

    else:
        # ── Fallback: per-document download ────────────────────────────────
        try:
            doc_data = client.get_documents(agreement_id)
            supporting = doc_data.get("supportingDocuments", [])

            if not supporting:
                logger.info("No supporting attachments found for %s", agreement_id)
            else:
                logger.info("Found %d supporting attachment(s).", len(supporting))

            for doc in supporting:
                doc_id = doc.get("id", "")
                field_name = doc.get("fieldName", "attachment")
                mime = doc.get("mimeType", "").lower()
                original_name = doc.get("name", "")
                safe_field = re.sub(r"[^\w]", "_", field_name).lower()

                orig_ext = Path(original_name).suffix.lower() if original_name else ""

                # ── Option A: Power Automate dropped the original file locally ──
                pa_source = _find_pa_original(agreement_id, original_name)
                if pa_source:
                    ext = pa_source.suffix.lower()
                    att_path = OUTPUT_DIR / f"{safe_name}_{safe_field}_{ts}{ext}"
                    shutil.copy2(str(pa_source), str(att_path))
                    output["attachment_files"].append(str(att_path))
                    output["attachment_ids"].append(doc_id)
                    logger.info(
                        "Attachment '%s' sourced from PA drop folder (original format preserved) → %s",
                        original_name, att_path.name,
                    )
                else:
                    # ── Option B: Download from Adobe Sign (may be PDF-converted) ─
                    adobe_converted = False
                    if "excel" in mime or "spreadsheetml" in mime or orig_ext in (".xlsx", ".xls", ".xlsm"):
                        if "pdf" in mime and orig_ext in (".xlsx", ".xls", ".xlsm"):
                            adobe_converted = True
                            ext = ".pdf"
                            logger.info(
                                "Attachment '%s' was originally Excel (%s) but Adobe Sign returned "
                                "mimeType '%s' — auto-converted to PDF during signing. "
                                "Will attempt table extraction for validation. "
                                "Consider configuring PA_ATTACHMENT_DROP_FOLDER for cleaner results.",
                                original_name, orig_ext, mime,
                            )
                        else:
                            ext = ".xlsx"
                    elif "pdf" in mime:
                        ext = ".pdf"
                    else:
                        ext = orig_ext or ""
                        if ext:
                            logger.info(
                                "Attachment '%s' has unrecognised mimeType '%s' — using original extension '%s'.",
                                original_name, mime, ext,
                            )

                    try:
                        raw_bytes = client.download_document(agreement_id, doc_id)
                        att_path = OUTPUT_DIR / f"{safe_name}_{safe_field}_{ts}{ext}"
                        att_path.write_bytes(raw_bytes)
                        output["attachment_files"].append(str(att_path))
                        output["attachment_ids"].append(doc_id)
                        logger.info("Attachment saved → %s  (%d bytes)", att_path.name, len(raw_bytes))
                    except Exception as doc_exc:
                        msg = f"Failed to download attachment '{field_name}': {doc_exc}"
                        logger.error(msg)
                        output["errors"].append(msg)
                        continue

                try:
                    if att_path.suffix.lower() in (".xlsx", ".pdf") and CIRCUIT_MASTER_PATH:
                        try:
                            val_result = validate_attachment(
                                att_path=att_path,
                                agreement_id=agreement_id,
                                reseller_name=output["reseller_name"] or partner_name,
                                convert_to_csv=CONVERT_ATTACHMENT_TO_CSV,
                            )
                            output["validation_results"].append(val_result)
                            # Add PDF-extracted XLSX to attachment list so it is included
                            # in the partial opt-out email for onward processing
                            xlsx_path = att_path.with_suffix(".xlsx")
                            if xlsx_path.exists() and str(xlsx_path) not in output["attachment_files"]:
                                output["attachment_files"].append(str(xlsx_path))
                            csv_path = att_path.with_suffix(".csv")
                            if csv_path.exists() and str(csv_path) not in output["attachment_files"]:
                                output["attachment_files"].append(str(csv_path))
                        except Exception as val_exc:
                            msg = f"Validation failed for attachment '{field_name}': {val_exc}"
                            logger.error(msg)
                            output["errors"].append(msg)

                    elif att_path.suffix.lower() == ".xlsx" and CONVERT_ATTACHMENT_TO_CSV:
                        df = pd.read_excel(att_path, dtype=str)
                        csv_path = OUTPUT_DIR / f"{safe_name}_{safe_field}_{ts}.csv"
                        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                        output["attachment_files"].append(str(csv_path))
                        logger.info(
                            "Attachment parsed → %s  (%d rows, %d cols)",
                            csv_path.name, len(df), len(df.columns),
                        )

                except Exception as doc_exc:
                    msg = f"Error processing attachment '{field_name}': {doc_exc}"
                    logger.error(msg)
                    output["errors"].append(msg)

        except AdobeSignError as exc:
            msg = f"Could not retrieve attachments: {exc}"
            logger.error(msg)
            output["errors"].append(msg)

    # 3. Signed PDF ────────────────────────────────────────────────────────────
    try:
        pdf_bytes = client.download_combined_document(agreement_id)
        pdf_path = OUTPUT_DIR / f"{safe_name}_signed_{ts}.pdf"
        pdf_path.write_bytes(pdf_bytes)
        output["signed_pdf_file"] = str(pdf_path)
        logger.info("Signed PDF saved → %s", pdf_path.name)
    except AdobeSignError as exc:
        msg = f"Could not download signed PDF: {exc}"
        logger.error(msg)
        output["errors"].append(msg)

    # 3b. Post-attachment: aggregate validation results and trigger alert emails ──────
    if output["validation_results"]:
        # Actual data quality issues: row-level CTL failures or Excel missing required columns.
        # Excludes PDF submissions that simply couldn't be parsed — those are routed to a
        # separate lower-urgency notification so the CTL alert is reserved for real failures.
        # Data quality issues = problems with IMMUTABLE columns only (grandparentname,
        # grandparent_id, circuit_id, address, product_name, line to migrate to).
        # Changes to MUTABLE columns (include in migration y/n, reason, email address)
        # are expected CP updates — they do NOT trigger a data alert.
        any_critical = any(
            v.critical_count > 0 or (bool(v.missing_columns) and not v.pdf_human_review_required)
            for v in output["validation_results"]
        )
        any_partial  = any(v.is_partial_opt_out for v in output["validation_results"])
        # PDF submissions that couldn't be auto-validated (process gap, not data failure)
        any_pdf_review = any(v.pdf_human_review_required for v in output["validation_results"])

        output["has_data_quality_issues"] = any_critical
        output["is_partial_opt_out"] = any_partial

        # Use the last report as the primary data quality report
        last_report = next(
            (v.data_quality_report_path for v in reversed(output["validation_results"]) if v.data_quality_report_path),
            None,
        )
        output["data_quality_report_file"] = last_report

        try:
            graph_token = _get_graph_token()

            # Send data quality email for ALL processed attachments (shows mutable changes + critical issues)
            # Only skip when PDF couldn't be parsed and needs separate human review notification
            if any_pdf_review:
                # PDF couldn't be auto-validated — lower-urgency process notification
                _send_pdf_review_email(
                    token=graph_token,
                    agreement_id=agreement_id,
                    reseller_name=output["reseller_name"] or partner_name,
                    validation_results=output["validation_results"],
                    report_path=last_report,
                )
            else:
                # Send data quality email for all successfully processed attachments
                # Shows mutable field changes (yellow) and critical issues (red) if any
                _send_data_alert_email(
                    token=graph_token,
                    agreement_id=agreement_id,
                    reseller_name=output["reseller_name"] or partner_name,
                    validation_results=output["validation_results"],
                    report_path=last_report,
                )

            # Also treat as partial when neither opt-out nor opt-in checkbox is ticked
            # (partner submitted circuit-level attachment for PS review without a clear full decision)
            is_partial_submission = not output["opt_out_full"] and not output["rwp_remain"]
            if any_partial or is_partial_submission:
                partial_circuits = [
                    r.circuit_id
                    for v in output["validation_results"]
                    for r in v.row_results
                    if r.is_partial_opt_out
                ]
                total_rows = sum(v.total_rows for v in output["validation_results"])
                _send_partial_opt_out_email(
                    token=graph_token,
                    agreement_id=agreement_id,
                    reseller_name=output["reseller_name"] or partner_name,
                    account_number=output["account_number"],
                    partial_circuits=partial_circuits,
                    total_rows=total_rows,
                    signed_pdf_path=output.get("signed_pdf_file"),
                    attachment_paths=[p for p in output["attachment_files"] if Path(p).exists()],
                )
        except Exception as email_exc:
            msg = f"Could not send data quality alert emails: {email_exc}"
            logger.error(msg)
            output["errors"].append(msg)

    # 4. Forward opt-out submission ────────────────────────────────────────────────
    if output["opt_out_full"] and OPT_OUT_FORWARD_EMAIL:
        try:
            graph_token = _get_graph_token()
            all_att_paths = [p for p in output["attachment_files"] if Path(p).exists()]
            _forward_opt_out(
                token=graph_token,
                reseller_name=output["reseller_name"] or partner_name,
                account_number=output["account_number"],
                agreement_id=agreement_id,
                signed_pdf_path=output["signed_pdf_file"],
                attachment_paths=all_att_paths,
            )
        except Exception as fwd_exc:
            msg = f"Could not forward opt-out email: {fwd_exc}"
            logger.error(msg)
            output["errors"].append(msg)

    # 4b. Full opt-out — update master sheet for all partner circuits ──────────────
    if output["opt_out_full"]:
        try:
            rows_queued, fo_errors = queue_full_opt_out_master_updates(
                reseller_name=output["reseller_name"] or partner_name,
                account_number=output["account_number"],
                agreement_id=agreement_id,
            )
            output["errors"].extend(fo_errors)
            if rows_queued:
                logger.info(
                    "✅  Full opt-out master update queued: %d row(s) will be set to N / 'Full Opt-Out'",
                    rows_queued,
                )
        except Exception as fo_exc:
            msg = f"Could not queue full opt-out master updates: {fo_exc}"
            logger.error(msg)
            output["errors"].append(msg)

    # 5. Write date_received back to Excel ───────────────────────────────────────
    try_update_excel_after_completion(agreement_id)

    # 6. Mark processed in state ──────────────────────────────────────────
    state = _load_state()
    for record in state:
        if record.get("agreement_id") == agreement_id:
            record["processed"]          = True
            record["opt_out_full"]       = output["opt_out_full"]
            record["rwp_remain"]         = output["rwp_remain"]
            record["confirm_amendments"] = output["confirm_amendments"]
            break
    _save_state(state)

    if output["errors"]:
        logger.warning(
            "Processing complete with %d error(s) for %s", len(output["errors"]), partner_name
        )
    else:
        logger.info("✅  Processing complete for %s", partner_name)

    return output


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Process a single signed agreement.")
    parser.add_argument("agreement_id", help="Adobe Sign agreement ID")
    parser.add_argument("partner_name", help="Partner name (used for output filenames)")
    args = parser.parse_args()
    result = process_signed_agreement(args.agreement_id, args.partner_name)
    print(json.dumps(result, indent=2))
