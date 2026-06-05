"""
git_backup.py — Pre-write snapshot backup of the circuit master sheet to GitHub.

Commits the current state of the master Excel file to a GitHub repository
before any update is applied. This provides a recoverable audit trail of
what the master contained before the automation modified it.

Usage is optional — if GITHUB_TOKEN / GITHUB_REPO / GITHUB_MASTER_BACKUP_PATH
are not configured in .env, the backup is silently skipped.

This module uses the GitHub Contents API (no git CLI required).
"""
from __future__ import annotations

import base64
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import requests

from config import GITHUB_MASTER_BACKUP_PATH, GITHUB_REPO, GITHUB_TOKEN

logger = logging.getLogger(__name__)

_GITHUB_API = "https://api.github.com"


def backup_master_to_github(local_path: Path) -> Optional[str]:
    """
    Upload the current state of local_path to GitHub as a pre-run backup.

    Steps:
      1. GET current file SHA (required by GitHub API for update vs create)
      2. Base64-encode the local file content
      3. PUT to GitHub with commit message including ISO timestamp

    Returns the commit SHA string on success, None if skipped or failed.
    Failures are logged as warnings — this function never raises.
    """
    if not GITHUB_TOKEN or not GITHUB_REPO or not GITHUB_MASTER_BACKUP_PATH:
        logger.debug(
            "GitHub backup skipped — GITHUB_TOKEN, GITHUB_REPO, or GITHUB_MASTER_BACKUP_PATH "
            "not configured in .env."
        )
        return None

    if not local_path.exists():
        logger.warning(
            "GitHub backup skipped — local master file not found: %s", local_path
        )
        return None

    ts_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ts_label = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    api_url = f"{_GITHUB_API}/repos/{GITHUB_REPO}/contents/{GITHUB_MASTER_BACKUP_PATH}"

    # Step 1: Get current file SHA (needed for update; None means file doesn't exist yet)
    current_sha: Optional[str] = None
    try:
        get_resp = requests.get(api_url, headers=headers, timeout=30)
        if get_resp.status_code == 200:
            current_sha = get_resp.json().get("sha")
            logger.debug("GitHub existing file SHA: %s", current_sha)
        elif get_resp.status_code == 404:
            logger.info("GitHub file does not exist yet — will create: %s", GITHUB_MASTER_BACKUP_PATH)
        else:
            logger.warning(
                "GitHub backup: unexpected status checking existing file: HTTP %s — %s",
                get_resp.status_code, get_resp.text[:200],
            )
    except requests.RequestException as exc:
        logger.warning("GitHub backup: failed to check existing file SHA: %s", exc)
        return None

    # Step 2: Read and encode local file
    try:
        content_bytes = local_path.read_bytes()
        content_b64 = base64.b64encode(content_bytes).decode("ascii")
    except Exception as exc:
        logger.warning("GitHub backup: failed to read local master file: %s", exc)
        return None

    # Step 3: PUT to GitHub
    payload: dict = {
        "message": f"pre-run backup {ts_iso} — automated PSTN widget monitor",
        "content": content_b64,
        "branch": "main",
    }
    if current_sha:
        payload["sha"] = current_sha

    try:
        put_resp = requests.put(api_url, headers=headers, json=payload, timeout=60)
        if put_resp.status_code in (200, 201):
            commit_sha = put_resp.json().get("commit", {}).get("sha", "unknown")
            logger.info(
                "✅  GitHub backup committed: %s (commit SHA: %s)",
                GITHUB_MASTER_BACKUP_PATH, commit_sha,
            )
            return commit_sha
        else:
            logger.warning(
                "GitHub backup failed: HTTP %s — %s",
                put_resp.status_code, put_resp.text[:400],
            )
            return None
    except requests.RequestException as exc:
        logger.warning("GitHub backup: PUT request failed: %s", exc)
        return None
