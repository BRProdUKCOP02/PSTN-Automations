"""
adobe_sign_client.py — REST API v6 wrapper for Adobe Acrobat Sign.

Authentication: OAuth 2.0 refresh token flow.
The refresh token is exchanged for a short-lived access token on first use
and refreshed automatically when it expires.

All methods raise AdobeSignError on non-2xx responses.
Rate-limit (HTTP 429) responses are retried automatically with the
server-specified Retry-After delay.
"""
import csv
import io
import logging
import time
import zipfile
from typing import Optional
from urllib.parse import quote

import requests

from config import (
    ADOBE_SIGN_BASE_URL,
    ADOBE_SIGN_CLIENT_ID,
    ADOBE_SIGN_CLIENT_SECRET,
    ADOBE_SIGN_REFRESH_TOKEN,
)

logger = logging.getLogger(__name__)

# Adobe Sign uses /oauth/v2/refresh (not /token) for refresh token grants
def _token_url() -> str:
    return f"{ADOBE_SIGN_BASE_URL}/oauth/v2/refresh"


class AdobeSignError(Exception):
    """Raised when the Adobe Sign API returns an error response."""

    def __init__(self, status_code: int, body: str, url: str = ""):
        self.status_code = status_code
        self.body = body
        self.url = url
        super().__init__(f"Adobe Sign API error {status_code} at {url}: {body}")


class AdobeSignClient:
    """Thin wrapper around the Adobe Sign REST API v6 using OAuth 2.0."""

    _MAX_RETRIES = 3

    def __init__(self):
        self._base_url = ADOBE_SIGN_BASE_URL
        self._access_token: Optional[str] = None
        self._token_expiry: float = 0.0

    def _get_access_token(self) -> str:
        """Refresh the OAuth access token using the /oauth/v2/refresh endpoint."""
        if self._access_token and time.time() < self._token_expiry - 60:
            return self._access_token
        payload = {
            "grant_type": "refresh_token",
            "client_id": ADOBE_SIGN_CLIENT_ID,
            "client_secret": ADOBE_SIGN_CLIENT_SECRET,
            "refresh_token": ADOBE_SIGN_REFRESH_TOKEN,
        }
        url = _token_url()
        resp = requests.post(url, data=payload, timeout=30)
        if not resp.ok:
            raise AdobeSignError(resp.status_code, resp.text, url)
        data = resp.json()
        self._access_token = data["access_token"]
        self._token_expiry = time.time() + data.get("expires_in", 3600)
        return self._access_token

    def _auth_headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Content-Type": "application/json",
        }

    # ── Private helpers ───────────────────────────────────────────────────────

    def _request(
        self,
        method: str,
        path: str,
        *,
        json: Optional[dict] = None,
        params: Optional[dict] = None,
        stream: bool = False,
    ) -> requests.Response:
        """Execute a request and handle 429 retry-after automatically."""
        url = f"{self._base_url}/api/rest/v6{path}"
        for attempt in range(1, self._MAX_RETRIES + 1):
            response = requests.request(
                method,
                url,
                headers=self._auth_headers(),
                json=json,
                params=params,
                stream=stream,
                timeout=60,
            )
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", 60))
                logger.warning(
                    "Rate limited (429). Waiting %s seconds before retry %s/%s.",
                    retry_after,
                    attempt,
                    self._MAX_RETRIES,
                )
                time.sleep(retry_after)
                continue
            if not response.ok:
                raise AdobeSignError(response.status_code, response.text, url)
            return response
        raise AdobeSignError(429, "Max retries exceeded after repeated 429 responses.", url)

    def _multipart_request(self, method: str, path: str, files: dict, data: Optional[dict] = None) -> requests.Response:
        """Execute a multipart/form-data request (used for file uploads)."""
        url = f"{self._base_url}/api/rest/v6{path}"
        # Strip Content-Type so requests sets the correct multipart boundary
        headers = {k: v for k, v in self._auth_headers().items() if k != "Content-Type"}
        for attempt in range(1, self._MAX_RETRIES + 1):
            response = requests.request(
                method, url, headers=headers, files=files, data=data, timeout=120
            )
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", 60))
                logger.warning("Rate limited (429). Waiting %s seconds.", retry_after)
                time.sleep(retry_after)
                continue
            if not response.ok:
                raise AdobeSignError(response.status_code, response.text, url)
            return response
        raise AdobeSignError(429, "Max retries exceeded after repeated 429 responses.", url)

    # ── Library documents ─────────────────────────────────────────────────────

    def list_library_documents(self) -> list[dict]:
        """Return all library documents visible to the integration key user."""
        response = self._request("GET", "/libraryDocuments")
        return response.json().get("libraryDocumentList", [])

    # ── Widgets ───────────────────────────────────────────────────────────────

    def list_widgets(self) -> list[dict]:
        """Return all widgets (webforms) visible to the integration key user."""
        response = self._request("GET", "/widgets")
        return response.json().get("userWidgetList", [])

    def get_widget(self, widget_id: str) -> dict:
        """Return full widget details including name, status, and webform URL."""
        encoded_id = quote(widget_id, safe="")
        response = self._request("GET", f"/widgets/{encoded_id}")
        return response.json()

    # ── Agreements ────────────────────────────────────────────────────────────

    def create_agreement(
        self,
        library_document_id: str,
        recipient_email: str,
        recipient_name: str,
        agreement_name: str,
        merge_fields: Optional[list[dict]] = None,
        message: str = "",
    ) -> str:
        """
        Create and send an agreement from a Library Document.

        merge_fields format:
            [{"fieldName": "partner_name", "defaultValue": "Acme Ltd"}, ...]

        Returns the agreementId string.
        """
        payload = {
            "fileInfos": [
                {"libraryDocumentId": library_document_id}
            ],
            "name": agreement_name,
            "participantSetsInfo": [
                {
                    "memberInfos": [
                        {"email": recipient_email, "name": recipient_name}
                    ],
                    "order": 1,
                    "role": "SIGNER",
                }
            ],
            "signatureType": "ESIGN",
            "state": "IN_PROCESS",
            "message": message,
        }
        if merge_fields:
            payload["mergeFieldInfo"] = merge_fields

        response = self._request("POST", "/agreements", json=payload)
        agreement_id = response.json().get("id")
        logger.info("Agreement created: %s → %s", agreement_id, recipient_email)
        return agreement_id

    def get_agreement(self, agreement_id: str) -> dict:
        """Return full agreement info including current status."""
        response = self._request("GET", f"/agreements/{agreement_id}")
        return response.json()

    def list_agreements(self, status_filter: Optional[str] = None) -> list[dict]:
        """
        Return agreements. Optionally filter by status string, e.g. 'OUT_FOR_SIGNATURE'.
        Adobe Sign does not support server-side status filter on GET /agreements,
        so filtering is applied client-side.
        """
        response = self._request("GET", "/agreements")
        all_agreements = response.json().get("userAgreementList", [])
        if status_filter:
            return [a for a in all_agreements if a.get("status") == status_filter]
        return all_agreements

    def list_widget_agreements(self, widget_id: str) -> list[dict]:
        """
        Return all agreements submitted via a specific widget (webform).
        Each entry represents one CP submission.
        The widget ID is URL-encoded to handle special characters (e.g. trailing *).
        Handles pagination to retrieve all agreements.
        """
        encoded_id = quote(widget_id, safe="")
        all_agreements = []
        cursor = None
        page_num = 1
        
        while True:
            params = {"cursor": cursor} if cursor else {}
            response = self._request("GET", f"/widgets/{encoded_id}/agreements", params=params)
            data = response.json()
            
            agreements = data.get("userAgreementList", [])
            all_agreements.extend(agreements)
            logger.info("Page %d: fetched %d agreements", page_num, len(agreements))
            
            # Check for next page
            page_info = data.get("page", {})
            cursor = page_info.get("nextCursor")
            
            if cursor:
                logger.info("Found nextCursor, fetching page %d...", page_num + 1)
            
            # If no next cursor, we've retrieved all pages
            if not cursor:
                break
            
            page_num += 1
        
        logger.info("Total agreements fetched across %d page(s): %d", page_num, len(all_agreements))
        return all_agreements

    def get_signing_url(self, agreement_id: str) -> Optional[str]:
        """
        Return the first signer's signing URL for embedding in chaser emails.
        Returns None if the agreement is no longer awaiting signature.
        """
        try:
            response = self._request("GET", f"/agreements/{agreement_id}/signingUrls")
            sets = response.json().get("signingUrlSetInfos", [])
            if sets:
                urls = sets[0].get("signingUrls", [])
                if urls:
                    return urls[0].get("esignUrl")
        except AdobeSignError as exc:
            logger.debug("Could not retrieve signing URL for %s: %s", agreement_id, exc)
        return None

    # ── Form data & attachments ───────────────────────────────────────────────

    def get_form_data(self, agreement_id: str) -> list[dict]:
        """
        Download the form field data for a completed agreement.
        Returns a list of dicts (one row per signer session).
        Adobe Sign returns CSV — this parses it into a list of dicts.
        """
        response = self._request("GET", f"/agreements/{agreement_id}/formData", stream=True)
        content = response.content.decode("utf-8-sig")  # strip BOM if present
        reader = csv.DictReader(io.StringIO(content))
        return list(reader)

    def get_documents(self, agreement_id: str) -> dict:
        """
        Return all document metadata for the agreement, including
        'documents' (the agreement PDF) and 'supportingDocuments'
        (files uploaded by the signer).
        """
        response = self._request(
            "GET",
            f"/agreements/{agreement_id}/documents",
        )
        return response.json()

    def get_supporting_documents(self, agreement_id: str) -> list[dict]:
        """
        Return only the signer-uploaded file attachments (supporting documents).
        Uses the /documents endpoint with attachments=true which returns
        participant-uploaded files separately from the main agreement document.
        """
        response = self._request(
            "GET",
            f"/agreements/{agreement_id}/documents",
            params={"attachments": "true"},
        )
        data = response.json()
        # Supporting docs come back under 'supportingDocuments' key
        return data.get("supportingDocuments", [])

    def download_document(self, agreement_id: str, document_id: str) -> bytes:
        """Download and return raw bytes for a specific document."""
        response = self._request(
            "GET",
            f"/agreements/{agreement_id}/documents/{document_id}",
            stream=True,
        )
        return response.content

    def download_combined_document(self, agreement_id: str) -> bytes:
        """Download the combined signed PDF for the agreement."""
        response = self._request(
            "GET",
            f"/agreements/{agreement_id}/combinedDocument",
            stream=True,
        )
        return response.content

    def download_original_attachments_zip(
        self, agreement_id: str
    ) -> dict[str, bytes]:
        """
        Download supporting documents in their original file format.

        Calls GET /agreements/{id}/combinedDocument?attachSupportingDocuments=true
        which Adobe Sign returns as a ZIP package when supporting documents are
        present. The ZIP contains the original uploaded files (e.g. the actual
        .xlsx, not Adobe's PDF-converted wrapper).

        Returns a dict of {filename: raw_bytes} for every non-PDF file in the
        ZIP (i.e. the CP-uploaded attachments only — the signed PDF itself is
        excluded).

        Returns an empty dict when:
          - No supporting documents were uploaded (Adobe returns a plain PDF)
          - The response is not a ZIP for any other reason

        Raises AdobeSignError on non-2xx HTTP responses.
        """
        response = self._request(
            "GET",
            f"/agreements/{agreement_id}/combinedDocument",
            params={"attachSupportingDocuments": "true", "auditReport": "false"},
            stream=True,
        )
        content = response.content

        # Adobe returns a plain PDF when there are no supporting documents.
        # is_zipfile is the most reliable check — Content-Type alone is not.
        if not zipfile.is_zipfile(io.BytesIO(content)):
            return {}

        result: dict[str, bytes] = {}
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            for info in zf.infolist():
                name = info.filename
                if name.endswith("/"):          # skip directory entries
                    continue
                ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                if ext in ("xlsx", "xls", "xlsm", "csv"):
                    result[name] = zf.read(name)
        return result
