"""Microsoft Graph client: client-secret app-only auth, paging, throttling retries, and $batch."""
from __future__ import annotations

import logging
import time
from typing import Any, Iterator, Optional

import msal
import requests

from config import settings

logger = logging.getLogger("intune_report.graph_client")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA_BASE = "https://graph.microsoft.com/beta"
_AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant}"
_MAX_RETRIES = 5
_SCOPE = ["https://graph.microsoft.com/.default"]


class GraphClient:
    """Thin wrapper over Microsoft Graph REST calls used by the Intune report pipeline."""

    def __init__(self) -> None:
        self._app = msal.ConfidentialClientApplication(
            client_id=settings.client_id,
            authority=_AUTHORITY_TEMPLATE.format(tenant=settings.tenant_id),
            client_credential=settings.client_secret,
        )
        self._session = requests.Session()

    def _token(self) -> str:
        result = self._app.acquire_token_silent(scopes=_SCOPE, account=None)
        if not result:
            result = self._app.acquire_token_for_client(scopes=_SCOPE)
        if "access_token" not in result:
            raise RuntimeError(f"Failed to acquire Graph token: {result.get('error_description')}")
        return result["access_token"]

    def _request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        headers = {"Authorization": f"Bearer {self._token()}", "Content-Type": "application/json"}
        headers.update(kwargs.pop("headers", None) or {})
        response = None
        for attempt in range(1, _MAX_RETRIES + 1):
            response = self._session.request(method, url, headers=headers, **kwargs)
            if response.status_code in (429, 503, 504):
                retry_after = int(response.headers.get("Retry-After", 2 ** attempt))
                logger.warning(
                    "Graph throttled (%s) on %s %s, retrying in %ss (attempt %s)",
                    response.status_code, method, url, retry_after, attempt,
                )
                time.sleep(retry_after)
                continue
            response.raise_for_status()
            return response
        assert response is not None
        response.raise_for_status()
        return response

    def get_paged(self, url: str, params: Optional[dict[str, Any]] = None) -> Iterator[dict[str, Any]]:
        """Yield every item across all pages of a Graph collection endpoint."""
        next_url, next_params = url, params
        while next_url:
            response = self._request("GET", next_url, params=next_params)
            payload = response.json()
            yield from payload.get("value", [])
            next_url = payload.get("@odata.nextLink")
            next_params = None  # nextLink already contains the full query string

    def get(self, url: str, params: Optional[dict[str, Any]] = None) -> dict[str, Any]:
        return self._request("GET", url, params=params).json()

    def batch(self, requests_: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Execute write operations individually so SharePoint throttling receives HTTP retries."""
        responses: list[dict[str, Any]] = []
        for request in requests_:
            kwargs: dict[str, Any] = {}
            if "body" in request:
                kwargs["json"] = request["body"]
            if "headers" in request:
                kwargs["headers"] = request["headers"]
            try:
                response = self._request(
                    request["method"], f"{GRAPH_BASE}{request['url']}", **kwargs
                )
                body = response.json() if response.content else None
                responses.append({"id": request["id"], "status": response.status_code, "body": body})
            except requests.HTTPError as exc:
                error_response = exc.response
                body = error_response.json() if error_response.content else None
                responses.append({
                    "id": request["id"],
                    "status": error_response.status_code,
                    "body": body,
                })
        return responses

    def upload_file_to_list_drive(
        self, site_id: str, list_id: str, file_name: str, content: bytes
    ) -> dict[str, Any]:
        """Upload (overwrite) a file into the document library backing a SharePoint list."""
        drive = self._request("GET", f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/drive").json()
        drive_id = drive["id"]
        response = self._request(
            "PUT",
            f"{GRAPH_BASE}/drives/{drive_id}/root:/{file_name}:/content",
            data=content,
        )
        return response.json()

    def retrieve_device_app_installation_status(
        self,
        app_id: str,
        select: list[str],
    ) -> list[dict[str, Any]]:
        """Return per-device install rows from the current Intune streamed report action."""
        body = {
            "name": "DeviceInstallStatusByApp",
            "filter": f"(ApplicationId eq '{app_id}')",
            "select": select,
            "search": "",
            "groupBy": [],
            "orderBy": [],
            "skip": 0,
            "top": 10000,
            "sessionId": "",
        }
        result = self._request(
            "POST",
            f"{GRAPH_BETA_BASE}/deviceManagement/reports/retrieveDeviceAppInstallationStatusReport",
            json=body,
        ).json()
        columns = [descriptor["Column"] for descriptor in result.get("Schema", [])]
        return [dict(zip(columns, values)) for values in result.get("Values", [])]
