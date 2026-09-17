"""Syncs normalized Intune records into SharePoint lists using an upsert + mark-and-sweep pattern.

Each SharePoint list must have a "Key" column (single line text, indexed) matching a record's
business key, and a "LastRunId" column. After upserting all current records, any existing item
whose key is absent from the current run's records is stale (no longer reported by Intune) and
is deleted -- so each list always mirrors the latest Intune state instead of accumulating history.
"""
from __future__ import annotations

import logging
from typing import Any

from graph_client import GRAPH_BASE, GraphClient

logger = logging.getLogger("intune_report.sharepoint_sync")

_PAGE_SIZE = 999


def _field_value(value: Any) -> Any:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    return value


def _list_items_url(site_id: str, list_id: str) -> str:
    return f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"


def _fetch_existing(client: GraphClient, site_id: str, list_id: str) -> dict[str, str]:
    """Return {businessKey: itemId} for every item currently in the list."""
    existing: dict[str, str] = {}
    params = {"$expand": "fields($select=Key)", "$top": _PAGE_SIZE}
    for item in client.get_paged(_list_items_url(site_id, list_id), params=params):
        key = item.get("fields", {}).get("Key")
        if key:
            existing[key] = item["id"]
    return existing


def sync_records(
    client: GraphClient,
    site_id: str,
    list_id: str,
    records: list[dict[str, Any]],
    run_id: str,
) -> dict[str, int]:
    """Upsert `records` into the target list, then delete any item not present in this run."""
    existing = _fetch_existing(client, site_id, list_id)
    batch_requests: list[dict[str, Any]] = []
    created = updated = 0

    for index, record in enumerate(records):
        fields = {name: _field_value(value) for name, value in record.items()}
        fields["LastRunId"] = run_id
        key = fields.pop("key")
        item_id = existing.get(key)
        if item_id:
            batch_requests.append({
                "id": f"upd-{index}",
                "method": "PATCH",
                "url": f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields",
                "body": fields,
                "headers": {"Content-Type": "application/json"},
            })
            updated += 1
        else:
            batch_requests.append({
                "id": f"new-{index}",
                "method": "POST",
                "url": f"/sites/{site_id}/lists/{list_id}/items",
                "body": {"fields": {**fields, "Key": key}},
                "headers": {"Content-Type": "application/json"},
            })
            created += 1

    _log_failures(client.batch(batch_requests))

    current_keys = {record["key"] for record in records}
    stale_item_ids = [item_id for key, item_id in existing.items() if key not in current_keys]
    deleted = _delete_items(client, site_id, list_id, stale_item_ids)

    return {"created": created, "updated": updated, "deleted": deleted}


def _delete_items(client: GraphClient, site_id: str, list_id: str, item_ids: list[str]) -> int:
    if not item_ids:
        return 0
    batch_requests = [
        {"id": f"del-{i}", "method": "DELETE", "url": f"/sites/{site_id}/lists/{list_id}/items/{item_id}"}
        for i, item_id in enumerate(item_ids)
    ]
    _log_failures(client.batch(batch_requests))
    return len(item_ids)


def _log_failures(responses: list[dict[str, Any]]) -> None:
    failures = [response for response in responses if response.get("status", 200) >= 400]
    for response in failures:
        logger.error("SharePoint batch operation failed: %s", response)
    if failures:
        first = failures[0]
        raise RuntimeError(
            f"{len(failures)} SharePoint batch operation(s) failed; first failure: "
            f"status={first.get('status')} body={first.get('body')}"
        )
