"""Central configuration loaded from Azure Function app settings (environment variables)."""
from __future__ import annotations

import os


def _required(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"Missing required app setting: {name}")
    return value


class Settings:
    """Lazily-validated wrapper around required Azure Function app settings."""

    @property
    def tenant_id(self) -> str:
        return _required("TENANT_ID")

    @property
    def client_id(self) -> str:
        return _required("CLIENT_ID")

    @property
    def client_secret(self) -> str:
        # Store this as a Key Vault reference in the Function App's application settings, not plain text.
        return _required("CLIENT_SECRET")

    @property
    def sharepoint_site_id(self) -> str:
        return _required("SHAREPOINT_SITE_ID")

    @property
    def sharepoint_list_ids(self) -> dict[str, str]:
        return {
            "apps": _required("SP_LIST_APPS_ID"),
            "compliance_policies": _required("SP_LIST_COMPLIANCE_ID"),
            "config_profiles": _required("SP_LIST_CONFIG_PROFILES_ID"),
            "patch_compliance": _required("SP_LIST_PATCH_ID"),
            "devices": _required("SP_LIST_DEVICES_ID"),
            "application_inventory": _required("SP_LIST_APP_INVENTORY_ID"),
            "autopilot": _required("SP_LIST_AUTOPILOT_ID"),
            "device_risks": _required("SP_LIST_DEVICE_RISKS_ID"),
            "health_summary": _required("SP_LIST_HEALTH_SUMMARY_ID"),
            "compliance_inventory": _required("SP_LIST_COMPLIANCE_INVENTORY_ID"),
            "configuration_inventory": _required("SP_LIST_CONFIGURATION_INVENTORY_ID"),
            "update_ring_inventory": _required("SP_LIST_UPDATE_RING_INVENTORY_ID"),
        }


settings = Settings()
