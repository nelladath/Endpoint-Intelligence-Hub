"""Known Intune error codes mapped to a human-readable issue and recommended remediation.

This is a living catalog -- extend ERROR_CATALOG as new recurring error codes are observed
in the tenant. Unrecognized codes fall back to a generic "review the admin center" message.
"""
from __future__ import annotations

from typing import Optional, Union

ERROR_CATALOG: dict[str, dict[str, str]] = {
    "0x87D13B94": {
        "issue": "App install timed out on the device.",
        "recommendation": "Confirm the device has network connectivity and free disk space, then retry sync.",
    },
    "0x87D1041C": {
        "issue": "The app failed a dependency or requirement check.",
        "recommendation": "Verify OS version/architecture requirements and that required dependent apps are assigned.",
    },
    "0x80070005": {
        "issue": "Access denied while installing the app.",
        "recommendation": "Check the install context (system vs. user) and local permissions on the target device.",
    },
    "0x8007000E": {
        "issue": "Device ran out of memory/resources during install.",
        "recommendation": "Free up device storage/memory and re-trigger the app installation.",
    },
    "0x87D30604": {
        "issue": "Content download failed (network or proxy issue).",
        "recommendation": "Check proxy/firewall rules for Intune/Delivery Optimization endpoints and retry.",
    },
    "0x8018002A": {
        "issue": "Device is not Azure AD joined/registered correctly.",
        "recommendation": "Re-run dsregcmd /status on the device and re-join Azure AD if needed.",
    },
}

DEFAULT_ISSUE = "No specific error recorded."
DEFAULT_RECOMMENDATION = "No action required."
UNKNOWN_ISSUE = "Unrecognized error code -- not yet cataloged."
UNKNOWN_RECOMMENDATION = "Review the Intune admin center troubleshooting blade for this device/app."


def lookup_error(error_code: Optional[Union[str, int]]) -> tuple[str, str]:
    """Return (issue, recommendation) for a Graph-reported error code, or a default pair for no error."""
    if not error_code or str(error_code) in ("0", "0x0"):
        return DEFAULT_ISSUE, DEFAULT_RECOMMENDATION
    entry = ERROR_CATALOG.get(str(error_code).upper())
    if entry:
        return entry["issue"], entry["recommendation"]
    return UNKNOWN_ISSUE, UNKNOWN_RECOMMENDATION
