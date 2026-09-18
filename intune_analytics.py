"""Derives status normalization, success ratios, and health classification for the Intune report."""
from __future__ import annotations

import uuid
from collections import Counter, defaultdict
from datetime import datetime, timezone
from typing import Any

SUCCESS_STATES = {"installed", "compliant", "success"}
FAILED_STATES = {"failed", "uninstallfailed", "noncompliant", "error", "conflict"}
PENDING_STATES = {"pendinginstall", "installpending", "pending", "ingraceperiod"}
NOT_APPLICABLE_STATES = {"notapplicable"}

HEALTHY_THRESHOLD = 0.95
WARNING_THRESHOLD = 0.80

HEALTH_WEIGHTS = {
    "compliance": 0.25,
    "encryption": 0.20,
    "patchCurrency": 0.20,
    "threatFree": 0.15,
    "deviceActivity": 0.10,
    "managementHealth": 0.10,
}


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() == "true"


def _parse_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def classify_os_currency(operating_system: str | None, os_version: str | None) -> str:
    """Classify Windows servicing currency; non-Windows platforms are treated as current."""
    if (operating_system or "").lower() != "windows":
        return "Current"
    parts = (os_version or "").split(".")
    if len(parts) < 3 or not parts[2].isdigit():
        return "Unknown"
    build = int(parts[2])
    if build < 22000:
        return "Critical"
    if build < 22631:
        return "Very outdated"
    if build == 22631:
        return "Slightly behind"
    if build < 26200:
        return "Recent"
    return "Current"


def enrich_device_health(
    devices: list[dict[str, Any]], now: datetime | None = None
) -> list[dict[str, Any]]:
    """Add storage, activity, OS-currency, certificate, and risk fields to device records."""
    reference_time = now or datetime.now(timezone.utc)
    risks: list[dict[str, Any]] = []
    for device in devices:
        last_sync = _parse_datetime(device.get("lastSyncDateTime"))
        days_inactive = max(0, (reference_time - last_sync).days) if last_sync else None
        device["daysInactive"] = days_inactive
        device["isStale"] = days_inactive is None or days_inactive > 30

        total_storage = device.get("totalStorageSpaceInBytes") or 0
        free_storage = device.get("freeStorageSpaceInBytes") or 0
        storage_percent = round((free_storage / total_storage) * 100, 1) if total_storage else None
        device["storagePercentFree"] = storage_percent
        if storage_percent is None:
            device["storageHealth"] = "Unknown"
        elif storage_percent < 2:
            device["storageHealth"] = "Critical"
        elif storage_percent < 10:
            device["storageHealth"] = "Low"
        else:
            device["storageHealth"] = "Healthy"

        device["patchStatus"] = classify_os_currency(
            device.get("operatingSystem"), device.get("osVersion")
        )
        certificate_expiry = _parse_datetime(device.get("managementCertExpirationDate"))
        device["certificateStatus"] = (
            "Unknown" if certificate_expiry is None
            else "Expired" if certificate_expiry < reference_time
            else "Valid"
        )

        score = 0
        reasons: list[str] = []
        threat_state = (device.get("threatState") or "unknown").lower()
        management_state = (device.get("managementState") or "unknown").lower()
        if _as_bool(device.get("jailBroken")):
            score += 3
            reasons.append("Device jailbroken/rooted")
        if threat_state == "highseverity":
            score += 3
            reasons.append("High-severity threat reported")
        elif threat_state == "mediumseverity":
            score += 2
            reasons.append("Medium-severity threat reported")
        elif threat_state == "lowseverity":
            score += 1
            reasons.append("Low-severity threat reported")
        if device.get("isEncrypted") is False:
            score += 1
            reasons.append("Disk not encrypted")
        if management_state == "wipefailed":
            score += 3
            reasons.append("Remote wipe failed")
        elif management_state == "wipepending":
            score += 1
            reasons.append("Remote wipe pending")
        if (device.get("complianceState") or "").lower() == "noncompliant":
            score += 1
            reasons.append("Non-compliant")
        if days_inactive is not None and days_inactive > 90:
            score += 2
            reasons.append(f"Inactive {days_inactive} days")

        risk_level = "High" if score >= 4 else "Medium" if score >= 2 else "Low" if score else "None"
        device["riskScore"] = score
        device["riskLevel"] = risk_level
        device["riskReasons"] = "; ".join(reasons)
        if score >= 2:
            risks.append({
                "key": device["deviceId"],
                "deviceId": device["deviceId"],
                "deviceName": device.get("deviceName"),
                "userPrincipalName": device.get("userPrincipalName"),
                "operatingSystem": device.get("operatingSystem"),
                "osVersion": device.get("osVersion"),
                "riskLevel": risk_level,
                "riskScore": str(score),
                "riskReasons": device["riskReasons"],
                "daysInactive": str(days_inactive) if days_inactive is not None else "",
                "lastSyncDateTime": device.get("lastSyncDateTime"),
            })
    return risks


def build_health_metrics(
    devices: list[dict[str, Any]],
    autopilot: list[dict[str, Any]],
    applications: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Build auditable fleet health metrics suitable for the SharePoint summary list."""
    total = len(devices)
    denominator = total or 1

    def ratio(count: int, base: int = denominator) -> float:
        return count / (base or 1)

    compliance_healthy = sum(
        (device.get("complianceState") or "").lower() in {"compliant", "configmanager"}
        for device in devices
    )
    encrypted = sum(device.get("isEncrypted") is True for device in devices)
    patch_healthy = sum(device.get("patchStatus") in {"Current", "Recent"} for device in devices)
    known_threat_states = {"secured", "lowseverity", "mediumseverity", "highseverity"}
    known_threat = sum(
        (device.get("threatState") or "unknown").lower() in known_threat_states
        for device in devices
    )
    threat_free = sum(
        (device.get("threatState") or "unknown").lower() == "secured"
        for device in devices
    )
    active = sum(not device.get("isStale") for device in devices)
    managed = sum((device.get("managementState") or "").lower() == "managed" for device in devices)
    components = {
        "compliance": ratio(compliance_healthy),
        "encryption": ratio(encrypted),
        "patchCurrency": ratio(patch_healthy),
        "threatFree": ratio(threat_free),
        "deviceActivity": ratio(active),
        "managementHealth": ratio(managed),
    }
    overall_score = round(sum(components[name] * weight for name, weight in HEALTH_WEIGHTS.items()) * 100, 1)
    overall_status = "Excellent" if overall_score >= 95 else "Good" if overall_score >= 85 else "Fair" if overall_score >= 70 else "Poor"

    metrics: list[dict[str, Any]] = []

    def add(category: str, name: str, value: Any, percentage: float | None = None, status: str = "", description: str = "") -> None:
        metrics.append({
            "key": f"{category}:{name}",
            "category": category,
            "metricName": name,
            "metricValue": str(value),
            "percentage": f"{percentage:.1f}" if percentage is not None else "",
            "status": status,
            "description": description,
        })

    add(
        "Health Score",
        "Overall Endpoint Health Score",
        overall_score,
        overall_score,
        overall_status,
        "Weighted score for the current managed-device scope: "
        "Compliance 25%, Encryption 20%, Patch/OS currency 20%, Threat-free 15%, "
        "Device activity 10%, Management health 10%. "
        f"Scope: {total} managed devices returned by Microsoft Graph.",
    )
    component_labels = {
        "compliance": "Compliance (25%)",
        "encryption": "Encryption (20%)",
        "patchCurrency": "Patch/OS currency (20%)",
        "threatFree": "Threat-free (15%)",
        "deviceActivity": "Device activity (10%)",
        "managementHealth": "Management health (10%)",
    }
    component_counts = {
        "compliance": compliance_healthy,
        "encryption": encrypted,
        "patchCurrency": patch_healthy,
        "threatFree": threat_free,
        "deviceActivity": active,
        "managementHealth": managed,
    }
    component_descriptions = {
        "compliance": "Devices with complianceState compliant or configManager.",
        "encryption": "Devices where isEncrypted is true.",
        "patchCurrency": "Devices classified Current or Recent by OS build.",
        "threatFree": "Devices explicitly reporting secured; unknown threat telemetry is not counted as threat-free.",
        "deviceActivity": "Devices synced within the 30-day activity window.",
        "managementHealth": "Devices whose managementState is managed.",
    }
    for key, value in components.items():
        add(
            "Health Component",
            component_labels[key],
            round(value * 100, 1),
            value * 100,
            "",
            f"{component_counts[key]} of {total} managed devices. {component_descriptions[key]}",
        )

    counts = {
        "Total Devices": total,
        "Unique Users": len({device.get("userId") for device in devices if device.get("userId")}),
        "Non-Compliant Devices": sum((device.get("complianceState") or "").lower() == "noncompliant" for device in devices),
        "Unencrypted Devices": sum(device.get("isEncrypted") is False for device in devices),
        "Stale / Inactive Devices": sum(bool(device.get("isStale")) for device in devices),
        "Orphaned Devices": sum(not device.get("userId") for device in devices),
        "Approval Pending": sum(bool(device.get("approvalPending")) for device in devices),
        "Expired Management Certificates": sum(device.get("certificateStatus") == "Expired" for device in devices),
        "Critical Storage (<2%)": sum(device.get("storagePercentFree") is not None and device["storagePercentFree"] < 2 for device in devices),
        "Low Storage (<10%)": sum(device.get("storagePercentFree") is not None and device["storagePercentFree"] < 10 for device in devices),
        "High/Medium Risk Devices": sum(device.get("riskLevel") in {"High", "Medium"} for device in devices),
        "Jailbroken / Rooted": sum(_as_bool(device.get("jailBroken")) for device in devices),
        "Threat Telemetry Coverage": known_threat,
        "Autopilot Registered": len(autopilot),
        "Applications": len(applications),
        "Win32 Applications": sum(app.get("appType") == "Win32" for app in applications),
        "Store Applications": sum(app.get("appType") == "Microsoft Store" for app in applications),
    }
    for name, count in counts.items():
        add(
            "Fleet",
            name,
            count,
            ratio(count) * 100 if name not in {"Autopilot Registered", "Applications", "Win32 Applications", "Store Applications"} else None,
            "",
            f"{count} of {total} managed devices." if name not in {"Autopilot Registered", "Applications", "Win32 Applications", "Store Applications"} else "Current tenant inventory count.",
        )

    distributions = {
        "Operating System": Counter(device.get("operatingSystem") or "Unknown" for device in devices),
        "Compliance": Counter(device.get("complianceState") or "Unknown" for device in devices),
        "Patch Currency": Counter(device.get("patchStatus") or "Unknown" for device in devices),
        "Threat State": Counter(device.get("threatState") or "Unknown" for device in devices),
        "Management State": Counter(device.get("managementState") or "Unknown" for device in devices),
        "Join Type": Counter(device.get("joinType") or "Unknown" for device in devices),
        "Autopilot Enrollment": Counter(device.get("enrollmentState") or "Unknown" for device in autopilot),
        "Application Type": Counter(app.get("appType") or "Other" for app in applications),
    }
    for category, distribution in distributions.items():
        base = len(autopilot) if category == "Autopilot Enrollment" else len(applications) if category == "Application Type" else total
        for name, count in sorted(distribution.items()):
            add(category, name, count, ratio(count, base) * 100)
    return metrics


def normalize_status(raw_status: str | None) -> str:
    """Collapse the many Graph-specific state strings into Success/Failed/Pending/NotApplicable/Unknown."""
    value = "".join(character for character in (raw_status or "unknown").lower() if character.isalnum())
    if value in SUCCESS_STATES:
        return "Success"
    if value in FAILED_STATES:
        return "Failed"
    if value in PENDING_STATES:
        return "Pending"
    if value in NOT_APPLICABLE_STATES:
        return "NotApplicable"
    return "Unknown"


def classify_health(success_ratio: float) -> str:
    if success_ratio >= HEALTHY_THRESHOLD:
        return "Healthy"
    if success_ratio >= WARNING_THRESHOLD:
        return "Warning"
    return "Critical"


def summarize_domain(records: list[dict[str, Any]], status_field: str) -> dict[str, Any]:
    """Aggregate a domain's records into total/success/failed/pending counts and a health rating.

    Mutates each record in-place, adding a `normalizedStatus` field used by the SharePoint sync.
    """
    counts: dict[str, int] = defaultdict(int)
    for record in records:
        normalized = normalize_status(record.get(status_field))
        record["normalizedStatus"] = normalized
        counts[normalized] += 1
    applicable_total = sum(v for k, v in counts.items() if k != "NotApplicable")
    success_ratio = (counts["Success"] / applicable_total) if applicable_total else 1.0
    return {
        "total": len(records),
        "success": counts["Success"],
        "failed": counts["Failed"],
        "pending": counts["Pending"],
        "notApplicable": counts["NotApplicable"],
        "unknown": counts["Unknown"],
        "successRatio": round(success_ratio, 4),
        "healthStatus": classify_health(success_ratio),
    }


def build_report(
    apps: list[dict[str, Any]],
    compliance_policies: list[dict[str, Any]],
    config_profiles: list[dict[str, Any]],
    patch_compliance: list[dict[str, Any]],
    devices: list[dict[str, Any]],
    application_inventory: list[dict[str, Any]],
    autopilot: list[dict[str, Any]],
    device_risks: list[dict[str, Any]],
    health_metrics: list[dict[str, Any]],
) -> dict[str, Any]:
    """Assemble the full Intune metadata report: raw records + per-domain summaries + a run id.

    `runId` is stamped onto every synced SharePoint item so stale rows from a prior run can be
    detected and removed (mark-and-sweep), keeping SharePoint an exact mirror of current Intune state.
    """
    return {
        "runId": str(uuid.uuid4()),
        "generatedAtUtc": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "apps": summarize_domain(apps, "installState"),
            "compliancePolicies": summarize_domain(compliance_policies, "complianceState"),
            "configProfiles": summarize_domain(config_profiles, "status"),
            "patchCompliance": summarize_domain(patch_compliance, "status"),
        },
        "apps": apps,
        "compliancePolicies": compliance_policies,
        "configProfiles": config_profiles,
        "patchCompliance": patch_compliance,
        "devices": devices,
        "applicationInventory": application_inventory,
        "autopilot": autopilot,
        "deviceRisks": device_risks,
        "healthMetrics": health_metrics,
    }
