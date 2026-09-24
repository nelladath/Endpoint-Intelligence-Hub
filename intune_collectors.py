"""Fetches raw Intune app / compliance / configuration / patch / device data from Microsoft Graph."""
from __future__ import annotations

from typing import Any

from error_catalog import lookup_error
from graph_client import GRAPH_BASE, GRAPH_BETA_BASE, GraphClient


def _classify_app_type(odata_type: str) -> str:
    value = odata_type.rsplit(".", 1)[-1]
    normalized = value.lower()
    if value == "win32LobApp":
        return "Win32"
    if any(marker in normalized for marker in ("microsoftstore", "windowsstore", "winget")):
        return "Microsoft Store"
    if value in {"officeSuiteApp", "macOSOfficeSuiteApp"}:
        return "Microsoft 365 Apps"
    if "ios" in normalized:
        return "iOS"
    if "android" in normalized:
        return "Android"
    if "macos" in normalized:
        return "macOS"
    return value or "Other"


def collect_application_inventory(client: GraphClient) -> list[dict[str, Any]]:
    """Application-level inventory, including Win32 and Store classification."""
    records: list[dict[str, Any]] = []
    for app in client.get_paged(f"{GRAPH_BASE}/deviceAppManagement/mobileApps"):
        app_id = app["id"]
        records.append({
            "key": app_id,
            "appId": app_id,
            "appName": app.get("displayName", ""),
            "appType": _classify_app_type(app.get("@odata.type", "")),
            "publisher": app.get("publisher", ""),
            "displayVersion": app.get("displayVersion") or "",
            "owner": app.get("owner") or "",
            "developer": app.get("developer") or "",
            "publishingState": app.get("publishingState") or "",
            "isAssigned": app.get("isAssigned"),
            "createdDateTime": app.get("createdDateTime"),
            "lastModifiedDateTime": app.get("lastModifiedDateTime"),
        })
    return records


def collect_autopilot_devices(client: GraphClient) -> list[dict[str, Any]]:
    """Windows Autopilot registration and deployment profile assignment inventory."""
    records: list[dict[str, Any]] = []
    for device in client.get_paged(
        f"{GRAPH_BETA_BASE}/deviceManagement/windowsAutopilotDeviceIdentities"
    ):
        autopilot_id = device["id"]
        records.append({
            "key": autopilot_id,
            "autopilotId": autopilot_id,
            "serialNumber": device.get("serialNumber"),
            "manufacturer": device.get("manufacturer"),
            "model": device.get("model"),
            "groupTag": device.get("groupTag"),
            "enrollmentState": device.get("enrollmentState"),
            "profileAssignmentStatus": device.get("deploymentProfileAssignmentStatus"),
            "profileAssignmentDetail": device.get("deploymentProfileAssignmentDetailedStatus"),
            "profileAssignedDateTime": device.get("deploymentProfileAssignedDateTime"),
            "lastContactedDateTime": device.get("lastContactedDateTime"),
            "userPrincipalName": device.get("userPrincipalName"),
            "managedDeviceId": device.get("managedDeviceId"),
            "azureAdDeviceId": device.get("azureAdDeviceId") or device.get("azureActiveDirectoryDeviceId"),
            "userlessEnrollmentStatus": device.get("userlessEnrollmentStatus"),
        })
    return records


def _policy_inventory_record(policy: dict[str, Any], policy_type: str) -> dict[str, Any]:
    policy_id = policy["id"]
    role_scope_tags = policy.get("roleScopeTagIds") or []
    if isinstance(role_scope_tags, list):
        role_scope_tags = ",".join(str(tag) for tag in role_scope_tags)
    return {
        "key": policy_id,
        "policyId": policy_id,
        "policyName": policy.get("displayName", ""),
        "policyType": policy_type,
        "odataType": policy.get("@odata.type", ""),
        "description": policy.get("description") or "",
        "platforms": policy.get("platforms") or "",
        "technologies": policy.get("technologies") or "",
        "isAssigned": policy.get("isAssigned"),
        "roleScopeTagIds": role_scope_tags,
        "createdDateTime": policy.get("createdDateTime"),
        "lastModifiedDateTime": policy.get("lastModifiedDateTime"),
    }


def collect_compliance_policy_inventory(client: GraphClient) -> list[dict[str, Any]]:
    """Return every compliance policy definition, including policies with no status rows."""
    return [
        _policy_inventory_record(policy, "Compliance Policy")
        for policy in client.get_paged(f"{GRAPH_BASE}/deviceManagement/deviceCompliancePolicies")
    ]


def collect_configuration_profile_inventory(client: GraphClient) -> list[dict[str, Any]]:
    """Return every configuration profile definition, including profiles with no status rows."""
    return [
        _policy_inventory_record(policy, "Configuration Profile")
        for policy in client.get_paged(f"{GRAPH_BASE}/deviceManagement/deviceConfigurations")
    ]


def collect_update_ring_inventory(client: GraphClient) -> list[dict[str, Any]]:
    """Return every Windows Update for Business ring definition."""
    profiles = client.get_paged(
        f"{GRAPH_BASE}/deviceManagement/deviceConfigurations",
        params={"$filter": "isof('microsoft.graph.windowsUpdateForBusinessConfiguration')"},
    )
    return [
        _policy_inventory_record(policy, "Windows Update Ring")
        for policy in profiles
    ]


def collect_app_deployment_status(client: GraphClient) -> list[dict[str, Any]]:
    """Per-device install status for every mobile app managed in Intune.

    Graph removed mobileApp.deviceStatuses, so this uses the current Intune reporting action.
    """
    records: list[dict[str, Any]] = []
    apps = client.get_paged(f"{GRAPH_BASE}/deviceAppManagement/mobileApps")
    for app in apps:
        app_id = app["id"]
        app_name = app.get("displayName", "")
        publisher = app.get("publisher", "")
        statuses = client.retrieve_device_app_installation_status(
            app_id,
            [
                "DeviceId", "DeviceName", "UserPrincipalName",
                "AppInstallState", "AppInstallStateDetails", "HexErrorCode", "LastModifiedDateTime",
            ],
        )
        for status in statuses:
            device_id = status.get("DeviceId")
            user_principal_name = status.get("UserPrincipalName")
            install_state = status.get("AppInstallState_loc") or status.get("AppInstallState", "unknown")
            error_code = status.get("HexErrorCode")
            issue, recommendation = lookup_error(error_code)
            records.append({
                "key": f"{app_id}_{device_id}_{user_principal_name}",
                "appId": app_id,
                "appName": app_name,
                "publisher": publisher,
                "deviceId": device_id,
                "deviceName": status.get("DeviceName"),
                "userPrincipalName": user_principal_name,
                "installState": install_state,
                "errorCode": error_code,
                "issue": issue,
                "recommendedSolution": recommendation,
                "lastSyncDateTime": status.get("LastModifiedDateTime"),
            })
    return records


def collect_compliance_policy_status(client: GraphClient) -> list[dict[str, Any]]:
    """Per-device compliance state for every device compliance policy."""
    records: list[dict[str, Any]] = []
    policies = client.get_paged(f"{GRAPH_BASE}/deviceManagement/deviceCompliancePolicies")
    for policy in policies:
        policy_id = policy["id"]
        policy_name = policy.get("displayName", "")
        statuses = client.get_paged(
            f"{GRAPH_BASE}/deviceManagement/deviceCompliancePolicies/{policy_id}/deviceStatuses"
        )
        for status in statuses:
            device_id = status.get("deviceId") or status.get("id")
            is_error = status.get("status") == "error"
            issue, recommendation = lookup_error(status.get("errorCode")) if is_error else (
                "No specific error recorded.", "No action required.",
            )
            records.append({
                "key": f"{policy_id}_{device_id}",
                "policyId": policy_id,
                "policyName": policy_name,
                "deviceId": device_id,
                "deviceName": status.get("deviceDisplayName"),
                "userPrincipalName": status.get("userPrincipalName"),
                "complianceState": status.get("status", "unknown"),
                "issue": issue,
                "recommendedSolution": recommendation,
                "lastReportedDateTime": status.get("lastReportedDateTime"),
            })
    return records


def collect_config_profile_status(client: GraphClient) -> list[dict[str, Any]]:
    """Per-device deployment status for every device configuration profile (excludes update rings)."""
    records: list[dict[str, Any]] = []
    profiles = client.get_paged(f"{GRAPH_BASE}/deviceManagement/deviceConfigurations")
    for profile in profiles:
        odata_type = profile.get("@odata.type", "")
        if odata_type.endswith("windowsUpdateForBusinessConfiguration"):
            continue  # handled separately as patch/update-ring compliance
        profile_id = profile["id"]
        profile_name = profile.get("displayName", "")
        statuses = client.get_paged(
            f"{GRAPH_BASE}/deviceManagement/deviceConfigurations/{profile_id}/deviceStatuses"
        )
        for status in statuses:
            device_id = status.get("deviceId") or status.get("id")
            records.append({
                "key": f"{profile_id}_{device_id}",
                "profileId": profile_id,
                "profileName": profile_name,
                "profileType": odata_type.split(".")[-1],
                "deviceId": device_id,
                "deviceName": status.get("deviceDisplayName"),
                "userPrincipalName": status.get("userName"),
                "status": status.get("status", "unknown"),
                "lastReportedDateTime": status.get("lastReportedDateTime"),
            })
    return records


def collect_patch_compliance_status(client: GraphClient) -> list[dict[str, Any]]:
    """Per-device deployment status for Windows Update rings (patch compliance)."""
    records: list[dict[str, Any]] = []
    profiles = client.get_paged(
        f"{GRAPH_BASE}/deviceManagement/deviceConfigurations",
        params={"$filter": "isof('microsoft.graph.windowsUpdateForBusinessConfiguration')"},
    )
    for profile in profiles:
        profile_id = profile["id"]
        profile_name = profile.get("displayName", "")
        statuses = client.get_paged(
            f"{GRAPH_BASE}/deviceManagement/deviceConfigurations/{profile_id}/deviceStatuses"
        )
        for status in statuses:
            device_id = status.get("deviceId") or status.get("id")
            records.append({
                "key": f"{profile_id}_{device_id}",
                "ringId": profile_id,
                "ringName": profile_name,
                "deviceId": device_id,
                "deviceName": status.get("deviceDisplayName"),
                "userPrincipalName": status.get("userName"),
                "status": status.get("status", "unknown"),
                "lastReportedDateTime": status.get("lastReportedDateTime"),
            })
    return records


def collect_managed_devices(client: GraphClient) -> list[dict[str, Any]]:
    """Managed device inventory with endpoint health, lifecycle, and hardware fields."""
    records: list[dict[str, Any]] = []
    select = ",".join([
        "id", "deviceName", "userId", "userPrincipalName", "operatingSystem", "osVersion",
        "manufacturer", "model", "complianceState", "managementState", "deviceEnrollmentType",
        "jailBroken", "lastSyncDateTime", "enrolledDateTime", "isEncrypted", "serialNumber",
        "totalStorageSpaceInBytes", "freeStorageSpaceInBytes", "partnerReportedThreatState",
        "complianceGracePeriodExpirationDateTime", "managementCertificateExpirationDate",
        "azureADDeviceId", "managedDeviceOwnerType", "managementAgent", "autopilotEnrolled",
        "requireUserEnrollmentApproval", "joinType", "physicalMemoryInBytes",
        "processorArchitecture", "enrollmentProfileName",
    ])
    devices = client.get_paged(
        f"{GRAPH_BETA_BASE}/deviceManagement/managedDevices", params={"$select": select}
    )
    for device in devices:
        records.append({
            "key": device["id"],
            "deviceId": device["id"],
            "deviceName": device.get("deviceName"),
            "userId": device.get("userId"),
            "userPrincipalName": device.get("userPrincipalName"),
            "operatingSystem": device.get("operatingSystem"),
            "osVersion": device.get("osVersion"),
            "manufacturer": device.get("manufacturer"),
            "model": device.get("model"),
            "serialNumber": device.get("serialNumber"),
            "complianceState": device.get("complianceState", "unknown"),
            "managementState": device.get("managementState", "unknown"),
            "enrollmentType": device.get("deviceEnrollmentType"),
            "jailBroken": device.get("jailBroken"),
            "isEncrypted": device.get("isEncrypted"),
            "totalStorageSpaceInBytes": device.get("totalStorageSpaceInBytes"),
            "freeStorageSpaceInBytes": device.get("freeStorageSpaceInBytes"),
            "threatState": device.get("partnerReportedThreatState"),
            "gracePeriodExpirationDateTime": device.get("complianceGracePeriodExpirationDateTime"),
            "managementCertExpirationDate": device.get("managementCertificateExpirationDate"),
            "azureAdDeviceId": device.get("azureADDeviceId"),
            "ownerType": device.get("managedDeviceOwnerType"),
            "managementAgent": device.get("managementAgent"),
            "autopilotEnrolled": device.get("autopilotEnrolled"),
            "approvalPending": device.get("requireUserEnrollmentApproval"),
            "joinType": device.get("joinType"),
            "physicalMemoryInBytes": device.get("physicalMemoryInBytes"),
            "processorArchitecture": device.get("processorArchitecture"),
            "enrollmentProfileName": device.get("enrollmentProfileName"),
            "enrolledDateTime": device.get("enrolledDateTime"),
            "lastSyncDateTime": device.get("lastSyncDateTime"),
        })
    return records
