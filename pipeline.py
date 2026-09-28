"""Endpoint Intelligence Hub collection, analytics, and SharePoint pipeline."""
from __future__ import annotations

import json
import logging
import csv
import io

from config import settings
from graph_client import GraphClient
from intune_analytics import (
    aggregate_application_deployments,
    build_health_metrics,
    build_report,
    enrich_device_health,
    format_record_datetimes,
)
from intune_collectors import (
    collect_app_deployment_status,
    collect_application_inventory,
    collect_autopilot_devices,
    collect_compliance_policy_inventory,
    collect_compliance_policy_status,
    collect_config_profile_status,
    collect_configuration_profile_inventory,
    collect_managed_devices,
    collect_update_ring_inventory,
)
from patch_compliance import build_patch_compliance, fetch_patch_catalog
from sharepoint_sync import sync_records

REPORT_NAME = "Endpoint Intelligence Hub"
REPORT_DESCRIPTION = (
    "Centralized portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation, "
    "patch management, application delivery, device compliance, endpoint health, security risk, "
    "hardware intelligence, and operational analytics."
)

logger = logging.getLogger("endpoint_intelligence_hub")


def records_to_csv(records: list[dict[str, object]]) -> bytes:
    """Serialize flat report records as an Excel-friendly UTF-8 CSV."""
    if not records:
        return b""
    field_names = list(dict.fromkeys(
        key for record in records for key in record if key != "key"
    ))
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=field_names, extrasaction="ignore")
    writer.writeheader()
    for record in records:
        writer.writerow({
            key: json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else value
            for key, value in record.items() if key != "key"
        })
    return output.getvalue().encode("utf-8-sig")


def run_pipeline() -> dict[str, object]:
    """Collect Intune telemetry, calculate health, and synchronize all SharePoint lists."""
    client = GraphClient()

    apps = collect_app_deployment_status(client)
    compliance_policies = collect_compliance_policy_status(client)
    compliance_inventory = collect_compliance_policy_inventory(client)
    config_profiles = collect_config_profile_status(client)
    configuration_inventory = collect_configuration_profile_inventory(client)
    update_ring_inventory = collect_update_ring_inventory(client)
    devices = collect_managed_devices(client)
    patch_compliance = build_patch_compliance(devices, fetch_patch_catalog())
    application_inventory = collect_application_inventory(client)
    aggregate_application_deployments(application_inventory, apps)
    autopilot = collect_autopilot_devices(client)
    device_risks = enrich_device_health(devices)
    health_metrics = build_health_metrics(devices, autopilot, application_inventory)
    for device in devices:
        device.pop("patchStatus", None)

    # Calculations above use raw Graph timestamps; SharePoint/JSON receive readable UTC values.
    for records in (
        apps, compliance_policies, compliance_inventory, config_profiles,
        configuration_inventory, patch_compliance, update_ring_inventory,
        devices, application_inventory, autopilot, device_risks,
    ):
        format_record_datetimes(records)

    report = build_report(
        apps,
        compliance_policies,
        config_profiles,
        patch_compliance,
        devices,
        application_inventory,
        autopilot,
        device_risks,
        health_metrics,
    )
    run_id = report["runId"]
    health_metrics.append({
        "key": "Report Run:Last Refreshed (UTC)",
        "category": "Report Run",
        "metricName": "Last Refreshed (UTC)",
        "metricValue": report["generatedAtUtc"],
        "percentage": "",
        "status": "",
        "description": "Timestamp when this pipeline run collected data from Microsoft Graph, in UTC.",
    })
    health_metrics.append({
        "key": "Report Run:Run ID",
        "category": "Report Run",
        "metricName": "Run ID",
        "metricValue": run_id,
        "percentage": "",
        "status": "",
        "description": "Unique identifier for this run; every list item from this run shares this value in LastRunId.",
    })
    list_ids = settings.sharepoint_list_ids
    site_id = settings.sharepoint_site_id
    sync_results = {
        "apps": sync_records(client, site_id, list_ids["apps"], apps, run_id),
        "compliancePolicies": sync_records(
            client, site_id, list_ids["compliance_policies"], compliance_policies, run_id
        ),
        "configProfiles": sync_records(
            client, site_id, list_ids["config_profiles"], config_profiles, run_id
        ),
        "patchCompliance": sync_records(
            client, site_id, list_ids["patch_compliance"], patch_compliance, run_id
        ),
        "devices": sync_records(client, site_id, list_ids["devices"], devices, run_id),
        "applicationInventory": sync_records(
            client, site_id, list_ids["application_inventory"], application_inventory, run_id
        ),
        "autopilot": sync_records(client, site_id, list_ids["autopilot"], autopilot, run_id),
        "deviceRisks": sync_records(
            client, site_id, list_ids["device_risks"], device_risks, run_id
        ),
        "healthSummary": sync_records(
            client, site_id, list_ids["health_summary"], health_metrics, run_id
        ),
        "complianceInventory": sync_records(
            client, site_id, list_ids["compliance_inventory"], compliance_inventory, run_id
        ),
        "configurationInventory": sync_records(
            client, site_id, list_ids["configuration_inventory"], configuration_inventory, run_id
        ),
        "updateRingInventory": sync_records(
            client, site_id, list_ids["update_ring_inventory"], update_ring_inventory, run_id
        ),
    }
    overall_health = next(
        metric for metric in health_metrics
        if metric["metricName"] == "Overall Endpoint Health Score"
    )

    json_export = {
        "reportName": REPORT_NAME,
        "description": REPORT_DESCRIPTION,
        "runId": run_id,
        "generatedAtUtc": report["generatedAtUtc"],
        "summary": report["summary"],
        "apps": apps,
        "compliancePolicies": compliance_policies,
        "complianceInventory": compliance_inventory,
        "configProfiles": config_profiles,
        "configurationInventory": configuration_inventory,
        "patchCompliance": patch_compliance,
        "updateRingInventory": update_ring_inventory,
        "devices": devices,
        "applicationInventory": application_inventory,
        "autopilot": autopilot,
        "deviceRisks": device_risks,
        "healthMetrics": health_metrics,
        "syncResults": sync_results,
    }
    client.upload_file_to_list_drive(
        site_id,
        list_ids["json_exports"],
        "EndpointIntelligenceHub_Latest.json",
        json.dumps(json_export, indent=2, default=str).encode("utf-8"),
        "application/json",
    )
    csv_exports = {
        "Application_Inventory.csv": application_inventory,
        "Application_Health_Details.csv": apps,
        "Devices.csv": devices,
        "Patch_Compliance.csv": patch_compliance,
        "Compliance_Policy_Status.csv": compliance_policies,
        "Compliance_Policy_Inventory.csv": compliance_inventory,
        "Configuration_Profile_Status.csv": config_profiles,
        "Configuration_Profile_Inventory.csv": configuration_inventory,
        "Update_Ring_Inventory.csv": update_ring_inventory,
        "Autopilot.csv": autopilot,
        "Device_Risks.csv": device_risks,
        "Health_Summary.csv": health_metrics,
    }
    for file_name, records in csv_exports.items():
        client.upload_file_to_list_drive(
            site_id, list_ids["json_exports"], file_name,
            records_to_csv(records), "text/csv; charset=utf-8",
        )

    logger.info("%s run %s complete: %s", REPORT_NAME, run_id, sync_results)
    return {
        "reportName": REPORT_NAME,
        "description": REPORT_DESCRIPTION,
        "runId": run_id,
        "summary": report["summary"],
        "endpointHealth": overall_health,
        "syncResults": sync_results,
    }