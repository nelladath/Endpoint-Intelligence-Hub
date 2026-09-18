"""Endpoint Intelligence Hub collection, analytics, and SharePoint pipeline."""
from __future__ import annotations

import logging

from config import settings
from graph_client import GraphClient
from intune_analytics import build_health_metrics, build_report, enrich_device_health
from intune_collectors import (
    collect_app_deployment_status,
    collect_application_inventory,
    collect_autopilot_devices,
    collect_compliance_policy_inventory,
    collect_compliance_policy_status,
    collect_config_profile_status,
    collect_configuration_profile_inventory,
    collect_managed_devices,
    collect_patch_compliance_status,
    collect_update_ring_inventory,
)
from sharepoint_sync import sync_records

REPORT_NAME = "Endpoint Intelligence Hub"
REPORT_DESCRIPTION = (
    "Centralized portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation, "
    "patch management, application delivery, device compliance, endpoint health, security risk, "
    "hardware intelligence, and operational analytics."
)

logger = logging.getLogger("endpoint_intelligence_hub")


def run_pipeline() -> dict[str, object]:
    """Collect Intune telemetry, calculate health, and synchronize all SharePoint lists."""
    client = GraphClient()

    apps = collect_app_deployment_status(client)
    compliance_policies = collect_compliance_policy_status(client)
    compliance_inventory = collect_compliance_policy_inventory(client)
    config_profiles = collect_config_profile_status(client)
    configuration_inventory = collect_configuration_profile_inventory(client)
    patch_compliance = collect_patch_compliance_status(client)
    update_ring_inventory = collect_update_ring_inventory(client)
    devices = collect_managed_devices(client)
    application_inventory = collect_application_inventory(client)
    autopilot = collect_autopilot_devices(client)
    device_risks = enrich_device_health(devices)
    health_metrics = build_health_metrics(devices, autopilot, application_inventory)

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
    logger.info("%s run %s complete: %s", REPORT_NAME, run_id, sync_results)
    return {
        "reportName": REPORT_NAME,
        "description": REPORT_DESCRIPTION,
        "runId": run_id,
        "summary": report["summary"],
        "endpointHealth": overall_health,
        "syncResults": sync_results,
    }