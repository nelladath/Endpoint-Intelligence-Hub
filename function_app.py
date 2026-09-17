"""Endpoint Intelligence Hub HTTP and daily scheduled Azure Function entry points."""
from __future__ import annotations

import json
import logging

import azure.functions as func

from pipeline import REPORT_NAME, run_pipeline

app = func.FunctionApp()
logger = logging.getLogger("endpoint_intelligence_hub")


@app.function_name(name="IntuneReportHttpTrigger")
@app.route(route="intune-report", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def http_trigger(req: func.HttpRequest) -> func.HttpResponse:
    """On-demand entry point for Power Automate: run the full collect + sync pipeline immediately."""
    try:
        result = run_pipeline()
        return func.HttpResponse(json.dumps(result), status_code=200, mimetype="application/json")
    except Exception as exc:  # noqa: BLE001 - surface any failure to the calling flow
        logger.exception("%s pipeline failed", REPORT_NAME)
        return func.HttpResponse(json.dumps({"error": str(exc)}), status_code=500, mimetype="application/json")


@app.function_name(name="IntuneReportTimerTrigger")
@app.schedule(schedule="0 0 8 * * *", arg_name="timer", run_on_startup=False, use_monitor=True)
def timer_trigger(timer: func.TimerRequest) -> None:
    """Run Endpoint Intelligence Hub collection daily at 08:00 UTC."""
    run_pipeline()
