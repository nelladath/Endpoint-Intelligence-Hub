"""Current-month Windows patch compliance derived from Microsoft update history."""
from __future__ import annotations

import html
import json
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timezone
from html.parser import HTMLParser
from typing import Any
from urllib.parse import urljoin

import requests


UPDATE_HISTORY_URLS = (
    "https://aka.ms/Windows11UpdateHistory",
    "https://support.microsoft.com/en-us/help/4043454",
)
STALE_AFTER_DAYS = 14

OS_NAMES = {
    "28000": "Windows 11 26H1",
    "26200": "Windows 11 25H2",
    "26100": "Windows 11 24H2",
    "22631": "Windows 11 23H2",
    "22621": "Windows 11 22H2",
    "22000": "Windows 11 21H2",
    "19045": "Windows 10 22H2",
    "19044": "Windows 10 21H2",
    "19043": "Windows 10 21H1",
    "19042": "Windows 10 20H2",
    "19041": "Windows 10 2004",
    "18363": "Windows 10 1909",
    "18362": "Windows 10 1903",
    "17763": "Windows 10 1809",
    "17134": "Windows 10 1803",
    "16299": "Windows 10 1709",
    "15063": "Windows 10 1703",
    "14393": "Windows 10 1607",
    "10586": "Windows 10 1511",
    "10240": "Windows 10 1507",
}


@dataclass(frozen=True)
class PatchBaseline:
    major_build: str
    revision: int
    kb_number: str
    release_date: date
    source_url: str

    @property
    def full_build(self) -> str:
        return f"10.0.{self.major_build}.{self.revision}"


class _UpdateLinkParser(HTMLParser):
    def __init__(self, base_url: str) -> None:
        super().__init__()
        self.base_url = base_url
        self.links: list[tuple[str, str]] = []
        self._href = ""
        self._text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        values = dict(attrs)
        href = values.get("href") or ""
        if href:
            self._href = urljoin(self.base_url, href)
            self._text = []

    def handle_data(self, data: str) -> None:
        if self._href:
            self._text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "a" and self._href:
            self.links.append((" ".join(self._text), self._href))
            self._href = ""
            self._text = []


def patch_tuesday(year: int, month: int) -> date:
    first = date(year, month, 1)
    first_tuesday_offset = (1 - first.weekday()) % 7
    return date.fromordinal(first.toordinal() + first_tuesday_offset + 7)


def reporting_month(today: date) -> tuple[int, int]:
    """Use the current month after Patch Tuesday; otherwise use the prior month."""
    if today >= patch_tuesday(today.year, today.month):
        return today.year, today.month
    previous_month_end = date.fromordinal(date(today.year, today.month, 1).toordinal() - 1)
    return previous_month_end.year, previous_month_end.month


def _parse_release_date(value: str) -> date | None:
    normalized = re.sub(r"\s+", " ", value).strip(" -\u2013\u2014")
    for pattern in ("%B %d, %Y", "%B %d %Y"):
        try:
            return datetime.strptime(normalized, pattern).date()
        except ValueError:
            continue
    return None


def parse_update_history(page_html: str, base_url: str) -> list[PatchBaseline]:
    """Parse security cumulative update links from a Microsoft update-history index."""
    parser = _UpdateLinkParser(base_url)
    parser.feed(page_html)
    baselines: list[PatchBaseline] = []
    for raw_text, source_url in parser.links:
        text = re.sub(r"\s+", " ", html.unescape(raw_text)).strip()
        lowered = text.lower()
        if "kb" not in lowered or "os build" not in lowered:
            continue
        if "preview" in lowered or "out-of-band" in lowered or "mobile" in lowered:
            continue
        kb_match = re.search(r"\bKB\d{6,8}\b", text, re.IGNORECASE)
        date_match = re.search(r"([A-Z][a-z]+\s+\d{1,2},?\s+\d{4})", text)
        build_match = re.search(r"OS Builds?\s+(.+?)(?:\s+[\u2014-]\s+|$)", text, re.IGNORECASE)
        if not kb_match or not date_match or not build_match:
            continue
        released = _parse_release_date(date_match.group(1))
        if not released:
            continue
        for major, revision in re.findall(r"\b(\d{5})\.(\d+)\b", build_match.group(1)):
            baselines.append(PatchBaseline(
                major_build=major,
                revision=int(revision),
                kb_number=kb_match.group(0).upper(),
                release_date=released,
                source_url=source_url,
            ))
    return baselines


def fetch_patch_catalog(session: requests.Session | None = None) -> list[PatchBaseline]:
    web = session or requests.Session()
    catalog: list[PatchBaseline] = []
    for url in UPDATE_HISTORY_URLS:
        response = web.get(url, timeout=30, headers={"User-Agent": "Endpoint-Intelligence-Hub/1.0"})
        response.raise_for_status()
        catalog.extend(parse_update_history(response.text, response.url))
    unique: dict[tuple[str, int], PatchBaseline] = {}
    for entry in catalog:
        unique[(entry.major_build, entry.revision)] = entry
    if not unique:
        raise RuntimeError("Microsoft update history returned no usable security update entries.")
    return list(unique.values())


def _parse_datetime(value: Any) -> datetime | None:
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)
    except ValueError:
        return None


def build_patch_compliance(
    devices: list[dict[str, Any]],
    catalog: list[PatchBaseline],
    now: datetime | None = None,
) -> list[dict[str, Any]]:
    """Compare each managed Windows device with its current Patch Tuesday baseline."""
    checked_at = now or datetime.now(timezone.utc)
    year, month = reporting_month(checked_at.date())
    month_key = f"{year:04d}-{month:02d}"
    monthly = [entry for entry in catalog if (entry.release_date.year, entry.release_date.month) == (year, month)]
    required: dict[str, PatchBaseline] = {}
    for entry in monthly:
        current = required.get(entry.major_build)
        if current is None or entry.revision > current.revision:
            required[entry.major_build] = entry
    known_builds = {entry.major_build for entry in catalog}

    records: list[dict[str, Any]] = []
    for device in devices:
        if str(device.get("operatingSystem") or "").lower() != "windows":
            continue
        version = str(device.get("osVersion") or "")
        version_match = re.fullmatch(r"10\.0\.(\d+)\.(\d+)", version)
        major = version_match.group(1) if version_match else ""
        installed_revision = int(version_match.group(2)) if version_match else None
        baseline = required.get(major)
        last_sync = _parse_datetime(device.get("lastSyncDateTime"))
        check_in_days = max(0, (checked_at - last_sync).days) if last_sync else None

        if not version_match:
            status = "Unknown"
        elif last_sync is None or check_in_days is None or check_in_days > STALE_AFTER_DAYS:
            status = "Stale"
        elif baseline:
            status = "Compliant" if installed_revision >= baseline.revision else "Non-Compliant"
        elif major in known_builds:
            status = "OS End of Service"
        else:
            status = "Unknown"

        records.append({
            "key": device["deviceId"],
            "deviceId": device["deviceId"],
            "azureAdDeviceId": device.get("azureAdDeviceId"),
            "deviceName": device.get("deviceName"),
            "userPrincipalName": device.get("userPrincipalName"),
            "serialNumber": device.get("serialNumber"),
            "manufacturer": device.get("manufacturer"),
            "model": device.get("model"),
            "osBranch": OS_NAMES.get(major, f"Build {major}" if major else "Unknown"),
            "installedBuild": version,
            "installedRevision": installed_revision if installed_revision is not None else "",
            "reportingMonth": month_key,
            "patchTuesdayDate": patch_tuesday(year, month).isoformat(),
            "requiredKB": baseline.kb_number if baseline else "",
            "requiredBuild": baseline.full_build if baseline else "",
            "requiredRevision": baseline.revision if baseline else "",
            "baselineReleaseDate": baseline.release_date.isoformat() if baseline else "",
            "complianceStatus": status,
            "lastCheckIn": device.get("lastSyncDateTime"),
            "lastCheckInDays": check_in_days if check_in_days is not None else "",
            "baselineSourceUrl": baseline.source_url if baseline else "",
            "evaluatedAtUtc": checked_at.isoformat(),
        })
    return records


if __name__ == "__main__":
    input_devices = json.load(sys.stdin)
    json.dump(build_patch_compliance(input_devices, fetch_patch_catalog()), sys.stdout)