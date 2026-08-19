"""AIHOT leaderboard price/score reference sync.

AIHOT exposes the leaderboard as a public HTML page rather than a leaderboard
JSON endpoint.  This module keeps the HTML adapter isolated from token
accounting: USD/cache prices remain canonical, while AIHOT supplies a current
CNY input/output reference and provenance.
"""
from __future__ import annotations

import html as html_lib
import json
import os
import re
from datetime import date
from pathlib import Path
from typing import Callable, Optional
from urllib.request import Request, urlopen

DEFAULT_ENDPOINT = "https://aihot.virxact.com/leaderboard"
_ROW_RE = re.compile(
    r'<a\s+class="lb-row"[^>]*href="(?P<href>/leaderboard/[^"?]+)"[^>]*>(?P<body>.*?)</a>',
    re.I | re.S,
)
_FIELD_RE = re.compile(
    r'<span\s+class="(?P<class>[^" ]*(?:lb-model|lb-release-date|lb-completeness|lb-input-price|lb-output-price|lb-score)[^" ]*)"(?P<attrs>[^>]*)>(?P<body>.*?)</span>',
    re.I | re.S,
)
_TEXT_RE = re.compile(r"<[^>]+>", re.S)
_ATTR_RE = re.compile(r'(?P<key>[\w-]+)="(?P<value>[^"]*)"', re.I)


def _text(value: str) -> str:
    return re.sub(r"\s+", " ", html_lib.unescape(_TEXT_RE.sub("", value))).strip()


def _attrs(value: str) -> dict[str, str]:
    return {m.group("key"): html_lib.unescape(m.group("value")) for m in _ATTR_RE.finditer(value)}


def parse_leaderboard_html(page: str) -> list[dict]:
    """Parse public AIHOT leaderboard rows into stable, JSON-serializable data."""
    result = []
    for row in _ROW_RE.finditer(page):
        body = row.group("body")
        fields: dict[str, str] = {}
        attrs: dict[str, str] = {}
        for field in _FIELD_RE.finditer(body):
            cls = field.group("class")
            key = next((x for x in ("lb-model", "lb-release-date", "lb-completeness", "lb-input-price", "lb-output-price", "lb-score") if x in cls), None)
            if key:
                fields[key] = _text(field.group("body"))
                attrs[key] = _attrs(field.group("attrs")).get("aria-label", "")
        model_match = re.search(r"<strong>(.*?)</strong><small>(.*?)</small>", body, re.S | re.I)
        if not model_match or "lb-score" not in fields:
            continue
        model = _text(model_match.group(1))
        provider = _text(model_match.group(2))
        score = _number(fields.get("lb-score", ""))
        completeness = _number(fields.get("lb-completeness", ""))
        result.append({
            "slug": row.group("href").rsplit("/", 1)[-1],
            "name": model,
            "provider": provider,
            "released_at": fields.get("lb-release-date", ""),
            "coverage_percent": completeness,
            "input_cny_per_million": _number(fields.get("lb-input-price", "")),
            "output_cny_per_million": _number(fields.get("lb-output-price", "")),
            "score": score,
            "pricing_source": attrs.get("lb-input-price", "") or attrs.get("lb-output-price", ""),
            "leaderboard_url": "https://aihot.virxact.com" + row.group("href"),
        })
    return result


def _number(value: str) -> Optional[float]:
    m = re.search(r"\d+(?:\.\d+)?", value or "")
    return float(m.group(0)) if m else None


def normalize_model(model: str) -> str:
    value = (model or "").strip().lower()
    for prefix in ("litellm/", "openai/", "anthropic/", "google/"):
        if value.startswith(prefix):
            value = value[len(prefix):]
    return re.sub(r"[^a-z0-9]+", "-", value).strip("-")


def find_reference(rows: list[dict], model: str) -> Optional[dict]:
    wanted = normalize_model(model)
    if not wanted:
        return None
    for row in rows:
        candidates = {normalize_model(row.get("slug", "")), normalize_model(row.get("name", ""))}
        if wanted in candidates:
            return row
    for row in rows:
        slug = normalize_model(row.get("slug", ""))
        if slug and (wanted.startswith(slug) or slug.startswith(wanted)):
            return row
    return None


def sync_leaderboard(path: str | os.PathLike[str], *, today: Optional[str] = None,
                     endpoint: str = DEFAULT_ENDPOINT, timeout: int = 20,
                     opener: Callable = urlopen) -> tuple[list[dict], bool, Optional[str]]:
    """Refresh once per day; preserve the previous cache on network/parse failure."""
    today = today or date.today().isoformat()
    target = Path(path)
    try:
        old = json.loads(target.read_text(encoding="utf-8"))
        if old.get("date") == today and isinstance(old.get("models"), list) and old["models"]:
            return old["models"], False, None
    except (OSError, ValueError, TypeError):
        old = {"models": []}
    try:
        request = Request(endpoint, headers={"User-Agent": "NumDesTools-token-stats/1.0"})
        with opener(request, timeout=timeout) as response:
            page = response.read().decode("utf-8", errors="replace")
        rows = parse_leaderboard_html(page)
        if not rows:
            raise ValueError("AIHOT leaderboard returned no model rows")
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = Path(str(target) + ".tmp")
        tmp.write_text(json.dumps({"date": today, "source": endpoint, "models": rows}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        os.replace(tmp, target)
        return rows, True, None
    except (OSError, ValueError, TypeError) as exc:
        return old.get("models", []), False, f"AIHOT leaderboard sync failed: {exc}"
