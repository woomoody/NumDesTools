"""Daily LiteLLM model catalog synchronization for token statistics."""
from __future__ import annotations

import json
import os
from datetime import date
from typing import Callable, Dict, List, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

DEFAULT_ENDPOINT = "https://litellm.solotopia.net/v1/models"


def extract_model_ids(payload: Dict) -> List[str]:
    """Return stable, sorted model IDs from an OpenAI-compatible /models payload."""
    result = set()
    for item in payload.get("data", []):
        if isinstance(item, dict) and isinstance(item.get("id"), str) and item["id"].strip():
            result.add(item["id"].strip())
    return sorted(result, key=str.casefold)


def load_catalog(path: str) -> dict:
    try:
        with open(path, encoding="utf-8") as fh:
            value = json.load(fh)
        if isinstance(value, dict) and isinstance(value.get("models"), list):
            return value
    except (OSError, ValueError, TypeError):
        pass
    return {"date": None, "models": []}


def sync_daily_catalog(
    path: str,
    api_key: str,
    *,
    endpoint: str = DEFAULT_ENDPOINT,
    today: Optional[str] = None,
    timeout: int = 20,
    opener: Callable = urlopen,
) -> Tuple[List[str], bool, Optional[str]]:
    """Fetch /models once per day and persist its model IDs.

    Returns ``(models, refreshed, warning)``. A failed refresh keeps the last
    successful catalog, so token reporting remains available offline.
    """
    today = today or date.today().isoformat()
    old = load_catalog(path)
    if old.get("date") == today and old.get("models"):
        return old["models"], False, None
    if not api_key:
        return old.get("models", []), False, "LiteLLM API key unavailable"
    try:
        req = Request(endpoint, headers={"Authorization": f"Bearer {api_key}"})
        with opener(req, timeout=timeout) as response:
            payload = json.load(response)
        models = extract_model_ids(payload)
        if not models:
            raise ValueError("LiteLLM /v1/models returned no model IDs")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump({"date": today, "models": models}, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return models, True, None
    except (OSError, ValueError, TypeError, HTTPError, URLError) as exc:
        return old.get("models", []), False, f"LiteLLM model sync failed: {exc}"


def load_api_key_from_hermes_config(path: str) -> str:
    """Read the custom provider key without importing YAML or printing it."""
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                if "api_key:" not in line:
                    continue
                value = line.split("api_key:", 1)[1].strip().strip("'\"")
                if value and not value.startswith("#"):
                    return value
    except OSError:
        pass
    return os.environ.get("LITELLM_API_KEY", "")
