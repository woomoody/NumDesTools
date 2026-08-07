import json
from io import BytesIO

from token_model_catalog import extract_model_ids, sync_daily_catalog


def test_extract_model_ids_deduplicates_and_sorts():
    assert extract_model_ids({"data": [{"id": "qwen3.8-max"}, {"id": "glm-5.2"}, {"id": "qwen3.8-max"}, {"id": ""}]}) == ["glm-5.2", "qwen3.8-max"]


def test_sync_daily_catalog_refreshes_once_and_persists_gateway_models(tmp_path):
    catalog = tmp_path / "model_catalog.json"
    response = BytesIO(json.dumps({"data": [{"id": "qwen3.8-max"}, {"id": "gpt-5.6-luna"}]}).encode())
    response.__enter__ = lambda: response
    response.__exit__ = lambda *args: None

    def opener(request, timeout):
        assert request.get_header("Authorization") == "Bearer secret"
        assert timeout == 20
        return response

    models, refreshed, warning = sync_daily_catalog(str(catalog), "secret", today="2026-08-07", opener=opener)
    assert models == ["gpt-5.6-luna", "qwen3.8-max"]
    assert refreshed is True
    assert warning is None
    assert json.loads(catalog.read_text(encoding="utf-8"))["models"] == models


def test_sync_daily_catalog_does_not_call_gateway_twice_same_day(tmp_path):
    catalog = tmp_path / "model_catalog.json"
    catalog.write_text(json.dumps({"date": "2026-08-07", "models": ["qwen3.8-max"]}), encoding="utf-8")

    def should_not_call(*args, **kwargs):
        raise AssertionError("gateway should only be queried once per day")

    models, refreshed, warning = sync_daily_catalog(str(catalog), "secret", today="2026-08-07", opener=should_not_call)
    assert models == ["qwen3.8-max"]
    assert refreshed is False
    assert warning is None


def test_sync_daily_catalog_keeps_last_catalog_when_gateway_fails(tmp_path):
    catalog = tmp_path / "model_catalog.json"
    catalog.write_text(json.dumps({"date": "2026-08-06", "models": ["qwen3.8-max"]}), encoding="utf-8")

    def failing_opener(*args, **kwargs):
        raise OSError("network down")

    models, refreshed, warning = sync_daily_catalog(str(catalog), "secret", today="2026-08-07", opener=failing_opener)
    assert models == ["qwen3.8-max"]
    assert refreshed is False
    assert "sync failed" in warning
