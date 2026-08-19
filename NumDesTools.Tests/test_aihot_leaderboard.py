import json
from io import BytesIO

from aihot_leaderboard import find_reference, parse_leaderboard_html, sync_leaderboard


ROW = '''<a class="lb-row" href="/leaderboard/gpt-5-6-sol"><span class="lb-rank"><b>01</b></span><span class="lb-model"><span><strong>GPT-5.6 Sol</strong><small>OpenAI</small></span></span><span class="lb-release-date"><strong>2026-07-09</strong></span><span class="lb-completeness"><strong>75.5%</strong></span><span class="lb-input-price" aria-label="输入 ¥33.71。OpenAI API 官网，官方型号 gpt-5.6-sol，核验于 2026-08-10"><strong>¥33.71</strong></span><span class="lb-output-price" aria-label="输出 ¥202.27。OpenAI API 官网，官方型号 gpt-5.6-sol，核验于 2026-08-10"><strong>¥202.27</strong></span><span class="lb-score"><strong>85.2</strong></span></a>'''


def test_parse_row_preserves_prices_and_provenance():
    rows = parse_leaderboard_html(ROW)
    assert len(rows) == 1
    assert rows[0]["slug"] == "gpt-5-6-sol"
    assert rows[0]["name"] == "GPT-5.6 Sol"
    assert rows[0]["input_cny_per_million"] == 33.71
    assert rows[0]["output_cny_per_million"] == 202.27
    assert rows[0]["score"] == 85.2
    assert "OpenAI API" in rows[0]["pricing_source"]


def test_find_reference_normalizes_harness_prefix():
    rows = parse_leaderboard_html(ROW)
    assert find_reference(rows, "litellm/gpt-5-6-sol")["score"] == 85.2


def test_sync_is_daily_and_keeps_cache_when_network_fails(tmp_path):
    path = tmp_path / "aihot.json"
    response = BytesIO(ROW.encode())
    response.__enter__ = lambda: response
    response.__exit__ = lambda *args: None

    rows, refreshed, warning = sync_leaderboard(path, today="2026-08-19", opener=lambda req, timeout: response)
    assert refreshed and warning is None and rows[0]["name"] == "GPT-5.6 Sol"
    rows, refreshed, warning = sync_leaderboard(path, today="2026-08-19", opener=lambda *_: (_ for _ in ()).throw(AssertionError()))
    assert not refreshed and warning is None

    rows, refreshed, warning = sync_leaderboard(path, today="2026-08-20", opener=lambda *_: (_ for _ in ()).throw(OSError("offline")))
    assert not refreshed and rows[0]["score"] == 85.2 and "sync failed" in warning
