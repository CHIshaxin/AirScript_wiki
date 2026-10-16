"""HTTP 客户端测试：URL 构造、退避重试、哨兵识别限流。"""

from airscript_wiki import site
from airscript_wiki.http import Fetcher


def test_asset_url_builds_cdn_path():
    url = site.asset_url("apiV2/excel/workbook/Range/属性/Count 属性.md", "abc123")
    assert url == (
        "https://qn.cache.wpscdn.cn/asdocs/assets/"
        "apiV2_excel_workbook_Range_%E5%B1%9E%E6%80%A7_Count%20%E5%B1%9E%E6%80%A7.md.abc123.js"
    )


def test_page_url_uses_html_suffix():
    assert site.page_url("apiV2/excel/workbook/Range/Range 对象.md") == (
        "https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/Range%20%E5%AF%B9%E8%B1%A1.html"
    )


def test_get_retries_then_succeeds():
    calls = []

    def fake_get(url):
        calls.append(url)
        return (404, b"") if len(calls) == 1 else (200, b"ok")

    fetcher = Fetcher(delay=0, get=fake_get, sleep=lambda s: None, sentinel_url=None)
    status, body = fetcher.get("https://example.com/a.js")
    assert (status, body) == (200, b"ok")
    assert len(calls) == 2
    assert "cb=" in calls[1]  # 重试带 cache-busting，绕过负缓存
    assert fetcher.stats.retries == 1


def test_probe_with_zero_retries_does_not_retry():
    calls = []

    def fake_get(url):
        calls.append(url)
        return 404, b""

    fetcher = Fetcher(delay=0, get=fake_get, sleep=lambda s: None, sentinel_url=None)
    status, _ = fetcher.get("https://example.com/missing.js", retries=0)
    assert status == 404
    assert len(calls) == 1


def test_sentinel_404_triggers_pause_then_recovers():
    slept = []
    state = {"sentinel_hits": 0}

    def fake_get(url):
        if "sentinel" in url:
            state["sentinel_hits"] += 1
            return (404, b"") if state["sentinel_hits"] == 1 else (200, b"")
        return (404, b"")

    fetcher = Fetcher(delay=0, get=fake_get, sleep=slept.append,
                      sentinel_url="https://example.com/sentinel.js",
                      max_throttle_waits=2)
    assert fetcher.check_sentinel() is True
    assert slept and slept[0] >= 60  # 限流时全局暂停
    assert fetcher.stats.throttle_pauses == 1


def test_sentinel_interval_limits_checks():
    probes = []

    def fake_get(url):
        probes.append(url)
        return 200, b""

    fetcher = Fetcher(delay=0, get=fake_get, sleep=lambda s: None,
                      sentinel_url="https://example.com/sentinel.js", sentinel_interval=1000)
    fetcher.throttled()
    fetcher.throttled()
    fetcher.throttled()
    assert len(probes) == 1  # 同一时间窗内只探测一次，避免自身放大请求量
