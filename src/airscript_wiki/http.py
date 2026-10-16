"""带节流、哨兵与退避重试的 HTTP 客户端。

站点 CDN 在短时间内收到大量请求（尤其是大量 404）后会进入数分钟的“全站 404”状态。
因此任何 404 都必须先区分“页面真的不存在”和“我们被临时限流了”，再决定是否记为失败。
"""

from __future__ import annotations

import sys
import time
from dataclasses import dataclass, field

import requests

from . import site

THROTTLE_WAITS = (60.0, 120.0, 300.0)


@dataclass
class Stats:
    requests: int = 0
    retries: int = 0
    throttle_pauses: int = 0
    bytes: int = 0


class Fetcher:
    """``get`` 返回 ``(status, body)``；404 只有在退避重试后仍失败才返回。"""

    def __init__(self, delay: float = 0.5, max_retries: int = 3, get=None,
                 sleep=time.sleep, sentinel_url: str | None = None, log=None,
                 max_throttle_waits: int = 3, sentinel_interval: float = 20.0):
        self.delay = delay
        self.max_retries = max_retries
        self._get = get or self._requests_get
        self._sleep = sleep
        self.sentinel_url = sentinel_url
        self.max_throttle_waits = max_throttle_waits
        self.sentinel_interval = sentinel_interval
        self.stats = Stats()
        self._last_request = 0.0
        self._last_sentinel_check = 0.0
        self._last_sentinel_ok = True
        self._log = log or (lambda msg: print(msg, file=sys.stderr, flush=True))

    # ---- 底层 ----
    def _requests_get(self, url: str) -> tuple[int, bytes]:
        resp = requests.get(url, headers={"User-Agent": site.USER_AGENT}, timeout=45)
        return resp.status_code, resp.content

    def _throttle(self):
        if self.delay <= 0:
            return
        wait = self.delay - (time.monotonic() - self._last_request)
        if wait > 0:
            self._sleep(wait)

    def _one(self, url: str) -> tuple[int, bytes]:
        self._throttle()
        self._last_request = time.monotonic()
        self.stats.requests += 1
        status, body = self._get(url)
        self.stats.bytes += len(body or b"")
        return status, body

    # ---- 哨兵：识别“被限流” ----
    def check_sentinel(self) -> bool:
        """探测哨兵 URL；若也 404，说明处于限流窗口，暂停后退避重试。"""
        if not self.sentinel_url:
            return True
        for attempt, wait in enumerate(THROTTLE_WAITS[: self.max_throttle_waits + 1]):
            status, _ = self._one(self.sentinel_url)
            if status == 200:
                return True
            if attempt >= self.max_throttle_waits:
                break
            self.stats.throttle_pauses += 1
            self._log(f"[throttle] sentinel {status} -> pause {wait:.0f}s")
            self._sleep(wait)
        return False

    # ---- 对外 ----
    def throttled(self) -> bool:
        """当前是否处于限流窗口（哨兵探测按 ``sentinel_interval`` 限频，避免自身放大请求量）。"""
        if not self.sentinel_url:
            return False
        now = time.monotonic()
        if now - self._last_sentinel_check < self.sentinel_interval:
            return not self._last_sentinel_ok
        self._last_sentinel_check = now
        self._last_sentinel_ok = self.check_sentinel()
        return not self._last_sentinel_ok

    def get(self, url: str, retries: int | None = None) -> tuple[int, bytes]:
        """GET 一个资源；404 时先用 cache-busting 退避重试，再确认是否限流。

        探测真实路径的调用（``retries=0``）不应重试：那里的 404 是预期结果。
        """
        attempts = self.max_retries if retries is None else retries
        backoff = (1.0, 3.0, 8.0)
        for i in range(attempts + 1):
            bust = "" if i == 0 else ("&" if "?" in url else "?") + f"cb={int(time.time() * 1000)}-{i}"
            status, body = self._one(url + bust)
            if status == 200:
                return status, body
            if status in (404, 403, 429, 500, 502, 503, 504):
                # 已缓存的 404 可能是限流期间的产物：重试前先确认是否被限流
                if self.throttled():
                    if not self.check_sentinel():
                        return status, body
                if i < attempts:
                    self.stats.retries += 1
                    self._sleep(backoff[min(i, len(backoff) - 1)])
                continue
            return status, body
        return 404, b""

    def get_text(self, url: str, retries: int | None = None) -> str:
        status, body = self.get(url, retries=retries)
        if status != 200:
            raise OSError(f"GET {url} -> {status}")
        return body.decode("utf-8", errors="replace")

    def find_sentinel(self) -> str | None:
        """从站点外壳里挑一个确定存在的资源作为哨兵（逐个验证直到 200）。"""
        import re
        try:
            shell = self.get_text(site.SITE_ORIGIN + site.BASE)
        except OSError:
            return None
        for m in re.finditer(r'assets/([\w./-]+\.js)', shell):
            candidate = site.CDN + "assets/" + m.group(1)
            status, _ = self._one(candidate)
            if status == 200:
                self.sentinel_url = candidate
                return candidate
        return None
