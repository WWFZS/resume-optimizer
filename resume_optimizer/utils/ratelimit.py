# -*- coding: utf-8 -*-
"""轻量进程内限流。

单机部署够用；多 worker / 多实例部署时应替换为 Redis 实现。
窗口算法：每个 key 维护一个时间戳队列，滑动窗口内计数超限即拒绝。
"""
import threading
import time
from collections import defaultdict, deque

_LOCK = threading.Lock()
_HITS = defaultdict(deque)


def _hit(key: str, limit: int, window: float):
    """记一次访问。返回 (allowed, retry_after_seconds)。"""
    now = time.time()
    with _LOCK:
        q = _HITS[key]
        while q and now - q[0] > window:
            q.popleft()
        if len(q) >= limit:
            return False, max(1, int(window - (now - q[0])) + 1)
        q.append(now)
        # 惰性清理，避免长期运行字典无限增长
        if len(_HITS) > 5000:
            for k in [x for x, dq in _HITS.items() if not dq or now - dq[-1] > window]:
                _HITS.pop(k, None)
        return True, 0


def _is_private_host(host: str) -> bool:
    try:
        import ipaddress
        ip = ipaddress.ip_address((host or "").split("%")[0])
        return ip.is_private or ip.is_loopback
    except ValueError:
        return False


def client_ip(request) -> str:
    """取客户端 IP；仅当直连方是本机/内网（反向代理）时才信任 X-Forwarded-For。"""
    try:
        host = request.client.host if request.client else "unknown"
        xff = (request.headers.get("x-forwarded-for") or "").strip()
        if xff and _is_private_host(host):
            first = xff.split(",")[0].strip()
            if first:
                return first
        return host or "unknown"
    except Exception:  # noqa: BLE001
        return "unknown"


def check(request, scope: str, key: str = "", limit: int = 10, window: float = 300):
    """超限时抛 429（带 Retry-After）。key 为空时按 IP 限流。"""
    ident = f"{client_ip(request)}|{key}" if key else client_ip(request)
    allowed, retry = _hit(f"{scope}:{ident}", limit, window)
    if not allowed:
        from fastapi import HTTPException
        raise HTTPException(
            status_code=429,
            detail=f"请求过于频繁，请 {retry} 秒后再试",
            headers={"Retry-After": str(retry)},
        )
