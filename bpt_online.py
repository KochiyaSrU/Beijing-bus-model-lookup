"""联网查询：连 bjbus.jinzihao.me 那台 MCP 服务（**只用标准库**）。

为什么要单独一个模块：
- 这一层只做「网络 + 解析」，**不 import tkinter**，所以可以脱离界面直接单测；
- 全用标准库（urllib / json / threading / queue），不引入新依赖，onefile 打包不变形；
- 所有纯解析函数（`parse_*` / `history_to_links`）**不碰网络**，测试直接喂真实回包夹具；
- 出网相关的失败一律抛 `OnlineError` 的子类，由界面决定怎么说人话。

服务端的现实（实测）：
- 它是 MCP over SSE：`GET /sse` 先回一条 `event: endpoint`，里面给 `/messages/?session_id=…`；
  之后 POST JSON-RPC 到那个地址，**响应从 SSE 流里回来**（不是 POST 的返回体）；
- 协议版本 2024-11-05，`serverInfo = My MCP Server 1.12.4`（FastMCP）；
- 内部 route id 只在"当前这份数据集"里有效、会变，所以**绝不落盘**，只存线路号与日期。
"""

from __future__ import annotations

import json
import os
import queue
import socket
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Protocol

from bpt_data import (
    LineLink,
    append_online_log,
    cache_entry,
    cache_prune,
    cache_put,
    load_online_cache,
    month_key,
    parse_arrivals,
    parse_depots,
    parse_journey_time,
    parse_routes_at_stop,
    parse_routes_between,
    parse_route_list,
    parse_schedules,
    parse_timetable,
    parse_track_latest,
    save_online_cache,
)

#: 实时数据（车在哪、到站、时刻表）只做这么久的**内存**缓存，绝不落盘
LIVE_CACHE_SECONDS = 30.0

DEFAULT_BASE_URL = "https://bjbus.jinzihao.me"
OFFLINE_ENV = "BPT_NO_NETWORK"
PROTOCOL_VERSION = "2024-11-05"
USER_AGENT = "BPT-BusLookup/1.0 (desktop; Beijing bus model lookup)"  # HTTP 头只能 latin-1，别写中文

#: 在线数据最多"新鲜"多久就算还在跑（天）。查不到 ≠ 已退出，所以只会用来把 recorded 提成 current。
CURRENT_WITHIN_DAYS = 7


class OnlineError(Exception):
    """联网相关的失败基类。"""


class OnlineUnavailable(OnlineError):
    """没网、或用户把联网关了、或对方连不上。"""


class OnlineProtocolError(OnlineError):
    """连上了，但对方回的东西看不懂（握手失败、缺字段、不是 JSON）。"""


@dataclass(frozen=True)
class OnlineSettings:
    """联网设置（键都放在 界面设置.json 的 network 段里，缺项用默认值）。"""

    enabled: bool = True
    base_url: str = DEFAULT_BASE_URL
    timeout_sec: int = 45          # 单次响应上限：历史类查询会慢
    connect_timeout_sec: int = 8   # 建连上限
    cache_days: int = 7            # 结果缓存多久
    min_interval_sec: float = 2.0  # 两次请求之间至少隔多久（对第三方客气一点）

    @classmethod
    def from_dict(cls, payload: Any) -> "OnlineSettings":
        data = payload if isinstance(payload, dict) else {}

        def number(key: str, fallback: float) -> float:
            try:
                return float(data.get(key, fallback))
            except (TypeError, ValueError):
                return fallback

        base = str(data.get("base_url", "") or "").strip() or DEFAULT_BASE_URL
        return cls(
            enabled=bool(data.get("enabled", True)),
            base_url=base.rstrip("/"),
            timeout_sec=max(5, int(number("timeout_sec", 45))),
            connect_timeout_sec=max(3, int(number("connect_timeout_sec", 8))),
            cache_days=max(0, int(number("cache_days", 7))),
            min_interval_sec=max(0.0, number("min_interval_sec", 2.0)),
        )

    @property
    def usable(self) -> bool:
        """现在能不能联网（开关 + 环境变量，两条任一条说不行就不行）。"""
        return self.enabled and not offline_forced()


def offline_forced() -> bool:
    """`BPT_NO_NETWORK=1` 时强制离线（给自己测"断网还能不能用"，也给用户一个彻底关掉的开关）。"""
    raw = os.environ.get(OFFLINE_ENV, "").strip().lower()
    return raw not in ("", "0", "false", "no")


# --------------------------------------------------------------------------- #
# 纯解析：不碰网络，测试直接喂夹具
# --------------------------------------------------------------------------- #


def parse_date_range(text: str) -> tuple[str, str] | None:
    """`20200815-20201012` → `("20200815", "20201012")`；认不出返回 None。"""
    parts = str(text).strip().split("-")
    if len(parts) != 2:
        return None
    start, end = parts[0].strip(), parts[1].strip()
    if not (start.isdigit() and end.isdigit() and len(start) == 8 and len(end) == 8):
        return None
    return start, end


def compact_to_month(text: str) -> str:
    """`20200815` → `2020.8`（程序里 first/last 一直用「年.月」这种写法）。"""
    value = str(text).strip()
    if len(value) != 8 or not value.isdigit():
        return value
    return f"{value[:4]}.{int(value[4:6])}"


def parse_history_map(payload: Any) -> dict[str, list[tuple[str, str]]]:
    """运营历史的两头都长这样（线路→日期区间 或 车辆→日期区间），共用这一个解析。

    顺手把区间按开始日期排好、去掉重复和认不出的。
    """
    if not isinstance(payload, dict):
        raise OnlineProtocolError("运营历史的返回不是一个字典")
    result: dict[str, list[tuple[str, str]]] = {}
    for key, ranges in payload.items():
        if not isinstance(ranges, list):
            continue
        collected: list[tuple[str, str]] = []
        for item in ranges:
            parsed = parse_date_range(str(item))
            if parsed and parsed not in collected:
                collected.append(parsed)
        if collected:
            result[str(key).strip()] = sorted(collected)
    return result


def parse_route_ids(payload: Any) -> list[tuple[str, str, str]]:
    """`query_route_id_by_route_number` 的返回 → [(内部id, 首站, 末站)]（内部 id 不落盘）。

    注意：**内部 id 不是线路号**，一个线路号对应两个方向、两个 id（实测每对都差 1：
    886→805/806、1路→5918/5919、919路→5299/5300），而且服务端声明它会随数据集变化，
    所以只当"本次会话里的指针"用，绝不能说"886 就是 805"。

    查不到的线路，服务端有时回的不是列表（会带 route_not_found 之类），这里按"没查到"处理，
    并把原文带上，免得界面报一句"返回不是一个列表"这种术语。
    """
    if isinstance(payload, dict):
        if "result" in payload:
            payload = payload.get("result")
        elif payload.get("error") or payload.get("detail"):
            raise OnlineProtocolError(f"服务端说这条线路查不到：{payload.get('error') or payload.get('detail')}")
    if not isinstance(payload, list):
        raise OnlineProtocolError(f"服务端没有给出方向（返回是 {type(payload).__name__}）")
    rows: list[tuple[str, str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        rows.append(
            (
                str(item.get("id", "")).strip(),
                str(item.get("first_stop", "")).strip(),
                str(item.get("last_stop", "")).strip(),
            )
        )
    return rows


def parse_stops(payload: Any) -> list[tuple[str, str, str, str, str]]:
    """`query_route_info_by_route_id` → [(序号, 站名, 距上一站米, 经度, 纬度)]。"""
    if not isinstance(payload, dict):
        raise OnlineProtocolError("线路详情的返回不是一个字典")
    stops = payload.get("stops")
    if not isinstance(stops, list):
        return []
    rows: list[tuple[str, str, str, str, str]] = []
    for index, item in enumerate(stops, start=1):
        if not isinstance(item, dict):
            continue
        # 注意别用 `value or ""`：距离 0 是有效值，会被真值判断吃掉
        distance = item.get("distance_from_previous_stop")
        rows.append(
            (
                str(index),
                str(item.get("name", "")).strip(),
                "" if distance is None else str(distance),
                "" if item.get("longitude") is None else str(item.get("longitude")),
                "" if item.get("latitude") is None else str(item.get("latitude")),
            )
        )
    return rows


def history_to_links(
    bus_id: str,
    history: dict[str, list[tuple[str, str]]],
    model_of: Callable[[str], str] | None = None,
    *,
    checked: str = "",
    current_within_days: int = CURRENT_WITHIN_DAYS,
    today: str = "",
) -> list[LineLink]:
    """把「一辆车的运营历史」变成线路记录。

    - `model_of` 用号段表反查车型；查不到就留空（界面会标「车型未知」）；
    - 最新区间结束日距今 ≤ `current_within_days` 天才算「现在还在跑」；
      否则一律 `recorded` —— **永不写 `retired`**，因为"查不到"不能证明车退出了；
    - `checked` 是查询日期（YYYYMMDD），写进 note 里，让人知道这是哪天的数据。
    """
    number = str(bus_id).strip()
    model = (model_of(number) if model_of else "") or ""
    note = f"网上查的（{_human_date(checked) if checked else '第三方实时数据'}）"
    links: list[LineLink] = []
    for line, ranges in history.items():
        # 容错：调用方可能已经解析过（元组），也可能直接喂原始 "20200815-20201012"
        cleaned: list[tuple[str, str]] = []
        for item in ranges:
            if isinstance(item, (tuple, list)) and len(item) == 2:
                cleaned.append((str(item[0]), str(item[1])))
            else:
                parsed = parse_date_range(str(item))
                if parsed:
                    cleaned.append(parsed)
        if not cleaned:
            continue
        starts = [start for start, _end in cleaned]
        ends = [end for _start, end in cleaned]
        latest = max(ends)
        status = "recorded"
        if _within_days(latest, today or _today(), current_within_days):
            status = "current"
        links.append(
            LineLink(
                line=line,
                model=model,
                numbers=(number,) if number else (),
                first=compact_to_month(min(starts)),
                last=compact_to_month(latest),
                status=status,
                source="online",
                note=note,
            )
        )
    return sorted(links, key=lambda item: (month_key(item.first), item.line))


def _today() -> str:
    return date.today().strftime("%Y%m%d")


def _human_date(value: str) -> str:
    parsed = datetime.strptime(value, "%Y%m%d") if len(value) == 8 and value.isdigit() else None
    return parsed.strftime("%Y-%m-%d") if parsed else str(value)


def _within_days(stamp: str, today: str, days: int) -> bool:
    try:
        end = datetime.strptime(stamp, "%Y%m%d").date()
        now = datetime.strptime(today, "%Y%m%d").date()
    except ValueError:
        return False
    return (now - end).days <= days


def model_lookup(ranges: list, number: str) -> str:
    """用号段表反查某个自编号的车型；查不到返回空串（界面会标「车型未知」）。"""
    from bpt_data import find_range  # 局部导入：这个模块只在这一个地方依赖号段查询

    try:
        found = find_range(int(str(number).strip()), ranges)
    except (TypeError, ValueError):
        return ""
    return found.model if found is not None else ""


def today_stamp() -> str:
    """今天的 YYYYMMDD（查询日期写进记录里，让人知道数据是哪天的）。"""
    return _today()


def error_text(error: BaseException) -> str:
    """把联网异常翻成一句人话（界面不弹红框，只在窗口里说明）。"""
    if isinstance(error, OnlineUnavailable):
        text = str(error)
        hint = _network_hint(text)
        return f"{text}{hint}。联网功能是可选的，本地数据完全不受影响；过一会儿点「查询」重试即可。"
    if isinstance(error, OnlineProtocolError):
        return f"对方返回的内容看不懂：{error}"
    if isinstance(error, OnlineError):
        return str(error)
    return f"{type(error).__name__}: {error}"


def short_error(error: BaseException) -> str:
    """一句话说清是哪类联网问题（主界面状态栏只有一行，塞不下长句子）。"""
    if isinstance(error, OnlineProtocolError):
        return "对方返回看不懂"
    text = str(error)
    lowered = text.casefold()
    if "10061" in text or "refused" in lowered or "积极拒绝" in text:
        return "连接被拒绝"
    if "timed out" in lowered or "超时" in text or "没有回应" in text or "等了太久" in text:
        return "等超时了"
    if "getaddrinfo" in lowered or "nodename" in lowered:
        return "域名解析不了"
    if "断了" in text:
        return "连接断了"
    return "连不上"


def _network_hint(text: str) -> str:
    """针对最常见的两种网络错误补一句"那大概是什么原因"。

    10061（连接被拒）：服务端没在听、或本机网络/防火墙/VPN 把它拦了；
    超时：链路慢或被墙。实测这类失败经常只是一下子，重试就好——所以要写清"可以重试"。
    """
    lowered = text.casefold()
    if "10061" in text or "refused" in lowered or "积极拒绝" in text:
        return "（连接被拒绝：可能是服务端一时没在听，或者本机网络/防火墙/VPN 拦了它）"
    if "timed out" in lowered or "超时" in text or "没有回应" in text or "等了太久" in text:
        return "（等超时了：网络慢或被拦）"
    if "getaddrinfo" in lowered or "名称" in text or "nodename" in lowered:
        return "（域名解析不了：检查本机 DNS / 网络）"
    return ""



def unpack_tool_result(result: Any) -> Any:
    """把 MCP 的工具返回拆成真正的数据。

    FastMCP 会把结果塞在 `content[0].text` 里（JSON 字符串）；能走 structuredContent 就走它。
    两条路都没有就抛协议错误——**绝不猜**。
    """
    if not isinstance(result, dict):
        raise OnlineProtocolError("工具返回不是一个对象")
    structured = result.get("structuredContent")
    if isinstance(structured, dict) and structured:
        if set(structured.keys()) == {"result"}:
            return structured["result"]
        return structured
    for item in result.get("content") or []:
        if not isinstance(item, dict):
            continue
        text = item.get("text")
        if not isinstance(text, str):
            continue
        stripped = text.strip()
        if not stripped:
            continue
        try:
            return json.loads(stripped)
        except json.JSONDecodeError:
            return stripped
    raise OnlineProtocolError("工具返回里没有可用的内容")


# --------------------------------------------------------------------------- #
# 传输层
# --------------------------------------------------------------------------- #


class Transport(Protocol):
    """给 `MCPClient` 用的最小传输接口（测试里塞假的就行）。"""

    def open(self) -> None: ...

    def post(self, payload: dict) -> None: ...

    def events(self) -> "queue.Queue[dict]": ...

    def close(self) -> None: ...


class SSETransport:
    """按 MCP 的 SSE 规矩连服务：GET /sse 拿会话地址，POST 提问，响应从流里读。

    关于超时：`urlopen(timeout=…)` 同时对"建连"和"读"生效，而这两件事要的时间差很远——
    建连要快（连不上就别让用户等），读流要久（服务端可能十几秒才回一次，中途还长时间静默）。
    所以先用 `socket.create_connection` 短超时探一下通不通，再用长超时去读流。
    """

    def __init__(
        self,
        base_url: str = DEFAULT_BASE_URL,
        *,
        connect_timeout: int = 8,
        read_timeout: int = 180,
        session_timeout: int = 30,
        post_timeout: int = 30,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.connect_timeout = connect_timeout
        self.read_timeout = read_timeout
        # 等会话地址、递请求分别用这两个独立上限：服务端建会话实测 ~6.5 秒，
        # 拿 connect_timeout（8s）去顶只剩 1.5 秒余量，会偶发失败。
        self.session_timeout = session_timeout
        self.post_timeout = post_timeout
        self._queue: "queue.Queue[dict]" = queue.Queue()
        self._stream = None
        self._reader: threading.Thread | None = None
        self._session_url = ""

    # -- 生命周期 --------------------------------------------------------- #
    def _preflight(self) -> None:
        """短超时探通：不通就立刻说"连不上"，不让用户干等长超时。"""
        parts = urllib.parse.urlsplit(self.base_url)
        host = parts.hostname or ""
        port = parts.port or (443 if parts.scheme == "https" else 80)
        if not host:
            raise OnlineUnavailable("联网地址不对")
        try:
            with socket.create_connection((host, port), timeout=self.connect_timeout):
                return
        except OSError as error:
            raise OnlineUnavailable(f"连不上联网服务（{error}）") from error

    def open(self) -> None:
        self._preflight()
        request = urllib.request.Request(
            f"{self.base_url}/sse",
            headers={
                "Accept": "text/event-stream",
                "Cache-Control": "no-cache",
                "User-Agent": USER_AGENT,
            },
        )
        try:
            self._stream = urllib.request.urlopen(request, timeout=self.read_timeout)
        except (urllib.error.URLError, OSError, TimeoutError) as error:
            raise OnlineUnavailable(f"连不上联网服务（{error}）") from error
        self._reader = threading.Thread(target=self._read_loop, name="bpt-online-sse", daemon=True)
        self._reader.start()
        # 等服务端给会话地址：实测要 6.5 秒，所以不能用 8 秒的"建连超时"（那样只剩 1.5 秒余量，
        # 网络稍慢就报"没有给出会话地址"）。用 session_timeout。
        deadline = time.time() + self.session_timeout
        while time.time() < deadline:
            if self._session_url:
                return
            time.sleep(0.05)
        raise OnlineProtocolError("联网服务没有给出会话地址")

    def post(self, payload: dict) -> None:
        if not self._session_url:
            raise OnlineProtocolError("会话还没建好")
        request = urllib.request.Request(
            self._session_url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json, text/event-stream",
                "User-Agent": USER_AGENT,
            },
            method="POST",
        )
        try:
            # POST 只是把请求递上去（真正的回包从 SSE 流里来），给它 post_timeout；
            # 用 8 秒的建连超时同样偏紧。
            with urllib.request.urlopen(request, timeout=self.post_timeout) as response:
                response.read()
        except (urllib.error.URLError, OSError, TimeoutError) as error:
            raise OnlineUnavailable(f"发不出去（{error}）") from error

    def events(self) -> "queue.Queue[dict]":
        return self._queue

    def close(self) -> None:
        stream, self._stream = self._stream, None
        if stream is not None:
            try:
                stream.close()
            except OSError:
                pass

    # -- 读流 ------------------------------------------------------------- #
    def _read_loop(self) -> None:
        stream = self._stream
        if stream is None:
            return
        try:
            for raw in stream:
                line = raw.decode("utf-8", "replace").rstrip("\r\n")
                if not line.startswith("data:"):
                    continue
                data = line[5:].strip()
                if not data:
                    continue
                if "session_id=" in data and not data.startswith("{"):
                    self._session_url = data if data.startswith("http") else f"{self.base_url}{data}"
                    continue
                try:
                    self._queue.put(json.loads(data))
                except json.JSONDecodeError:
                    continue
        except (OSError, ValueError):
            pass
        finally:
            self._queue.put({"__closed__": True})


class OfflineTransport:
    """离线替身：给"用户关了联网 / 没网"这条路用，也让测试不碰真网络。"""

    def open(self) -> None:
        raise OnlineUnavailable("联网已关闭")

    def post(self, payload: dict) -> None:
        raise OnlineUnavailable("联网已关闭")

    def events(self) -> "queue.Queue[dict]":
        return queue.Queue()

    def close(self) -> None:
        pass


# --------------------------------------------------------------------------- #
# 客户端
# --------------------------------------------------------------------------- #


class OnlineService:
    """全程序共用的联网服务：**一个会话 + 结果缓存**。

    为什么要有这一层（实测数据说话）：
    - 服务端建立会话要约 26 秒（SSE 6.5s + initialize 19s），**每开一个窗口都付一次太亏**；
    - 服务端每个工具调用固定约 6.4 秒（复用会话也一样），所以**缓存能把重复查询变成 0 秒**。

    所以：会话懒建、全程序复用；查询结果按 (工具, 参数) 缓存到 `联网缓存.json`，默认 7 天。
    线程安全：内部串行化（一个 SSE 会话同时只能处理一个请求）。
    """

    def __init__(
        self,
        settings: OnlineSettings,
        base_directory: Path | str,
        *,
        client_factory: Callable[..., MCPClient] | None = None,
    ) -> None:
        self.settings = settings
        self.base_directory = Path(base_directory)
        self._factory = client_factory or MCPClient
        self._client: MCPClient | None = None
        self._lock = threading.RLock()
        self.cache: dict = {}
        self._live: dict[str, tuple[Any, str, float]] = {}   # 实时数据的 30 秒内存缓存
        self.last_source = ""      # "cache" / "net"
        self.last_seconds = 0.0
        self.load_cache()

    # -- 缓存 ------------------------------------------------------------- #
    def load_cache(self) -> None:
        with self._lock:
            self.cache = load_online_cache(self.base_directory)

    def save_cache(self) -> None:
        with self._lock:
            try:
                save_online_cache(self.base_directory, self.cache)
            except Exception:  # noqa: BLE001 - 缓存写不进去不该影响查询
                pass

    def clear_cache(self) -> int:
        with self._lock:
            count = len(self.cache)
            self.cache = {}
            self.save_cache()
        return count

    @property
    def cache_size(self) -> int:
        return len(self.cache)

    def _cached(self, key: str) -> Any:
        item = cache_entry(self.cache, key, days=self.settings.cache_days)
        if item is None:
            return None
        self.last_source = "cache"
        self.last_seconds = 0.0
        return item.get("payload"), str(item.get("checked", ""))

    def _store(self, key: str, payload: Any, checked: str) -> None:
        with self._lock:
            self.cache = cache_put(cache_prune(self.cache, days=self.settings.cache_days), key, payload, checked=checked)
            self.save_cache()

    # -- 会话 ------------------------------------------------------------- #
    def client(self) -> MCPClient:
        with self._lock:
            if self._client is None:
                self._client = self._factory(settings=self.settings)
            return self._client

    def warm_up(self) -> None:
        """后台先把会话建起来（用户还在输字的时候握手就跑完了）。"""
        if not self.settings.usable:
            return
        try:
            self.client().connect()
        except OnlineError:
            pass

    def close(self) -> None:
        with self._lock:
            client, self._client = self._client, None
        if client is not None:
            client.close()

    @property
    def connected(self) -> bool:
        return self._client is not None

    # -- 各类查询（带缓存） ----------------------------------------------- #
    def _call(self, kind: str, key: str, worker: Callable[[MCPClient], Any], checked: str = "",
              *, cache: bool = True) -> tuple[Any, str]:
        # 这一层也要查"能不能联网"：不能只指望客户端，否则换个 client 就漏了
        if not self.settings.usable:
            raise OnlineUnavailable("联网功能没有打开")
        cache_key = f"{kind}:{key}"
        if cache:
            hit = self._cached(cache_key)
            if hit is not None:
                return hit
        else:
            # 实时数据（车在哪、到站、时刻表）绝不能按天缓存：只做 30 秒内存缓存
            hit = self._live_cached(cache_key)
            if hit is not None:
                return hit
        started = time.perf_counter()
        last_error: OnlineError | None = None
        for attempt in (1, 2):
            try:
                payload = worker(self.client())
                break
            except OnlineError as error:
                last_error = error
                append_online_log(self.base_directory, f"{kind} {key}\t失败\t第{attempt}次：{error}")
                # 连接类失败（被拒/超时/断了）常常只是一下子——换一个会话重试一次，
                # 用户看到的就是"点一下查出来了"，而不是"又连不上"。协议错误不重试（重试也一样）。
                if attempt == 2 or isinstance(error, OnlineProtocolError):
                    raise
                self.close()  # 丢掉坏会话，下一轮重新握手
                time.sleep(1.0)
        if last_error is not None and "payload" not in locals():
            raise last_error
        self.last_source = "net"
        self.last_seconds = time.perf_counter() - started
        append_online_log(self.base_directory, f"{kind} {key}\t成功\t{self.last_seconds:.1f}s")
        stamp = checked or today_stamp()
        if cache:
            self._store(cache_key, payload, stamp)
        else:
            with self._lock:
                self._live[cache_key] = (payload, stamp, time.time())
        return payload, stamp

    def _live_cached(self, key: str) -> Any:
        """实时数据的 30 秒内存缓存（只活在本次运行里，绝不落盘）。"""
        with self._lock:
            item = self._live.get(key)
        if item is None:
            return None
        payload, stamp, at = item
        if time.time() - at > LIVE_CACHE_SECONDS:
            return None
        self.last_source = "cache"
        self.last_seconds = 0.0
        return payload, stamp

    def bus_history(self, bus_id: str) -> tuple[dict, str]:
        """自编号 → (线路: [日期区间], 数据日期)。"""
        return self._call("bus", str(bus_id).strip(), lambda client: client.query_bus_history(bus_id))

    def route_ids(self, route_number: str) -> tuple[list, str]:
        """线路号 → ([(内部id, 首站, 末站)], 数据日期)。"""
        return self._call("route", str(route_number).strip(), lambda client: client.query_route_ids(route_number))

    def route_history(self, route_number: str) -> tuple[dict, str]:
        """线路号 → (车辆: [日期区间], 数据日期)。"""
        return self._call("rh", str(route_number).strip(), lambda client: client.query_route_history(route_number))

    def buses_on_route(self, route_id: str) -> tuple[list, str]:
        """内部 id → (在跑的车, 数据日期)。"""
        return self._call("buses", str(route_id).strip(), lambda client: client.query_buses_on_route(route_id))

    def route_stops(self, route_id: str) -> tuple[list, str]:
        """内部 id → (站点表, 数据日期)。"""
        return self._call("stops", str(route_id).strip(), lambda client: client.query_route_stops(route_id))

    # -- 实时类（不落盘缓存，只有 30 秒内存缓存） ------------------------- #
    def track_latest(self, bus_id: str) -> tuple[dict, str]:
        """某车现在跑到哪儿了（轨迹的最后一条）。"""
        payload, stamp = self._call(
            "track", str(bus_id).strip(), lambda client: client.query_track_record(bus_id), cache=False
        )
        return parse_track_latest(payload), stamp

    def arrivals(self, route_id: str, stop: str) -> tuple[list, str]:
        """某线在某站的到站记录。"""
        payload, stamp = self._call(
            "arrivals", f"{route_id}@{stop}", lambda client: client.query_arrivals(route_id, stop), cache=False
        )
        return parse_arrivals(payload), stamp

    def timetable(self, route_id: str) -> tuple[list, str]:
        """某线的时刻表概要（车号/始发/终点/站数）。"""
        payload, stamp = self._call(
            "timetable", str(route_id).strip(), lambda client: client.query_timetable(route_id), cache=False
        )
        return parse_timetable(payload if isinstance(payload, str) else str(payload)), stamp

    def routes_at_stop(self, stop: str) -> tuple[list, str]:
        """这个站有哪些线（线路拓扑变化慢，按天缓存没问题）。"""
        payload, stamp = self._call(
            "stop", str(stop).strip(), lambda client: client.query_routes_at_stop(stop)
        )
        return parse_routes_at_stop(payload), stamp

    def routes_between(self, from_stop: str, to_stop: str) -> tuple[list, str]:
        """从 A 站到 B 站有哪些直达线路。"""
        payload, stamp = self._call(
            "between", f"{from_stop}→{to_stop}",
            lambda client: client.query_routes_between(from_stop, to_stop),
        )
        return parse_routes_between(payload), stamp

    def schedules(self, route_id: str) -> tuple[list, str]:
        """某线每辆车的发车/到达时刻（当天的班次，按天缓存）。"""
        payload, stamp = self._call("schedules", str(route_id).strip(), lambda client: client.query_schedules(route_id))
        return parse_schedules(payload), stamp

    def depot(self, route_id: str) -> tuple[list, str]:
        """某线的出车记录（从哪个场站、几点出车）。"""
        payload, stamp = self._call("depot", str(route_id).strip(), lambda client: client.query_depot(route_id))
        return parse_depots(payload), stamp

    def journey_time(self, route_id: str, from_stop: str, to_stop: str) -> tuple[list, str]:
        """同一线路两个站之间每辆车花了多久（当天班次算出来的，按天缓存没问题）。"""
        payload, stamp = self._call(
            "journey", f"{route_id}:{from_stop}→{to_stop}",
            lambda client: client.query_journey_time(route_id, from_stop, to_stop),
        )
        return parse_journey_time(payload), stamp

    def route_list(self) -> tuple[dict, str]:
        """服务端全量线路目录 → ({内部id: (线路号, 首站, 末站)}, 数据日期)。

        这份目录**不进联网缓存**（168KB，塞进缓存文件太笨），由界面单独存进 线路目录.json，
        用来在本地判断"这个线路号服务端认不认"，省掉打错号白等 6 秒。
        """
        if not self.settings.usable:
            raise OnlineUnavailable("联网功能没有打开")
        started = time.perf_counter()
        payload = self.client().query_route_list()
        self.last_source = "net"
        self.last_seconds = time.perf_counter() - started
        stamp = today_stamp()
        append_online_log(self.base_directory, f"route_list\t成功\t{self.last_seconds:.1f}s")
        return parse_route_list(payload), stamp

    # -- 给界面看的一句话 ------------------------------------------------- #
    def source_text(self) -> str:
        if self.last_source == "cache":
            return "来自缓存"
        if self.last_source == "net":
            return f"联网查询（{self.last_seconds:.1f}s）"
        return ""

    def build_mark(self) -> str:
        """这个程序自己的构建时间（帮用户判断"我是不是在跑新版"）。"""
        import sys

        try:
            path = Path(sys.executable)
            if getattr(sys, "frozen", False) and path.is_file():
                stamp = datetime.fromtimestamp(path.stat().st_mtime)
                return f"程序构建 {stamp:%Y-%m-%d %H:%M}"
        except Exception:  # noqa: BLE001 - 只是给界面显示，失败就算了
            pass
        return ""


# --------------------------------------------------------------------------- #
# 全程序一个实例（窗口之间共享会话与缓存）
# --------------------------------------------------------------------------- #

_SERVICE: OnlineService | None = None
_SERVICE_LOCK = threading.Lock()


def shared_service(settings: OnlineSettings, base_directory: Path | str) -> OnlineService:
    """拿到（必要时创建）全程序共用的联网服务。"""
    global _SERVICE
    with _SERVICE_LOCK:
        if _SERVICE is None:
            _SERVICE = OnlineService(settings, base_directory)
        else:
            _SERVICE.settings = settings
        return _SERVICE


def reset_shared_service() -> None:
    """关掉并忘记共享服务（程序退出、或测试里换目录时用）。"""
    global _SERVICE
    with _SERVICE_LOCK:
        if _SERVICE is not None:
            _SERVICE.close()
        _SERVICE = None


class MCPClient:
    """MCP 客户端：握手 + 调工具（同步、带超时，供后台线程使用）。"""

    def __init__(
        self,
        transport: Transport | None = None,
        *,
        base_url: str = DEFAULT_BASE_URL,
        settings: OnlineSettings | None = None,
    ) -> None:
        self.settings = settings or OnlineSettings(base_url=base_url)
        self.transport: Transport = transport or SSETransport(
            self.settings.base_url,
            connect_timeout=self.settings.connect_timeout_sec,
            session_timeout=max(30, self.settings.timeout_sec),
            post_timeout=max(30, self.settings.timeout_sec),
        )
        self._next_id = 0
        self._lock = threading.Lock()
        # 一个 SSE 会话同时只处理一个请求：并发调用会互相抢流里的回包，所以串行化。
        # 服务端每个请求固定约 6 秒，想快就靠缓存（见 OnlineService）。
        self._call_lock = threading.Lock()
        self._last_call = 0.0
        self._opened = False
        self.last_call_seconds = 0.0

    # -- 生命周期 --------------------------------------------------------- #
    def __enter__(self) -> "MCPClient":
        self.connect()
        return self

    def __exit__(self, *_exc) -> None:
        self.close()

    def connect(self) -> None:
        if self._opened:
            return
        self.transport.open()
        # 注意：这里要等的是**服务端的回应**，实测 initialize 要约 19 秒（建会话 6.5s + 握手 19s），
        # 所以必须用 timeout_sec（默认 45s）。之前误用了 connect_timeout_sec（8s，那是"建连"的上限），
        # 结果握手慢一点就报"等了太久，对方没有回应"——偶发连不上就是这么来的。
        self._send({"jsonrpc": "2.0", "id": self._take_id(), "method": "initialize",
                    "params": {"protocolVersion": PROTOCOL_VERSION, "capabilities": {},
                               "clientInfo": {"name": "北京公交车型查询器", "version": "1.0"}}})
        self._wait_for(self._next_id, self.settings.timeout_sec)
        self._send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        self._opened = True

    def close(self) -> None:
        self._opened = False
        self.transport.close()

    # -- 请求 ------------------------------------------------------------- #
    def _take_id(self) -> int:
        with self._lock:
            self._next_id += 1
            return self._next_id

    def _send(self, payload: dict) -> None:
        self.transport.post(payload)

    def _throttle(self) -> None:
        gap = self.settings.min_interval_sec - (time.time() - self._last_call)
        if gap > 0:
            time.sleep(gap)
        self._last_call = time.time()

    def _wait_for(self, request_id: int, timeout: int | float) -> Any:
        deadline = time.time() + float(timeout)
        events = self.transport.events()
        while True:
            remaining = deadline - time.time()
            if remaining <= 0:
                raise OnlineUnavailable("等了太久，对方没有回应")
            try:
                event = events.get(timeout=min(remaining, 0.5))
            except queue.Empty:
                continue
            if event.get("__closed__"):
                raise OnlineUnavailable("联网服务的连接断了")
            if event.get("id") != request_id:
                continue
            if "error" in event:
                message = event["error"]
                if isinstance(message, dict):
                    message = message.get("message") or message.get("code") or message
                raise OnlineProtocolError(str(message))
            return event.get("result")

    def call_tool(self, name: str, arguments: dict | None = None, *, timeout: int | float | None = None) -> Any:
        """调一个 MCP 工具，返回拆好的数据。"""
        if not self.settings.usable:
            raise OnlineUnavailable("联网功能没有打开")
        with self._call_lock:
            self.connect()
            self._throttle()
            request_id = self._take_id()
            self._send({"jsonrpc": "2.0", "id": request_id, "method": "tools/call",
                        "params": {"name": name, "arguments": arguments or {}}})
            started = time.perf_counter()
            result = self._wait_for(request_id, timeout or self.settings.timeout_sec)
            self.last_call_seconds = time.perf_counter() - started
            return unpack_tool_result(result)

    # -- 具体工具（界面只用这几个） --------------------------------------- #
    def query_route_ids(self, route_number: str) -> list[tuple[str, str, str]]:
        """线路号 → [(内部id, 首站, 末站)]（两个方向各一条）。**内部 id 不要存**。"""
        return parse_route_ids(self.call_tool("query_route_id_by_route_number",
                                              {"route_number": str(route_number).strip()}))

    def query_bus_history(self, bus_id: str) -> dict[str, list[tuple[str, str]]]:
        """自编号 → {线路号: [日期区间]}（这辆车跑过哪些线）。"""
        return parse_history_map(self.call_tool("query_bus_operating_history",
                                                {"bus_id": str(bus_id).strip()}))

    def query_route_history(self, route_number: str) -> dict[str, list[tuple[str, str]]]:
        """线路号 → {车辆自编号: [日期区间]}（这条线用过哪些车）。"""
        return parse_history_map(self.call_tool("query_route_operating_history",
                                                {"route_number": str(route_number).strip()}))

    def query_route_stops(self, route_id: str) -> list[tuple[str, str, str, str, str]]:
        """内部 id → 站点表（首末站一次只给一个方向）。"""
        return parse_stops(self.call_tool("query_route_info_by_route_id", {"route_id": str(route_id).strip()}))

    def query_buses_on_route(self, route_id: str) -> list[str]:
        """内部 id → 现在在跑的车（**该接口实测有时不回包**，调用方要能降级）。"""
        payload = self.call_tool("query_bus_ids_by_route_id", {"route_id": str(route_id).strip()})
        if isinstance(payload, list):
            return [str(item).strip() for item in payload if str(item).strip()]
        if isinstance(payload, dict):
            values = payload.get("result") or []
            return [str(item).strip() for item in values if str(item).strip()]
        return []

    # -- 实时类（V5） ----------------------------------------------------- #
    def query_track_record(self, bus_id: str) -> Any:
        """某车的轨迹记录（原始列表，30 秒一条，很长）。"""
        return self.call_tool("query_bus_track_record_by_bus_id", {"bus_id": str(bus_id).strip()})

    def query_arrivals(self, route_id: str, stop: str) -> Any:
        """某线在某站的到站记录。"""
        return self.call_tool(
            "query_bus_arrivals_by_route_id_and_stop_name_or_id",
            {"route_id": str(route_id).strip(), "stop_name_or_id": str(stop).strip()},
        )

    def query_timetable(self, route_id: str) -> Any:
        """某线的时刻表（服务端给的是一个制表符表格字符串）。"""
        return self.call_tool("query_route_timetable", {"route_id": str(route_id).strip()})

    def query_routes_at_stop(self, stop: str) -> Any:
        """这个站有哪些线（站名要和服务端完全一致）。"""
        return self.call_tool("query_routes_at_stop", {"stop_name": str(stop).strip()})

    def query_routes_between(self, from_stop: str, to_stop: str) -> Any:
        """从 A 站到 B 站有哪些直达线路（同一方向）。"""
        return self.call_tool(
            "query_routes_from_stop_to_stop",
            {"from_stop_name": str(from_stop).strip(), "to_stop_name": str(to_stop).strip()},
        )

    def query_schedules(self, route_id: str) -> Any:
        """某线每辆车的发车/到达时刻。"""
        return self.call_tool("query_schedules_by_route_id", {"route_id": str(route_id).strip()})

    def query_depot(self, route_id: str) -> Any:
        """某线的出车记录：每辆车从哪个场站、几点出车。"""
        return self.call_tool("query_route_depot", {"route_id": str(route_id).strip()})

    def query_route_list(self) -> Any:
        """服务端全量线路目录（168KB 左右，只在"更新线路目录"时拉一次）。"""
        return self.call_tool("query_route_list", {}, timeout=90)

    def query_journey_time(self, route_id: str, from_stop: str, to_stop: str) -> Any:
        """同一条线上，从 A 站到 B 站每辆车花了多久。"""
        return self.call_tool(
            "query_bus_journey_time_by_route_id_and_stop_names_or_ids",
            {
                "route_id": str(route_id).strip(),
                "from_stop_name_or_id": str(from_stop).strip(),
                "to_stop_name_or_id": str(to_stop).strip(),
            },
        )
