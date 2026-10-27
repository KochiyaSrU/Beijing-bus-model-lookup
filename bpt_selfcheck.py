"""程序自检：把"能不能正常用"拆成一条条可读的检查。

为什么单独一个模块：
- 只用标准库 + `bpt_data`，**不 import tkinter**，所以命令行也能用（`--self-test`）；
- 每条检查都带 `hint`（失败时告诉用户下一步点哪里），而不是只报"出错了"；
- 正常数据下必须全过，坏数据下必须只有对应那几条失败——这两点都有测试守着。

调用方式：
    from bpt_selfcheck import run_self_checks, format_report
    items = run_self_checks(application_dir(), settings, deep=True)
"""

from __future__ import annotations

import json
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import bpt_data as d


@dataclass(frozen=True)
class CheckItem:
    """一条自检结果。

    `ok=True` 通过；`blocking=False` 表示"只是提示"（例如标签覆盖率不满），
    不影响命令行退出码——否则每台机器都会因为"还没标完标签"而报错。
    """

    key: str
    title: str
    ok: bool
    detail: str = ""
    hint: str = ""
    blocking: bool = True

    @property
    def mark(self) -> str:
        if self.ok:
            return "✓"
        return "✗" if self.blocking else "!"


def _json_ok(path: Path) -> tuple[bool, str]:
    """文件不存在算通过（程序本来就能在没有它时正常工作）；坏了才算失败。"""
    if not path.is_file():
        return True, "还没有这个文件（不影响使用）"
    try:
        json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        return False, f"读不出来：{error}"
    return True, f"{path.stat().st_size / 1024:.1f} KB"


def _check_sheet(base: Path) -> CheckItem:
    path = base / d.DATA_FILENAME
    if not path.is_file():
        return CheckItem(
            "sheet", "号段表", False, f"没找到 {d.DATA_FILENAME}",
            "把号段表（Excel）放回程序目录，或从「备份与恢复」里恢复一份",
        )
    try:
        ranges = d.load_ranges(path)
    except Exception as error:  # noqa: BLE001 - 自检要把原因原样告诉用户
        return CheckItem("sheet", "号段表", False, f"解析失败：{error}",
                         "在「数据体检」里点「恢复最近的备份」，或用 Excel 手工修一下")
    models = len(d.model_names(ranges))
    if not ranges:
        return CheckItem("sheet", "号段表", False, "里面一条号段都没有",
                         "确认 Excel 的表头和内容没被改动；必要时恢复备份")
    return CheckItem("sheet", "号段表", True, f"{len(ranges)} 条号段 / {models} 种车型")


def _check_overlap(base: Path) -> CheckItem:
    try:
        ranges = d.load_ranges(base / d.DATA_FILENAME)
    except Exception as error:  # noqa: BLE001
        return CheckItem("overlap", "号段重叠", False, f"读不到号段表：{error}", "先修号段表")
    ordered = sorted(ranges, key=lambda item: (item.start, item.end))
    overlaps = [
        (left, right)
        for left, right in zip(ordered, ordered[1:])
        if right.start <= left.end
    ]
    if overlaps:
        first = overlaps[0]
        return CheckItem(
            "overlap", "号段重叠", False,
            f"{len(overlaps)} 处重叠，例如 {first[0].start}-{first[0].end} 与 {first[1].start}-{first[1].end}",
            "在「数据体检」里看「号段重叠」那一类，按提示改号段",
        )
    return CheckItem("overlap", "号段重叠", True, "没有重叠")


def _check_tags(base: Path, tag_checker) -> CheckItem:
    """标签覆盖率交给调用方注入（标签推导在界面层，这里不能 import tkinter）。"""
    if tag_checker is None:
        return CheckItem("tags", "车型标签", True, "跳过（没有提供标签检查器）")
    try:
        return tag_checker(base)
    except Exception as error:  # noqa: BLE001 - 检查器自身出错不该让自检崩
        return CheckItem("tags", "车型标签", False, f"检查时出错：{error}", "在「数据体检」里逐个确认标签")


def _check_jsons(base: Path) -> CheckItem:
    """用户数据文件：不存在算通过，坏了才算失败。"""
    files = [
        d.METRO_DATA_FILENAME, d.PROFILE_FILENAME, d.TAG_FILENAME,
        d.LINES_FILENAME, d.LINE_INFO_FILENAME, d.ROUTE_LIST_FILENAME, d.ONLINE_CACHE_FILENAME,
    ]
    bad: list[str] = []
    for name in files:
        ok, detail = _json_ok(base / name)
        if not ok:
            bad.append(f"{name}（{detail}）")
    if bad:
        return CheckItem("json", "数据文件", False, "；".join(bad),
                         "这些文件坏掉不影响查询；可以在「备份与恢复」里恢复，或直接删掉让它重建")
    return CheckItem("json", "数据文件", True, "都能正常读")


def _check_writable(base: Path) -> CheckItem:
    probe = base / ".selfcheck.tmp"
    try:
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except OSError as error:
        return CheckItem(
            "write", "写入权限", False, f"程序目录写不进去：{error}",
            "程序要在这个目录写设置和记录。别装在 Program Files，改装到「文档」或桌面下的文件夹",
        )
    return CheckItem("write", "写入权限", True, f"{base} 可写")


def _check_disk(base: Path) -> CheckItem:
    try:
        usage = shutil.disk_usage(base)
    except OSError as error:
        return CheckItem("disk", "磁盘空间", True, f"没量出来（{error}），不影响使用")
    free_mb = usage.free / (1024 * 1024)
    if free_mb < 200:
        return CheckItem("disk", "磁盘空间", False, f"只剩 {free_mb:.0f} MB",
                         "清理一下磁盘；备份和照片会占空间")
    return CheckItem("disk", "磁盘空间", True, f"剩余 {free_mb / 1024:.1f} GB")


def _check_version() -> CheckItem:
    frozen = bool(getattr(sys, "frozen", False))
    where = Path(sys.executable) if frozen else Path(__file__).resolve()
    try:
        stamp = datetime.fromtimestamp(where.stat().st_mtime)
    except OSError:
        return CheckItem("version", "程序版本", True, "取不到构建时间")
    kind = "打包版 EXE" if frozen else "源码运行"
    return CheckItem("version", "程序版本", True,
                     f"{kind} · 构建 {stamp:%Y-%m-%d %H:%M} · Python {sys.version.split()[0]}")


def _check_network(base: Path, settings: dict) -> CheckItem:
    """只在 deep 模式做；调用联网探测（不弹窗、不写缓存）。"""
    try:
        from bpt_online import OnlineSettings, probe_network
    except Exception as error:  # noqa: BLE001 - 联网模块缺失不该让自检崩
        return CheckItem("network", "联网（可选）", True, f"跳过：{error}")
    options = OnlineSettings.from_dict(settings.get("network") or {})
    if not options.usable:
        return CheckItem("network", "联网（可选）", True, "已关闭（本地功能不受影响）")
    ok, detail = probe_network(options)
    if ok:
        return CheckItem("network", "联网（可选）", True, detail)
    return CheckItem("network", "联网（可选）", False, detail,
                     "联网只是可选功能；想用它就检查网络/防火墙，或看「联网日志.txt」")


def run_self_checks(
    base: Path, settings: dict | None = None, *, deep: bool = False, tag_checker=None
) -> list[CheckItem]:
    """跑一遍自检。`deep=True` 时额外探测联网（慢几秒）。"""
    base = Path(base)
    items = [
        _check_sheet(base),
        _check_overlap(base),
        _check_tags(base, tag_checker),
        _check_jsons(base),
        _check_writable(base),
        _check_disk(base),
        _check_version(),
    ]
    if deep:
        items.append(_check_network(base, settings or {}))
    return items


def format_report(items: list[CheckItem]) -> list[str]:
    """给命令行用：每行一条。"""
    lines = []
    for item in items:
        line = f"{item.mark} {item.title}：{item.detail}"
        lines.append(line)
        if not item.ok and item.hint:
            lines.append(f"    → {item.hint}")
    failed = [item for item in items if not item.ok and item.blocking]
    notes = [item for item in items if not item.ok and not item.blocking]
    lines.append("")
    if failed:
        lines.append(f"有 {len(failed)} 项需要处理（见上面的 → 提示）。")
    elif notes:
        lines.append(f"能正常用；另有 {len(notes)} 项是提示（不影响使用）。")
    else:
        lines.append("全部通过。")
    return lines


def has_blocking_failure(items: list[CheckItem]) -> bool:
    """命令行退出码用：只有 blocking 的失败才算失败。"""
    return any(not item.ok and item.blocking for item in items)


def as_json(items: list[CheckItem]) -> str:
    payload = [
        {"key": item.key, "title": item.title, "ok": item.ok, "detail": item.detail, "hint": item.hint}
        for item in items
    ]
    return json.dumps(payload, ensure_ascii=False, indent=2)
