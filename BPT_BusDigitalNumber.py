"""北京公交车型查询器。运行时从同目录的 Excel 文件加载号段映射。"""

from __future__ import annotations

import csv
import ctypes
import functools
import json
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from ctypes import wintypes
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, simpledialog, ttk

from PIL import Image, ImageGrab, ImageTk, UnidentifiedImageError

import bpt_web
from bpt_selfcheck import CheckItem  # 自检结果类型（自检模块本身不依赖界面）
from bpt_data import (  # noqa: F401 - 数据层搬到了 bpt_data.py，这里全部再导出，旧引用不变
    _column_letter,
    _first_sheet_path,
    _read_shared_strings,
    _worksheet_table,
    _write_json,
    APP_TITLE,
    append_spot,
    application_dir,
    apply_line_plan,
    apply_merge,
    applicable_suggestions,
    build_links_from_records,
    fact_for,
    group_by_suffix,
    learn_tag_rules,
    LineLink,
    LinePlan,
    LINE_STATUSES,
    LINE_STATUS_LABELS,
    LINE_STATUS_ORDER,
    line_discrepancies,
    line_display,
    line_family,
    line_family_suggestions,
    line_key,
    lines_for_model,
    lines_for_number,
    lines_path,
    load_lines,
    load_lines_lenient,
    load_district_bus,
    merge_lines,
    models_for_line,
    numbers_for_line,
    all_line_names,
    save_lines,
    save_district_bus_template,
    save_district_bus,
    district_bus_entry_cars,
    district_bus_entry_conflicts,
    district_bus_entry_summary,
    district_bus_apply_batch,
    district_bus_companies_seen,
    district_bus_set_fields,
    district_bus_duplicate_keys,
    district_bus_merge_duplicates,
    district_bus_plan_batch,
    district_bus_quality,
    district_bus_split_plates,
    district_bus_companies_for,
    district_bus_districts,
    district_bus_lookup,
    district_bus_lookup_any,
    district_bus_lookup_plate,
    district_bus_path,
    district_bus_template,
    looks_like_plate,
    normalize_plate,
    remove_district_bus_entry,
    upsert_district_bus_entry,
    DISTRICT_BUS_FILENAME,
    DISTRICT_BUS_NUMBER_KINDS,
    DistrictBusCompany,
    DistrictBusEntry,
    DistrictBusRange,
    DistrictBusResult,
    suggest_tag,
    suggest_tags_for_models,
    TagFact,
    TAG_FACTS,
    TagRuleHint,
    TagSuggestion,
    tag_suffix_candidates,
    backup_directory,
    BACKUP_DIRECTORY_NAME,
    backup_file,
    BACKUP_KEEP,
    BACKUP_SOURCES,
    BusRange,
    CLIPBOARD_INTERVAL_MS,
    CREATE_NO_WINDOW,
    DATA_FILENAME,
    DataError,
    ERROR_LOG_FILENAME,
    EXPECTED_HEADERS,
    extra_columns_by_row,
    filter_spots,
    find_model,
    find_range,
    group_missing_numbers,
    HISTORY_FILENAME,
    HISTORY_LIMIT,
    history_path,
    ICON_FILENAME,
    IMAGE_DIRECTORY_NAME,
    IMAGE_EXTENSIONS,
    JPEG_QUALITY,
    list_backups,
    load_history,
    load_ranges,
    load_ranges_lenient,
    load_spots,
    load_ride_records,
    looks_like_metro_model,
    looks_like_out_of_town,
    parse_ride_sheet,
    read_sheets,
    reconcile_rides,
    RideRecord,
    log_exception,
    make_spot,
    MAX_ERROR_LOG_BYTES,
    MAX_IMAGE_PIXELS,
    merge_conflict_keys,
    merge_ranges,
    MergePlan,
    METRO_DATA_FILENAME,
    METRO_SOURCE_FILENAME,
    MetroError,
    MetroRange,
    model_names,
    OFFICE_RELATIONSHIPS_NS,
    parse_number_list,
    PROFILE_FILENAME,
    ProfileError,
    PHOTO_DIRECTORY_NAME,
    prune_backups,
    RangeError,
    ranges_for_model,
    read_table,
    RELATIONSHIPS_NS,
    remember_number,
    replace_spot,
    resource_path,
    restore_backup,
    save_ranges,
    save_spots,
    segment_lookup,
    SETTINGS_FILENAME,
    SPOT_LOG_FILENAME,
    spot_log_path,
    spot_models,
    spot_summary,
    spot_summary_text,
    SpotEntry,
    SPREADSHEET_NS,
    TAG_FILENAME,
    validate_number,
    validate_range,
    write_table,
    # 联网写入线路数据用到的这一批（V2）
    LineEvidence,
    SOURCE_LABELS,
    cache_entry,
    cache_prune,
    cache_put,
    delete_by_source,
    lines_from_history,
    load_online_cache,
    month_key,
    online_cache_path,
    online_stale,
    save_online_cache,
    source_label,
    source_rank,
    # 线路资料（V3）
    LINE_INFO_FILENAME,
    line_info_directions,
    line_info_lines_without_stops,
    line_info_path,
    line_info_stops,
    lines_from_route_history,
    latest_link,
    load_line_info,
    save_line_info,
    update_line_info,
    # 联网日志与实时解析（V5）
    append_online_log,
    clear_online_log,
    online_log_path,
    parse_arrivals,
    parse_timetable,
    parse_track_latest,
    read_online_log,
    recent_online_errors,
    # 服务端线路目录（本地校验 / 线路号补全）
    ROUTE_LIST_FILENAME,
    load_route_list,
    parse_route_list,
    route_list_numbers,
    route_list_path,
    route_list_suggest,
    save_route_list,
    validate_line_numbers,
)
from bpt_online import (  # 联网查询：只用标准库，界面里按需调用    MCPClient,
    OnlineError,
    OnlineService,
    OnlineSettings,
    OnlineUnavailable,
    error_text,
    history_to_links,
    model_lookup,
    offline_forced,
    short_error,
    shared_service,
    today_stamp,
    USER_AGENT,
)


def online_settings() -> OnlineSettings:
    """从 界面设置.json 的 network 段读联网设置（缺项全用默认值）。"""
    try:
        payload = load_settings(application_dir()).get("network")
    except Exception:  # noqa: BLE001 - 设置读不出来就按默认走，绝不因为联网配置挡住启动
        payload = None
    return OnlineSettings.from_dict(payload)


def install_exception_handling(app: tk.Misc) -> None:
    """打包成 --windowed 后没有 stderr，异常会无声消失；这里改成写日志 + 弹窗。"""
    shown: set[tuple[str, str]] = set()

    def report(exc_type: type, exc_value: BaseException, exc_tb: object) -> None:
        path = log_exception(exc_type, exc_value, exc_tb, "界面回调")
        signature = (getattr(exc_type, "__name__", str(exc_type)), str(exc_value))
        if signature in shown:
            return
        shown.add(signature)
        message = "这个操作出错了，没有完成。"
        if path is not None:
            message += f"\n\n详细信息已写入：\n{path.name}"
        else:
            message += "\n\n（错误日志写入失败）"
        message += f"\n\n{signature[0]}: {signature[1]}"
        try:
            messagebox.showerror("程序出错", message, parent=app)
        except Exception:
            pass

    app.report_callback_exception = report  # type: ignore[method-assign]
    sys.excepthook = report  # type: ignore[assignment]


# --------------------------------------------------------------------------- #
# 拍车模式的数据：查询历史、拍车日志
# --------------------------------------------------------------------------- #


# --------------------------------------------------------------------------- #
# 拍车模式：剪贴板监听、全局热键、语音朗读
# --------------------------------------------------------------------------- #


def read_clipboard_text(widget: tk.Misc) -> str:
    """读剪贴板文本；剪贴板是图片或为空时返回空串。"""
    try:
        return str(widget.clipboard_get()).strip()
    except tk.TclError:
        return ""


def extract_clipboard_number(text: str) -> int | None:
    """从剪贴板内容里认出五位自编号（允许前后有空白、换行或中文说明）。"""
    value = text.strip()
    direct = validate_number(value)
    if direct is not None:
        return direct
    for token in re.split(r"[\s,，、;；|/\\]+", value):
        number = validate_number(token)
        if number is not None:
            return number
    match = re.search(r"(?<!\d)(\d{5})(?!\d)", value)
    if match:
        return validate_number(match.group(1))
    return None


class ClipboardWatcher:
    """定时看剪贴板里有没有出现新的五位自编号。"""

    def __init__(self, widget: tk.Misc, callback: Callable[[int], None], interval_ms: int = CLIPBOARD_INTERVAL_MS) -> None:
        self.widget = widget
        self.callback = callback
        self.interval_ms = interval_ms
        self._job: str | None = None
        self._last = ""

    @property
    def running(self) -> bool:
        return self._job is not None

    def start(self) -> None:
        if self._job is None:
            self._last = read_clipboard_text(self.widget)
            self._schedule()

    def stop(self) -> None:
        if self._job is not None:
            try:
                self.widget.after_cancel(self._job)
            except tk.TclError:
                pass
            self._job = None

    def _schedule(self) -> None:
        try:
            self._job = self.widget.after(self.interval_ms, self._tick)
        except tk.TclError:
            self._job = None

    def _tick(self) -> None:
        self._job = None
        text = read_clipboard_text(self.widget)
        if text != self._last:
            self._last = text
            number = extract_clipboard_number(text)
            if number is not None:
                self.callback(number)
        if self.widget.winfo_exists():
            self._schedule()


def clipboard_image() -> Image.Image | None:
    """取剪贴板里的图片（截图或复制的图片文件），拿不到返回 None。"""
    try:
        data = ImageGrab.grabclipboard()
    except Exception:
        return None
    if isinstance(data, Image.Image):
        try:
            data.load()
            return data.convert("RGB")
        except Exception:
            return None
    if isinstance(data, list):
        for item in data:
            path = Path(str(item))
            if path.suffix.lower() in IMAGE_EXTENSIONS and path.is_file():
                try:
                    return load_image(path)
                except ProfileError:
                    continue
    return None


def speak_text(text: str) -> bool:
    """用 Windows 自带语音念一句；没有 PowerShell 或调用失败就静默放弃。"""
    value = text.strip()
    if not value or sys.platform != "win32":
        return False
    script = (
        "Add-Type -AssemblyName System.Speech; "
        "(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('"
        + value.replace("'", "''")
        + "')"
    )
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            creationflags=CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except (OSError, ValueError):
        return False
    return True


class GlobalHotkey:
    """注册一个系统级热键（默认 Ctrl+Alt+B）。

    实现在后台线程里 RegisterHotKey + GetMessage；注册失败（被别的程序占用）
    只把 registered 置为 False，不影响程序其余部分。
    """

    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_NOREPEAT = 0x4000
    WM_HOTKEY = 0x0312
    WM_QUIT = 0x0012
    HOTKEY_ID = 0xB07

    def __init__(self, widget: tk.Misc, callback: Callable[[], None], modifiers: int | None = None, key: int = 0x42) -> None:
        self.widget = widget
        self.callback = callback
        self.modifiers = self.MOD_CONTROL | self.MOD_ALT if modifiers is None else modifiers
        self.key = key
        self.registered = False
        self._thread: threading.Thread | None = None
        self._thread_id = 0
        self._ready = threading.Event()
        self._stopping = False
        # 热键线程只往这里塞事件，界面回填由主线程的 _poll 做
        self._fired: "queue.Queue[bool]" = queue.Queue()

    @property
    def description(self) -> str:
        parts = []
        if self.modifiers & self.MOD_CONTROL:
            parts.append("Ctrl")
        if self.modifiers & self.MOD_ALT:
            parts.append("Alt")
        if self.modifiers & self.MOD_SHIFT:
            parts.append("Shift")
        parts.append(chr(self.key) if 0x41 <= self.key <= 0x5A else f"0x{self.key:X}")
        return "+".join(parts)

    def start(self) -> bool:
        if self._thread is not None:
            return self.registered
        if sys.platform != "win32":
            return False
        self._ready.clear()
        self._stopping = False
        self._thread = threading.Thread(target=self._run, name="bpt-hotkey", daemon=True)
        self._thread.start()
        self._ready.wait(timeout=3)
        if not self.registered:
            self._thread = None
            return self.registered
        # 由**主线程**轮询热键事件（Tkinter 不是线程安全的：以前在工作线程里直接
        # after_idle 回调，慢一点就抛 RuntimeError: main thread is not in main loop，
        # 而 except tk.TclError 抓不到它 —— 热键会静默失效）。
        self.widget.after(200, self._poll)
        return self.registered

    def _poll(self) -> None:
        """在主线程里把热键按下事件交给回调。"""
        fired = False
        while True:
            try:
                self._fired.get_nowait()
            except queue.Empty:
                break
            fired = True
        if fired and not self._stopping:
            self.callback()
        if self.registered and not self._stopping:
            self.widget.after(200, self._poll)

    def stop(self) -> None:
        self._stopping = True
        if self._thread is not None and self._thread_id:
            try:
                ctypes.windll.user32.PostThreadMessageW(self._thread_id, self.WM_QUIT, 0, 0)
            except Exception:
                pass
        self._thread = None
        self._thread_id = 0
        self.registered = False

    def _run(self) -> None:
        user32 = ctypes.windll.user32
        try:
            self._thread_id = ctypes.windll.kernel32.GetCurrentThreadId()
            if not user32.RegisterHotKey(None, self.HOTKEY_ID, self.modifiers | self.MOD_NOREPEAT, self.key):
                self.registered = False
                self._ready.set()
                return
        except Exception:
            self.registered = False
            self._ready.set()
            return
        self.registered = True
        self._ready.set()
        message = wintypes.MSG()
        try:
            while not self._stopping and user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
                if message.message == self.WM_HOTKEY:
                    self._fired.put(True)  # 只入队；碰界面的事交给主线程的 _poll
        finally:
            try:
                user32.UnregisterHotKey(None, self.HOTKEY_ID)
            except Exception:
                pass
            self.registered = False


DRAFT_TEMPLATE = "{厂商} {燃料} {车长}级客车，自编号 {号段}。"


def draft_profile(model: str, tags: ModelTags, ranges: list[BusRange], template: str = DRAFT_TEMPLATE, metro: bool = False) -> dict[str, str]:
    """按模板给没有资料的车型起草一段简介 + 参数。"""
    known = ranges_for_model(model, ranges)
    span = "、".join(f"{item.start}-{item.end}" for item in known) or "暂无对应号段"
    total = sum(item.end - item.start + 1 for item in known)
    values = {
        "厂商": tags.manufacturer,
        "燃料": tags.fuel,
        "车长": tags.length.replace(" 米级", "").strip() or tags.length,
        "车型": model,
        "号段": span,
        "编号数": str(total),
        "号段数": str(len(known)),
    }
    try:
        introduction = template.format(**values).strip()
    except (KeyError, IndexError, ValueError):
        introduction = f"{tags.manufacturer} {tags.fuel} {tags.length}客车。"
    specifications = (
        f"生产厂商：{tags.manufacturer}\n"
        f"燃料类型：{tags.fuel}\n"
        f"车长级别：{tags.length}\n"
        f"车体形式：{tags.body}\n"
        f"对应号段：{span}（共 {total} 个自编号）"
    )
    return {"introduction": introduction, "specifications": specifications}


def empty_profile_models(profiles: dict, ranges: list[BusRange], existing_images: set[str] | None = None) -> list[str]:
    """还没有任何资料的车型（顺序按车型名）。"""
    names = existing_images
    return [
        name
        for name in model_names(ranges)
        if profile_is_empty(profiles.get(name, {}), None, names)
    ]


def _web_bound(value: object) -> int | str:
    """网页版里的号段边界：纯数字转 int，带字母前缀的原样保留。"""
    text = str(value).strip()
    return int(text) if text.isdigit() else text


def web_payload(
    ranges: list[BusRange],
    overrides: dict | None,
    profiles: dict,
    metro_ranges: list[MetroRange],
    base_directory: Path | None = None,
    *,
    embed_images: bool = True,
    notices: list[str] | None = None,
    spots: list[SpotEntry] | None = None,
    max_spot_photos: int = 60,
) -> dict:
    """把当前数据整理成网页版需要的普通字典（bpt_web 只认这个结构）。"""
    names = model_names(ranges)
    tags: dict[str, list[str]] = {}
    aliases: dict[str, list[str]] = {}
    for name in names:
        tag = model_tags(name, overrides)
        tags[name] = [tag.manufacturer, tag.fuel, tag.length, tag.body]
        names_for_model = model_aliases(name, overrides)
        if names_for_model:
            aliases[name] = names_for_model
    payload_profiles: dict[str, dict[str, str]] = {}
    for name in names:
        profile = profiles.get(name, {})
        entry: dict[str, str] = {}
        introduction = str(profile.get("introduction", "")).strip()
        specifications = str(profile.get("specifications", "")).strip()
        if introduction:
            entry["intro"] = introduction
        if specifications:
            entry["spec"] = specifications
        images = profile_images(profile)
        if embed_images and base_directory is not None and images:
            uri = bpt_web.thumbnail_data_uri(stored_image_path(base_directory, images[0]["name"]))
            if uri:
                entry["image"] = uri
        if entry:
            payload_profiles[name] = entry
    metro: dict[str, list[list[object]]] = {}
    for item in sorted(metro_ranges, key=lambda entry: (entry.line, _metro_sort_key(entry.start))):
        metro.setdefault(item.line, []).append([_web_bound(item.start), _web_bound(item.end), item.model])
    total = sum(item.end - item.start + 1 for item in ranges)
    spot_rows: list[dict[str, str]] = []
    photo_budget = max_spot_photos if embed_images and base_directory is not None else 0
    for entry in (spots or [])[:300]:
        row = {"time": entry.time, "number": entry.number, "model": entry.model, "note": entry.note}
        if entry.photo and photo_budget > 0:
            uri = bpt_web.thumbnail_data_uri(stored_photo_path(base_directory, entry.photo), max_size=(280, 200), quality=62)
            if uri:
                row["photo"] = uri
                photo_budget -= 1
        spot_rows.append(row)
    return {
        "title": APP_TITLE,
        "subtitle": f"{len(ranges)} 个号段 · {len(names)} 种车型 · 覆盖 {total} 个自编号",
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "ranges": [[item.start, item.end, item.model] for item in sorted(ranges, key=lambda entry: entry.start)],
        "tags": tags,
        "aliases": aliases,
        "profiles": payload_profiles,
        "metro": metro,
        "spots": spot_rows,
        "spot_summary": spot_summary_text(spot_summary(list(spots or []))),
        "notices": notices or [],
    }


def export_web_page(
    path: Path,
    ranges: list[BusRange],
    overrides: dict | None,
    profiles: dict,
    metro_ranges: list[MetroRange],
    base_directory: Path | None = None,
    *,
    embed_images: bool = True,
    notices: list[str] | None = None,
    spots: list[SpotEntry] | None = None,
) -> tuple[Path, int]:
    """导出单文件网页版，返回 (路径, 字节数)。"""
    payload = web_payload(
        ranges,
        overrides,
        profiles,
        metro_ranges,
        base_directory,
        embed_images=embed_images,
        notices=notices,
        spots=spots,
    )
    try:
        bpt_web.write_html(path, payload)
    except OSError as error:
        raise DataError(f"无法写入 {path.name}：{error}") from error
    return path, path.stat().st_size


def model_card_payload(
    model: str,
    ranges: list[BusRange],
    overrides: dict | None,
    profile: dict,
    base_directory: Path | None = None,
    *,
    embed_images: bool = True,
) -> dict:
    """一个车型的「卡片」数据：标签、图片、简介参数、号段。"""
    tags = model_tags(model, overrides)
    known = ranges_for_model(model, ranges)
    covered = sum(item.end - item.start + 1 for item in known)
    photos: list[dict[str, str]] = []
    if embed_images and base_directory is not None:
        for item in profile_images(profile):
            uri = bpt_web.thumbnail_data_uri(stored_image_path(base_directory, item["name"]), max_size=(760, 560), quality=76)
            if uri:
                photos.append({"src": uri, "caption": item.get("caption", "")})
    return {
        "model": model,
        "subtitle": f"{tags.manufacturer} · {tags.fuel} · {tags.length} · {tags.body}",
        "tags": [tags.manufacturer, tags.fuel, tags.length, tags.body],
        "photos": photos,
        "introduction": str(profile.get("introduction", "")).strip(),
        "specifications": str(profile.get("specifications", "")).strip(),
        "ranges": [{"start": f"{item.start:05d}", "end": f"{item.end:05d}", "count": item.end - item.start + 1, "model": item.model} for item in known],
        "covered": covered,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def export_model_card(
    path: Path,
    model: str,
    ranges: list[BusRange],
    overrides: dict | None,
    profile: dict,
    base_directory: Path | None = None,
) -> tuple[Path, int]:
    """导出单车型卡片，返回 (路径, 字节数)。"""
    payload = model_card_payload(model, ranges, overrides, profile, base_directory)
    try:
        bpt_web.write_model_card(path, payload)
    except OSError as error:
        raise DataError(f"无法写入 {path.name}：{error}") from error
    return path, path.stat().st_size


def write_tagged_workbook(path: Path, ranges: list[BusRange], overrides: dict) -> int:
    """导出带厂商/燃料/车长列的 Workbook，返回写出的数据行数。"""
    header = ["号段起始", "号段结束", "车型", "厂商", "燃料", "车长"]
    rows = []
    for item in sorted(ranges, key=lambda entry: entry.start):
        tag = model_tags(item.model, overrides)
        rows.append([str(item.start), str(item.end), item.model, tag.manufacturer, tag.fuel, tag.length])
    try:
        write_table(path, header, rows, numeric_columns={0, 1})
    except OSError as error:
        raise DataError(f"无法写入 {path.name}，请确认目标文件夹可写入。") from error
    return len(rows)


TAG_COLUMN_NAMES = {"厂商": "manufacturer", "燃料": "fuel", "车长": "length", "车体形式": "body"}


def load_tag_columns(path: Path) -> dict[str, dict[str, str]]:
    """读取表里的「车型 / 厂商 / 燃料 / 车长」列，按车型整理成覆盖值。"""
    header, rows = read_table(path)
    if not header or "车型" not in header:
        return {}
    model_index = header.index("车型")
    field_indexes = {
        TAG_COLUMN_NAMES[name]: index for index, name in enumerate(header) if name in TAG_COLUMN_NAMES
    }
    if not field_indexes:
        return {}
    result: dict[str, dict[str, str]] = {}
    for row in rows:
        padded = (row + [""] * len(header))[: len(header)]
        model = padded[model_index].strip()
        if not model:
            continue
        entry = {field: padded[index].strip() for field, index in field_indexes.items() if padded[index].strip()}
        if entry:
            result.setdefault(model, {}).update(entry)
    return result


def merge_tag_columns(overrides: dict, columns: dict[str, dict[str, str]], known_models: set[str]) -> tuple[dict, list[str]]:
    """把导入的标签合并进覆盖表；返回（新的覆盖表, 被忽略的车型）。"""
    merged = clean_tag_overrides(overrides if isinstance(overrides, dict) else {})
    ignored: list[str] = []
    for model, values in columns.items():
        if model not in known_models:
            ignored.append(model)
            continue
        apply_model_tag_values(merged, model, values)
    return clean_tag_overrides(merged), ignored


# --------------------------------------------------------------------------- #
# Model code parsing (manufacturer / fuel / length class) and quick lookups
# --------------------------------------------------------------------------- #

# 型号前缀 → 厂商。只收录能确认的；其余前缀原样显示，可在 车型标签.json 里补全。
MANUFACTURER_NAMES = {
    "BJ": "北汽福田",
    "BJD": "北京电车",
    "BK": "北京客车",
    "HFF": "安凯客车",
    "ZK": "宇通客车",
    "XMQ": "厦门金龙",
    "XML": "厦门金旅",
    "JNP": "青年汽车",
    "KLQ": "苏州海格",
    "GTQ": "珠海广通",
    "LCK": "聊城中通",
    "CA": "一汽客车",
    "BYD": "比亚迪",
}
UNKNOWN_TAG = "未标注"
ANY_TAG = "全部"
TAG_FIELDS = ("manufacturer", "fuel", "length", "body")
# 判断顺序很重要：PHEV / SHEV / FCEV 都含 EV，必须排在纯电前面。
FUEL_RULES: tuple[tuple[str, str], ...] = (
    ("燃料电池", r"FCEV"),
    ("插电混动", r"PHEV"),
    ("增程混动", r"SHEV"),
    ("天然气", r"(CNG|LNG|HNG)"),
    ("无轨电车", r"^(BJD|WG)|-WG\d"),
    ("纯电动", r"EV"),
)

LENGTH_PRIMARY_PATTERN = re.compile(r"^[A-Za-z][A-Za-z\-]*6(\d{2,3})")
LENGTH_FALLBACK_PATTERN = re.compile(r"[A-Za-z\-]+(\d{2})\d")
MANUFACTURER_KEY_PATTERN = re.compile(r"^[A-Za-z]+")

# 车体形式：
# · 铰接：北京路上 16 / 18 米级基本就是铰接车，本地数据里 16 台全对得上。
# · 三轴：13 / 14 米级的单机基本都是三轴车（BJ6140C8CTD、BJ6140SHEVCA-1、BJ6147SHEVCA、
#   ZK6140HNG2…，多个来源都直接写"三轴"）；14 米以上的刚体客车不可能是两轴。
# · 双层：**没有可靠的型号规律**——"GS" 在不同厂家里含义不同，比如 DNC6107BEVGS1 是
#   远程牌纯电动低入口城市客车（BEVGS1 里的 GS 只是版本代号），按"GS 就是双层"去猜会猜错。
#   所以这里只列已核实的双层型号；其余按下面的规则走，遇到新的双层车用「批量设置标签」
#   或标签规则标一次即可（标过之后数据体检不会再提示）。
ARTICULATED_LENGTHS = ("16 米级", "18 米级")
THREE_AXLE_LENGTHS = ("13 米级", "14 米级")
KNOWN_DOUBLE_DECKERS = (
    "JNP6130GSC",  # 青年牌 JNP6130GSC 双层城市客车
    "HFF6110GS01D",  # 安凯 GS 系列双层（注意型号是 GS01D，不是 GS101D）
    "HFF6120GS01C",  # 安凯 GS 系列双层
    "HFF6122GS03EV",  # 安凯 GS 系列双层（纯电）
)
_DOUBLE_DECKER_KEYS = {code.upper() for code in KNOWN_DOUBLE_DECKERS}
BODY_VALUES = ("单机", "铰接", "双层", "三轴", UNKNOWN_TAG)


@dataclass(frozen=True)
class ModelTags:
    manufacturer: str
    fuel: str
    length: str
    body: str = UNKNOWN_TAG


def manufacturer_key(model: str) -> str:
    """Return the leading letter code of a model name, e.g. BJ for BJ6123C7C4D."""
    match = MANUFACTURER_KEY_PATTERN.match(model)
    return match.group(0).upper() if match else ""


def default_manufacturer(model: str) -> str:
    key = manufacturer_key(model)
    if not key:
        return UNKNOWN_TAG
    return MANUFACTURER_NAMES.get(key, key)


def default_fuel(model: str) -> str:
    upper = model.upper()
    for label, pattern in FUEL_RULES:
        if re.search(pattern, upper):
            return label
    return UNKNOWN_TAG


def default_length_class(model: str) -> str:
    """粗略车长级别：客车类别号 6 后面以 1 开头是 10~18 米，否则是 6~9 米。

    例如 BJ6123→12 米级、BJ6851→8 米级（实际 8.5 米）、BJ6105→10 米级（实际 10.5 米）。
    型号编号规则各厂并不统一，精确车长请在 车型标签.json 里按车型覆盖。
    """
    match = LENGTH_PRIMARY_PATTERN.match(model)
    if match:
        digits = match.group(1)
        metres = int(digits[:2]) if digits[0] == "1" else int(digits[0])
    else:
        fallback = LENGTH_FALLBACK_PATTERN.search(model)
        if fallback is None:
            return UNKNOWN_TAG
        metres = int(fallback.group(1))
    return f"{metres} 米级" if 6 <= metres <= 18 else UNKNOWN_TAG


def default_body(model: str, length: str) -> str:
    """车体形式：铰接 / 双层 / 三轴 / 单机。

    有依据的推断按这个顺序判：
    1. 车长 16 / 18 米级 → 铰接（本地数据 16 台全对）；
    2. 型号在 KNOWN_DOUBLE_DECKERS 里 → 双层；
    3. 车长 13 / 14 米级 → 三轴（BJ6140 / BJ6147 / BJ6149 / JNP6140 / ZK6140 这类 14 米
       刚体客车都是三轴，多个来源直接这么写）；
    4. 其余 → 单机。

    **型号里的字母不能当推断依据**：DNC6107BEVGS1 是低入口单机，`BEVGS1` 里的 GS 只是
    版本代号。清单之外的双层车请用「批量设置标签」或标签规则标一次；这个字段是自由文本，
    想加"双层铰接"之类的分类直接填就行。车长都没推断出来时留「未标注」，不硬猜。
    """
    if length == UNKNOWN_TAG:
        return UNKNOWN_TAG
    if length in ARTICULATED_LENGTHS:
        return "铰接"
    if model.strip().upper() in _DOUBLE_DECKER_KEYS:
        return "双层"
    if length in THREE_AXLE_LENGTHS:
        return "三轴"
    return "单机"


def body_is_guessed_double_decker(model: str) -> bool:
    """这条「双层」是按内置清单猜的（清单里的型号值得人工核对一下）。"""
    return model.strip().upper() in _DOUBLE_DECKER_KEYS


def tag_overrides_path(base_directory: Path) -> Path:
    return base_directory / TAG_FILENAME

def load_tag_overrides(base_directory: Path) -> dict:
    """Read optional hand-written corrections; a broken file never blocks startup."""
    path = tag_overrides_path(base_directory)
    if not path.is_file():
        return {}
    try:
        contents = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return contents if isinstance(contents, dict) else {}


def save_tag_overrides(base_directory: Path, overrides: dict) -> None:
    """Persist tag corrections atomically, keeping a one-generation backup."""
    path = tag_overrides_path(base_directory)
    temporary = path.with_suffix(".tmp")
    try:
        backup_file(path)
        temporary.write_text(json.dumps(overrides, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(path)
    except OSError as error:
        temporary.unlink(missing_ok=True)
        raise ProfileError(f"无法保存 {TAG_FILENAME}，请确认程序文件夹可写入。") from error


def default_manufacturer_name(prefix: str) -> str:
    """内置表里该前缀的厂商名（没有就返回前缀本身）。"""
    return MANUFACTURER_NAMES.get(prefix, prefix)


def apply_model_tag_values(overrides: dict, model: str, values: dict[str, str], *, force: bool = False) -> None:
    """把若干标签写进覆盖表；留空或与自动推断相同就等于不覆盖。

    `force=True` 用于「确认程序猜的就是对的」：即使值跟自动推断一样也**明确记下来**，
    这样数据体检里"待确认"的那几条确认过之后就不会再冒出来。
    """
    auto = model_tags(model)
    models = overrides.setdefault("models", {})
    entry = dict(models.get(model, {})) if isinstance(models.get(model), dict) else {}
    for field in TAG_FIELDS:
        if field not in values:
            continue
        value = str(values[field]).strip()
        if not value:
            entry.pop(field, None)
        elif value == getattr(auto, field) and not force:
            entry.pop(field, None)
        else:
            entry[field] = value
    if entry:
        models[model] = entry
    else:
        models.pop(model, None)


def apply_manufacturer_prefix(overrides: dict, prefix: str, name: str) -> None:
    """把某个前缀的厂商名写进覆盖表；等于内置名或留空则删除这条映射。"""
    names = overrides.setdefault("manufacturers", {})
    name = str(name).strip()
    if prefix and name and name != default_manufacturer_name(prefix):
        names[prefix] = name
    else:
        names.pop(prefix, None)


def manufacturer_is_unconfirmed(model: str, overrides: dict | None = None) -> bool:
    """厂商名仍是前缀代码、或被人为标了「待确认」时算未确认。"""
    return _manufacturer_unconfirmed(model_tags(model, overrides).manufacturer, model)


def _manufacturer_unconfirmed(manufacturer: str, model: str) -> bool:
    """同上，但厂商名已经算好了（批量体检时避免重复推断）。"""
    return manufacturer == manufacturer_key(model) or "待确认" in manufacturer


def clean_tag_overrides(overrides: dict) -> dict:
    """Drop empty or malformed sections so the saved file stays readable."""
    result: dict = {}
    names = overrides.get("manufacturers")
    if isinstance(names, dict):
        kept = {str(key): str(value).strip() for key, value in names.items() if str(value).strip()}
        if kept:
            result["manufacturers"] = kept
    models = overrides.get("models")
    if isinstance(models, dict):
        kept_models: dict[str, dict[str, str]] = {}
        for model, entry in models.items():
            if not isinstance(entry, dict):
                continue
            kept_fields = {
                field: str(entry[field]).strip()
                for field in TAG_FIELDS
                if isinstance(entry.get(field), str) and str(entry[field]).strip()
            }
            if kept_fields:
                kept_models[str(model)] = kept_fields
        if kept_models:
            result["models"] = kept_models
    rules = overrides.get("rules")
    if isinstance(rules, list):
        kept_rules: list[dict[str, str]] = []
        for rule in rules:
            if not isinstance(rule, dict):
                continue
            pattern = str(rule.get("pattern", "")).strip()
            if not pattern:
                continue
            values = {
                field: str(rule[field]).strip()
                for field in TAG_FIELDS
                if isinstance(rule.get(field), str) and str(rule[field]).strip()
            }
            if values:
                kept_rules.append({"pattern": pattern, **values})
        if kept_rules:
            result["rules"] = kept_rules
    aliases = overrides.get("aliases")
    if isinstance(aliases, dict):
        kept_aliases: dict[str, list[str]] = {}
        for model, names in aliases.items():
            if isinstance(names, str):
                names = [names]
            if not isinstance(names, list):
                continue
            cleaned: list[str] = []
            for name in names:
                text = str(name).strip()
                if text and text not in cleaned:
                    cleaned.append(text)
            if cleaned:
                kept_aliases[str(model)] = cleaned
        if kept_aliases:
            result["aliases"] = kept_aliases
    return result


def model_aliases(model: str, overrides: dict | None = None) -> list[str]:
    """车型的别名 / 俗称（用户手工维护）。"""
    if not isinstance(overrides, dict):
        return []
    aliases = overrides.get("aliases")
    if not isinstance(aliases, dict):
        return []
    names = aliases.get(model)
    if isinstance(names, str):
        names = [names]
    if not isinstance(names, list):
        return []
    return [str(name).strip() for name in names if str(name).strip()]


def alias_index(overrides: dict | None = None, models: list[str] | None = None) -> dict[str, list[str]]:
    """别名 → 车型列表（用于搜索匹配，忽略大小写）。"""
    index: dict[str, list[str]] = {}
    aliases = overrides.get("aliases") if isinstance(overrides, dict) else None
    if not isinstance(aliases, dict):
        return index
    known = set(models) if models is not None else None
    for model in aliases:
        if known is not None and model not in known:
            continue
        for name in model_aliases(str(model), overrides):
            index.setdefault(name.casefold(), []).append(str(model))
    return index


@dataclass(frozen=True)
class TagRule:
    pattern: str
    values: dict[str, str]


_PATTERN_CACHE: dict[str, re.Pattern | None] = {}


def compiled_pattern(pattern: str) -> re.Pattern | None:
    """编译并缓存正则；写错的规则返回 None（当作不匹配，不影响其它规则）。"""
    if pattern not in _PATTERN_CACHE:
        try:
            _PATTERN_CACHE[pattern] = re.compile(pattern, re.IGNORECASE)
        except re.error:
            _PATTERN_CACHE[pattern] = None
    return _PATTERN_CACHE[pattern]


def parse_rules(overrides: dict | None) -> list[TagRule]:
    if not isinstance(overrides, dict):
        return []
    rules = overrides.get("rules")
    if not isinstance(rules, list):
        return []
    parsed: list[TagRule] = []
    for rule in rules:
        if not isinstance(rule, dict):
            continue
        pattern = str(rule.get("pattern", "")).strip()
        values = {
            field: str(rule[field]).strip()
            for field in TAG_FIELDS
            if isinstance(rule.get(field), str) and str(rule[field]).strip()
        }
        if pattern and values:
            parsed.append(TagRule(pattern, values))
    return parsed


def rule_matches(model: str, pattern: str) -> bool:
    compiled = compiled_pattern(pattern)
    return bool(compiled and compiled.search(model))


def rule_hits(pattern: str, models: list[str]) -> list[str]:
    """这条规则会命中哪些车型（用于编辑规则时预览，避免误伤）。"""
    return [model for model in models if rule_matches(model, pattern)]


def _rule_values(model: str, overrides: dict | None) -> dict[str, str]:
    """按顺序应用规则：某个字段以第一条定义了它的匹配规则为准。"""
    values: dict[str, str] = {}
    for rule in parse_rules(overrides):
        if not rule_matches(model, rule.pattern):
            continue
        for field, value in rule.values.items():
            values.setdefault(field, value)
    return values


def model_tags(model: str, overrides: dict | None = None) -> ModelTags:
    """优先级：单车覆盖 > 匹配规则 > 前缀映射 > 型号自动推断。"""
    from_rules = _rule_values(model, overrides)
    manufacturer = from_rules.get("manufacturer") or default_manufacturer(model)
    fuel = from_rules.get("fuel") or default_fuel(model)
    length = from_rules.get("length") or default_length_class(model)
    body = from_rules.get("body") or default_body(model, length)
    if overrides:
        names = overrides.get("manufacturers")
        if isinstance(names, dict) and "manufacturer" not in from_rules:
            replacement = names.get(manufacturer_key(model))
            if isinstance(replacement, str) and replacement.strip():
                manufacturer = replacement.strip()
        per_model = overrides.get("models")
        if isinstance(per_model, dict):
            entry = per_model.get(model)
            if isinstance(entry, dict):
                manufacturer = _override_value(entry, "manufacturer", manufacturer)
                fuel = _override_value(entry, "fuel", fuel)
                length = _override_value(entry, "length", length)
                body = _override_value(entry, "body", body)
    return ModelTags(manufacturer, fuel, length, body)


def _override_value(entry: dict, field: str, current: str) -> str:
    value = entry.get(field)
    return value.strip() if isinstance(value, str) and value.strip() else current


def _tag_overridden(model: str, overrides: dict | None, field: str) -> bool:
    """这个字段是不是被人手工定过（规则或单车覆盖都算）。"""
    if field in _rule_values(model, overrides):
        return True
    if isinstance(overrides, dict):
        per_model = overrides.get("models")
        if isinstance(per_model, dict):
            entry = per_model.get(model)
            if isinstance(entry, dict) and str(entry.get(field, "")).strip():
                return True
    return False


# 拼音首字母：只覆盖厂商/客车相关的常用字，够用来搜「ft → 福田」「bqt → 北汽福田」。
PINYIN_INITIALS: dict[str, str] = {
    "北": "b", "京": "j", "汽": "q", "福": "f", "田": "t", "客": "k", "车": "c", "厂": "c",
    "安": "a", "凯": "k", "宇": "y", "通": "t", "金": "j", "龙": "l", "厦": "x", "门": "m",
    "中": "z", "苏": "s", "州": "z", "海": "h", "格": "g", "聊": "l", "城": "c",
    "青": "q", "年": "n", "珠": "z", "广": "g", "一": "y", "比": "b", "亚": "y", "迪": "d",
    "黄": "h", "申": "s", "沃": "w", "银": "y", "隆": "l", "蜀": "s", "都": "d", "恒": "h",
    "大": "d", "马": "m", "九": "j", "江": "j", "淮": "h", "合": "h", "丰": "f", "吉": "j",
    "利": "l", "远": "y", "程": "c", "丰": "f", "国": "g", "际": "j", "五": "w", "洲": "z",
    "双": "s", "层": "c", "铰": "j", "接": "j", "单": "d", "机": "j", "无": "w", "轨": "g",
    "电": "d", "燃": "r", "料": "l", "池": "c", "混": "h", "动": "d", "插": "c", "增": "z",
    "天": "t", "然": "r", "气": "q", "米": "m", "级": "j", "柴": "c", "油": "y", "未": "w",
    "标": "b", "注": "z", "全": "q", "部": "b", "型": "x", "号": "h", "段": "d", "资": "z",
    "料": "l", "图": "t", "片": "p", "简": "j", "介": "j", "参": "c", "数": "s", "线": "x",
    "路": "l", "地": "d", "铁": "t", "公": "g", "交": "j", "运": "y", "转": "z", "记": "j",
    "录": "l", "时": "s", "间": "j", "备": "b", "注": "z", "搜": "s", "索": "s", "查": "c",
    "询": "x", "编": "b", "号": "h", "车": "c", "型": "x", "厂": "c", "商": "s", "长": "c",
}


def pinyin_initials(text: str) -> str:
    """把中文串转成拼音首字母（不认识的字跳过）。"""
    return "".join(PINYIN_INITIALS[ch] for ch in text if ch in PINYIN_INITIALS)


def _is_subsequence(needle: str, haystack: str) -> bool:
    """needle 的字符是否按顺序出现在 haystack 里（"bqt" 能命中 "bqft"）。"""
    iterator = iter(haystack)
    return all(character in iterator for character in needle)


def search_key(text: str) -> str:
    """搜索用的一串：原文 + 拼音首字母，便于「ft / bqft / 福田」都能命中。"""
    return f"{text.casefold()} {pinyin_initials(text)}"


def search_hit(query: str, *values: object) -> bool:
    """query 是否命中任意一个值的原文、拼音首字母（支持跳字母）。"""
    needle = query.strip().casefold()
    if not needle:
        return True
    for value in values:
        if isinstance(value, (list, tuple)):
            if any(search_hit(needle, item) for item in value):
                return True
            continue
        text = str(value)
        if needle in text.casefold():
            return True
        initials = pinyin_initials(text)
        if initials and (needle in initials or _is_subsequence(needle, initials)):
            return True
    return False


def tag_suggestions(ranges: list[BusRange], overrides: dict | None = None) -> dict[str, list[str]]:
    """下拉框里给的候选值：现有数据里出现过的 + 内置常见值。"""
    tags = [model_tags(name, overrides) for name in model_names(ranges)]
    return {
        "manufacturer": sorted({tag.manufacturer for tag in tags} | set(MANUFACTURER_NAMES.values()) | {UNKNOWN_TAG}),
        "fuel": sorted({tag.fuel for tag in tags} | {label for label, _pattern in FUEL_RULES} | {UNKNOWN_TAG}),
        "length": sorted({tag.length for tag in tags} | {f"{metres} 米级" for metres in range(6, 19)} | {UNKNOWN_TAG}),
        "body": list(BODY_VALUES),
    }


@dataclass(frozen=True)
class HealthItem:
    """一类待办：标题、处理建议，以及具体的对象列表。"""

    key: str
    title: str
    hint: str
    rows: list[tuple[str, str]]


def health_report(
    base_directory: Path,
    ranges: list[BusRange],
    profiles: dict[str, dict[str, str]],
    overrides: dict | None = None,
    metro_ranges: list[MetroRange] | None = None,
    lines: list[LineLink] | None = None,
    line_info: dict | None = None,
) -> list[HealthItem]:
    """体检报告：把「还缺什么」整理成待办清单（公交 + 地铁）。"""
    names = model_names(ranges)
    tags = {name: model_tags(name, overrides) for name in names}
    image_names = existing_image_names(base_directory)
    items: list[HealthItem] = []

    unconfirmed = [
        name for name in names if _manufacturer_unconfirmed(tags[name].manufacturer, name)
    ]
    items.append(
        HealthItem(
            "manufacturer",
            "厂商待确认",
            "在车型详情里点「编辑标签」，勾上「同一前缀」可一次改完这个厂商的所有车型。",
            [(name, f"当前显示：{tags[name].manufacturer}") for name in unconfirmed],
        )
    )

    missing_fuel = [name for name in names if tags[name].fuel == UNKNOWN_TAG]
    items.append(
        HealthItem(
            "fuel",
            "燃料未标注",
            "型号里没有新能源关键字，多半是柴油车；可在目录里多选后批量设置，"
            "也可以点下面的「燃料待确认」按查证过的来源和已有标注一次定一批。",
            [(name, f"{tags[name].manufacturer} · {tags[name].length}") for name in missing_fuel],
        )
    )

    suggested_fuel, unsure_fuel = suggest_tags_for_models(overrides, "fuel", missing_fuel)
    items.append(
        HealthItem(
            "fuel_guess",
            "燃料待确认（有依据）",
            "这些都是有出处的：要么查证过公开的车辆公告/产品页（带链接），"
            "要么是从你已经手工标过的车型里归纳出来的同后缀规律。点「燃料待确认」逐条勾选，"
            "写入前会自动备份，写错了可以用「撤销上次标签修改」回退。",
            [
                (item.model, f"{item.value or '（查到但需你选）'} · {item.source}｜{item.evidence()}")
                for item in suggested_fuel
            ],
        )
    )
    items.append(
        HealthItem(
            "fuel_manual",
            "燃料需人工判断",
            "这些型号暂时没有可靠依据（公开来源查不到、你的标注里也没有同后缀的例子）。"
            "在「燃料待确认」里按后缀分组，一次给一组定一个值最快。",
            [
                (name, f"{tags[name].manufacturer} · 后缀 {('、'.join(tag_suffix_candidates(name)) or '（无）')}")
                for name in unsure_fuel
            ],
        )
    )

    line_issues = line_discrepancies(lines or [], ranges)    # 内置清单（双层型号、查证事实）和标签文件里写的型号，是不是都能在号段表里找到
    known_models = set(names)
    curated: list[tuple[str, str]] = []
    for name in KNOWN_DOUBLE_DECKERS:
        if name not in known_models:
            curated.append((name, "内置双层清单"))
    for fact in TAG_FACTS:
        if fact.model not in known_models:
            curated.append((fact.model, f"查证事实（{fact.source}）"))
    for name in ((overrides or {}).get("models") or {}):
        if name not in known_models:
            curated.append((name, "车型标签.json 里的覆盖"))
    items.append(
        HealthItem(
            "curated",
            "清单里的型号对不上号段表",
            "内置清单（双层型号、查证过的燃料事实）或车型标签.json 里写了一个号段表里没有的型号，"
            "多半是型号名写错了一个字符——写错的清单是静默失效的，不会报错也不会生效。",
            curated,
        )
    )
    items.append(
        HealthItem(
            "line_model",
            "线路记录里的车型没收录",
            "线路数据（线路.json）里提到的车型不在号段表里，多半是型号写法不一样或写错了。"
            "在地铁/公交号段里补上，或者在「线路管理」里改掉。",
            [(line, model) for line, model in line_issues["unknown_model"]],
        )
    )
    items.append(
        HealthItem(
            "line_number",
            "线路记录里的自编号对不上",
            "同一个自编号，线路记录里写的车型和号段表里查到的不一样。",
            [(f"{line} {number}", f"记录里是 {model}，号段表里是 {table_model}") for line, number, model, table_model in line_issues["number_mismatch"]],
        )
    )

    # 线路资料（联网查的站点表）：线路数据里出现过、但还没查过站点的线路
    without_stops = line_info_lines_without_stops(line_info or {}, lines or [])
    items.append(
        HealthItem(
            "line_stops",
            "线路还没查过站点表",
            f"{LINE_INFO_FILENAME} 里没有这些线路的站点。去「线路资料（联网）」查一次并「保存站点到线路资料」，"
            "以后就能离线看这条线有什么站。不联网也不影响（这一项可以不管）。",
            [(name, "还没查过站点") for name in without_stops],
        )
    )

    missing_length = [name for name in names if tags[name].length == UNKNOWN_TAG]
    items.append(
        HealthItem(
            "length",
            "车长未标注",
            "型号代号不符合常见规则，手动填一次即可。",
            [(name, tags[name].manufacturer) for name in missing_length],
        )
    )

    missing_body = [name for name in names if tags[name].body == UNKNOWN_TAG]
    items.append(
        HealthItem(
            "body",
            "车体形式未标注",
            "车长没推断出来，所以车体形式也定不了；填好车长后它会自动跟着出来。",
            [(name, tags[name].manufacturer) for name in missing_body],
        )
    )

    guessed_body = [
        name
        for name in names
        if tags[name].body == "双层"
        and _tag_overridden(name, overrides, "body") is False
        and body_is_guessed_double_decker(name)
    ]
    items.append(
        HealthItem(
            "body_guess",
            "车体形式待确认",
            "这几台是按内置清单判成「双层」的：对就选中后点上面的「确认」按钮（确认过就不再提示），"
            "不对就在详情页「编辑标签」里改。判定双层不看型号字母（DNC6107BEVGS1 名字里有 GS，其实是低入口单机），"
            "所以清单之外的双层车请在目录里多选后「批量设置标签」标一次。"
            "另外 16/18 米级按铰接、13/14 米级按三轴算，可用「车体」筛选抽查。",
            [(name, f"按内置清单判为双层：{name}") for name in guessed_body],
        )
    )

    without_profile = [name for name in names if profile_is_empty(profiles.get(name, {}), None, image_names)]
    items.append(
        HealthItem(
            "profile",
            "还没有资料",
            "在车型详情里补简介/参数/主图；有主图的会在这里显示「仅有主图」。",
            [
                (name, "仅有主图" if (profiles.get(name, {}).get("image") or "").strip() else "完全没有")
                for name in without_profile
            ],
        )
    )

    broken_images = [
        (name, profile.get("image", ""))
        for name, profile in profiles.items()
        if name in set(names)
        and (profile.get("image") or "").strip()
        and Path(profile["image"]).name not in image_names
    ]
    items.append(
        HealthItem("image", "主图文件丢失", "资料里记着图片名，但 车型图片 文件夹里找不到，重新上传或移除即可。", broken_images)
    )

    # 号段重叠：同一个自编号落在两个号段里（数据错），查车时只会匹配到先遇到的那个
    overlaps: list[tuple[str, str]] = []
    ordered = sorted(ranges, key=lambda item: (item.start, item.end))
    for index, current in enumerate(ordered):
        for other in ordered[index + 1:]:
            if other.start > current.end:
                break
            shared_start = max(current.start, other.start)
            shared_end = min(current.end, other.end)
            count = shared_end - shared_start + 1
            same_model = current.model == other.model
            note = f"重叠 {count} 个自编号（{shared_start}-{shared_end}）"
            note += "，两边车型一样" if same_model else f"，车型不同：{current.model} / {other.model}"
            overlaps.append((f"{current.start}-{current.end} 与 {other.start}-{other.end}", note))
    items.append(
        HealthItem(
            "overlap",
            "号段重叠",
            "同一个自编号被两个号段包住了——查这个号会按先遇到的号段回答，等于有一半是错的。"
            "进「管理号段」把重复的那段删掉或改掉；如果确实是两台车共用一个号，那号段表本身需要拆开。",
            overlaps,
        )
    )

    gaps = [(previous.end + 1, current.start - 1) for previous, current in zip(ranges, ranges[1:]) if current.start - previous.end > 1]
    items.append(
        HealthItem(
            "gap",
            "号段空档",
            "相邻号段之间没有收录的自编号区间；如果这些号段确实不存在，忽略即可。",
            [(f"{start}-{end}", f"{end - start + 1} 个自编号未收录") for start, end in gaps],
        )
    )
    if metro_ranges:
        items.extend(metro_health_items(base_directory, metro_ranges))
    return items


def profile_path(base_directory: Path) -> Path:
    return base_directory / PROFILE_FILENAME


def image_directory(base_directory: Path) -> Path:
    return base_directory / IMAGE_DIRECTORY_NAME


def load_profiles(base_directory: Path) -> dict[str, dict[str, object]]:
    path = profile_path(base_directory)
    if not path.exists():
        return {}
    try:
        contents = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ProfileError(f"无法读取 {PROFILE_FILENAME}。请确认文件未损坏。") from error
    if not isinstance(contents, dict):
        raise ProfileError(f"{PROFILE_FILENAME} 格式无效。")
    profiles: dict[str, dict[str, object]] = {}
    for model, profile in contents.items():
        if not (isinstance(model, str) and isinstance(profile, dict)):
            continue
        images = profile_images(profile)
        profiles[model] = {
            "introduction": str(profile.get("introduction", "")),
            "specifications": str(profile.get("specifications", "")),
            # image 始终保留为封面，兼容旧版本与体检逻辑
            "image": images[0]["name"] if images else "",
            "images": images,
        }
    return profiles


def profile_images(profile: dict) -> list[dict[str, str]]:
    """统一读出图片列表：新的 images 优先，旧结构 image 自动迁移成单张。"""
    images = profile.get("images")
    if isinstance(images, list):
        cleaned: list[dict[str, str]] = []
        for item in images:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            if name:
                cleaned.append({"name": name, "caption": str(item.get("caption", "")).strip()})
        if cleaned:
            return cleaned
    single = str(profile.get("image", "")).strip()
    return [{"name": single, "caption": ""}] if single else []


def set_profile_images(profile: dict, images: list[dict[str, str]]) -> None:
    """写入图片列表，同时把第一张同步到 image 字段（兼容旧结构）。"""
    cleaned = [item for item in images if str(item.get("name", "")).strip()]
    profile["images"] = cleaned
    profile["image"] = cleaned[0]["name"] if cleaned else ""


def save_profiles(base_directory: Path, profiles: dict[str, dict[str, object]]) -> None:
    path = profile_path(base_directory)
    temporary_path = path.with_suffix(".tmp")
    try:
        backup_file(path)
        temporary_path.write_text(json.dumps(profiles, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary_path.replace(path)
    except OSError as error:
        temporary_path.unlink(missing_ok=True)
        raise ProfileError(f"无法保存 {PROFILE_FILENAME}。请确认程序文件夹可写入。") from error


def _safe_image_name(model: str, index: int | None = None) -> str:
    safe = "".join(character if character.isalnum() else "_" for character in model)
    if index is None:
        return f"{safe[:80]}.jpg"
    return f"{safe[:72]}-{index}.jpg"


def next_image_name(base_directory: Path, model: str) -> str:
    """给相册里新加的一张图挑一个没被占用的文件名。"""
    used = existing_image_names(base_directory)
    for index in range(1, 100):
        candidate = _safe_image_name(model, index)
        if candidate not in used:
            return candidate
    return f"{_safe_image_name(model, 99)[:-4]}-{datetime.now().strftime('%H%M%S')}.jpg"


def stored_image_path(base_directory: Path, image_name: str) -> Path:
    """Resolve a stored image name inside the image folder only."""
    return image_directory(base_directory) / Path(image_name).name


def photo_directory(base_directory: Path) -> Path:
    return base_directory / PHOTO_DIRECTORY_NAME


def stored_photo_path(base_directory: Path, photo_name: str) -> Path:
    """拍车日志里用户自己拍的照片（只认文件名，不允许路径穿越）。"""
    return photo_directory(base_directory) / Path(photo_name).name


def save_spot_photo(base_directory: Path, source: Path | Image.Image) -> str:
    """把一张照片压进 拍车照片 文件夹，返回文件名。"""
    image = source if isinstance(source, Image.Image) else load_image(source)
    if image.mode != "RGB":
        image = image.convert("RGB")
    directory = photo_directory(base_directory)
    directory.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    name = f"拍车-{stamp}.jpg"
    index = 1
    while (directory / name).exists():
        index += 1
        name = f"拍车-{stamp}-{index}.jpg"
    try:
        image.thumbnail(MAX_IMAGE_PIXELS, Image.Resampling.LANCZOS)
        image.save(directory / name, format="JPEG", quality=JPEG_QUALITY, optimize=True)
    except OSError as error:
        raise ProfileError("无法保存照片，请确认程序文件夹可写入。") from error
    return name


def delete_spot_photo(base_directory: Path, photo_name: str) -> None:
    if not photo_name:
        return
    try:
        stored_photo_path(base_directory, photo_name).unlink(missing_ok=True)
    except OSError as error:
        raise ProfileError("无法删除照片，请确认程序文件夹可写入。") from error


def load_image(source: Path) -> Image.Image:
    """Open and fully decode an image, converting it to RGB."""
    if source.suffix.lower() not in IMAGE_EXTENSIONS:
        raise ProfileError("请选择 JPG、PNG、BMP、GIF 或 WEBP 格式的图片。")
    try:
        with Image.open(source) as image:
            image.load()
            return image.convert("RGB")
    except (OSError, UnidentifiedImageError) as error:
        raise ProfileError("无法读取该图片，请选择未损坏的常见图片文件。") from error


def import_image(
    base_directory: Path,
    model: str,
    source: Path | Image.Image,
    previous_image: str = "",
    filename: str | None = None,
    *,
    delete_previous: bool = True,
) -> str:
    """把一张图压缩后存进 车型图片，返回文件名。

    source 可以是文件路径，也可以是已经读进来的 PIL 图片（用于剪贴板贴图）。
    filename 为空时用「车型.jpg」这个封面名；相册里的第 N 张用 next_image_name()。
    """
    image = source if isinstance(source, Image.Image) else load_image(source)
    if image.mode != "RGB":
        image = image.convert("RGB")
    destination_directory = image_directory(base_directory)
    name = filename or _safe_image_name(model)
    destination = destination_directory / name
    temporary = destination.with_suffix(".tmp")
    try:
        destination_directory.mkdir(exist_ok=True)
        image.thumbnail(MAX_IMAGE_PIXELS, Image.Resampling.LANCZOS)
        image.save(temporary, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        temporary.replace(destination)
        if delete_previous and previous_image and previous_image != name:
            stored_image_path(base_directory, previous_image).unlink(missing_ok=True)
        return name
    except OSError as error:
        temporary.unlink(missing_ok=True)
        raise ProfileError("无法保存图片，请确认程序文件夹可写入。") from error


def delete_image(base_directory: Path, image_name: str) -> None:
    if not image_name:
        return
    try:
        stored_image_path(base_directory, image_name).unlink(missing_ok=True)
    except OSError as error:
        raise ProfileError("无法删除旧图片，请确认程序文件夹可写入。") from error


def existing_image_names(base_directory: Path) -> set[str]:
    """一次列举图片目录，避免逐行做磁盘检查。"""
    directory = image_directory(base_directory)
    try:
        return {item.name for item in directory.iterdir() if item.is_file()}
    except OSError:
        return set()


def profile_is_empty(
    profile: dict[str, str],
    base_directory: Path | None = None,
    existing_images: set[str] | None = None,
) -> bool:
    """Report whether a profile has neither text nor a usable main image.

    给了 existing_images 就查这一份名单（批量渲染时用，避免反复 stat）。
    """
    if profile.get("introduction", "").strip() or profile.get("specifications", "").strip():
        return False
    image_name = profile.get("image", "")
    if not image_name:
        return True
    if existing_images is not None:
        return Path(image_name).name not in existing_images
    if base_directory is None:
        return False
    return not stored_image_path(base_directory, image_name).is_file()


def metro_data_path(base_directory: Path) -> Path:
    return base_directory / METRO_DATA_FILENAME


def _metro_sort_key(number: str) -> tuple[str, int]:
    match = re.fullmatch(r"([A-Z]*)(\d+)", number.upper())
    if match is None:
        raise MetroError(f"地铁编号“{number}”格式无效。")
    return match.group(1), int(match.group(2))


def _metro_bounds_compatible(start: str, end: str) -> bool:
    return _metro_sort_key(start)[0] == _metro_sort_key(end)[0]


METRO_RANGE_PATTERN = re.compile(r"([A-Za-z]*\d+)-([A-Za-z]*\d+)\s+(.+)")
METRO_RANGE_LIKE_PATTERN = re.compile(r"[A-Za-z]*\d+\s*-\s*[A-Za-z]*\d+")
METRO_SPACED_RANGE_PATTERN = re.compile(r"[A-Za-z]*\d+\s+[A-Za-z]*\d+(\s|$)")


def parse_metro_source(path: Path) -> list[MetroRange]:
    """Read the plain-text seed file, rejecting malformed range lines loudly."""
    if not path.is_file():
        raise MetroError(f"未找到 {METRO_SOURCE_FILENAME}。")
    try:
        rows = path.read_text(encoding="utf-8-sig").splitlines()
    except OSError as error:
        raise MetroError("无法读取地铁车型数据文件。") from error
    line = ""
    ranges: list[MetroRange] = []
    for line_number, raw in enumerate(rows, start=1):
        text = raw.strip()
        if not text:
            continue
        match = METRO_RANGE_PATTERN.fullmatch(text)
        if match is None:
            if METRO_RANGE_LIKE_PATTERN.search(text):
                raise MetroError(f"地铁数据第 {line_number} 行号段格式无效，应形如“01001-01031 DKZ4”（号段与车型之间留一个空格）。")
            if METRO_SPACED_RANGE_PATTERN.search(text):
                raise MetroError(f"地铁数据第 {line_number} 行号段格式无效，应形如“01001-01031 DKZ4”（号段之间用半角连字符）。")
            line = text
            continue
        if not line:
            raise MetroError(f"地铁数据第 {line_number} 行缺少线路名称。")
        start, end, model = match.groups()
        if not _metro_bounds_compatible(start, end) or _metro_sort_key(start)[1] > _metro_sort_key(end)[1]:
            raise MetroError(f"地铁数据第 {line_number} 行号段无效。")
        ranges.append(MetroRange(line, start.upper(), end.upper(), model.strip()))
    if not ranges:
        raise MetroError("地铁车型数据中没有可用号段。")
    return ranges


def load_metro_ranges(base_directory: Path) -> list[MetroRange]:
    """Load saved metro ranges, seeding from the text file on first run.

    A read-only program folder must not disable metro lookups: the parsed seed
    data is still returned when it cannot be persisted.
    """
    data_file = metro_data_path(base_directory)
    if not data_file.exists():
        ranges = parse_metro_source(base_directory / METRO_SOURCE_FILENAME)
        try:
            save_metro_ranges(base_directory, ranges)
        except MetroError:
            pass
        return ranges
    try:
        content = json.loads(data_file.read_text(encoding="utf-8"))
        return [MetroRange(**item) for item in content]
    except (OSError, json.JSONDecodeError, TypeError) as error:
        raise MetroError(f"无法读取 {METRO_DATA_FILENAME}。") from error


def save_metro_ranges(base_directory: Path, ranges: list[MetroRange]) -> None:
    path = metro_data_path(base_directory)
    temporary = path.with_suffix(".tmp")
    try:
        backup_file(path)
        temporary.write_text(json.dumps([item.__dict__ for item in ranges], ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(path)
    except OSError as error:
        temporary.unlink(missing_ok=True)
        raise MetroError(f"无法保存 {METRO_DATA_FILENAME}，请确认程序文件夹可写入。") from error


def metro_lines(ranges: list[MetroRange]) -> list[str]:
    return sorted({item.line for item in ranges}, key=str.casefold)


def parse_metro_number_list(text: str) -> tuple[list[str], list[str]]:
    """Split pasted metro numbers into unique uppercase values plus ignored parts."""
    numbers: list[str] = []
    ignored: list[str] = []
    seen: set[str] = set()
    for token in re.split(r"[\s,，、;；|/\\]+", text.strip()):
        if not token:
            continue
        try:
            _metro_sort_key(token)
        except MetroError:
            ignored.append(token)
            continue
        value = token.upper()
        if value not in seen:
            seen.add(value)
            numbers.append(value)
    return numbers, ignored


def import_metro_source(base_directory: Path) -> list[MetroRange]:
    """用 txt 的内容覆盖当前地铁数据（保存前自动备份）。"""
    ranges = parse_metro_source(base_directory / METRO_SOURCE_FILENAME)
    save_metro_ranges(base_directory, ranges)
    return ranges


def export_metro_source(base_directory: Path, ranges: list[MetroRange]) -> Path:
    """把当前地铁数据写回 txt 种子文件，保持两边一致。"""
    path = base_directory / METRO_SOURCE_FILENAME
    blocks: list[str] = []
    for line in metro_lines(ranges):
        entries = [f"{item.start}-{item.end} {item.model}" for item in metro_ranges_for_line(line, ranges)]
        blocks.append("\n".join([line] + entries))
    text = "\n\n".join(blocks) + "\n"
    temporary = path.with_suffix(".tmp")
    try:
        temporary.write_text(text, encoding="utf-8")
        temporary.replace(path)
    except OSError as error:
        temporary.unlink(missing_ok=True)
        raise MetroError(f"无法写入 {METRO_SOURCE_FILENAME}，请确认程序文件夹可写入。") from error
    return path


def metro_source_diff(base_directory: Path, current: list[MetroRange] | None = None) -> tuple[list[MetroRange], list[MetroRange]]:
    """比较 txt 种子与当前数据，返回（只在 txt 里的, 只在数据里的）。"""
    source_path = base_directory / METRO_SOURCE_FILENAME
    if not source_path.is_file():
        return [], []
    try:
        seeded = parse_metro_source(source_path)
    except MetroError:
        return [], []
    if current is None:
        current = load_metro_ranges(base_directory)
    key = lambda item: (item.line, item.start, item.end, item.model)  # noqa: E731
    only_source = sorted(set(seeded) - set(current), key=key)
    only_current = sorted(set(current) - set(seeded), key=key)
    return only_source, only_current


def metro_health_items(base_directory: Path, ranges: list[MetroRange]) -> list["HealthItem"]:
    """地铁侧体检：种子漂移、号段重叠、同线路空档。"""
    items: list[HealthItem] = []

    only_source, only_current = metro_source_diff(base_directory, ranges)
    drift_rows = [(f"{METRO_SOURCE_FILENAME}：{item.line} {item.start}-{item.end} {item.model}", "只在种子文件里") for item in only_source]
    drift_rows += [(f"当前数据：{item.line} {item.start}-{item.end} {item.model}", "只在当前数据里") for item in only_current]
    items.append(
        HealthItem(
            "metro_source",
            "地铁 · 种子文件不一致",
            f"{METRO_SOURCE_FILENAME} 只是首次导入用的种子，现在两边内容不一样了。二选一："
            "要么以种子为准（「地铁数据工具」→ 重新从 txt 导入，会覆盖当前数据），"
            "要么以你现在的数据为准（「地铁数据工具」→ 把当前数据写回 txt 种子）。"
            "点上面的按钮直接过去。",
            drift_rows,
        )
    )

    overlaps: list[tuple[str, str]] = []
    gaps: list[tuple[str, str]] = []
    for line in metro_lines(ranges):
        groups: dict[str, list[MetroRange]] = {}
        for item in metro_ranges_for_line(line, ranges):
            groups.setdefault(_metro_sort_key(item.start)[0], []).append(item)
        for entries in groups.values():
            entries.sort(key=lambda item: _metro_sort_key(item.start)[1])
            for previous, current in zip(entries, entries[1:]):
                previous_end = _metro_sort_key(previous.end)[1]
                current_start = _metro_sort_key(current.start)[1]
                if current_start <= previous_end:
                    overlaps.append((f"{line} {previous.start}-{previous.end} 与 {current.start}-{current.end}", "号段重叠"))
                elif current_start - previous_end > 1:
                    missing = current_start - previous_end - 1
                    note = f"{missing} 个编号未收录"
                    if missing > 50:
                        note += "（跨度很大，通常是两批不同的编号，可不理会）"
                    gaps.append((f"{line} {previous.end} → {current.start}", note))
    items.append(
        HealthItem("metro_overlap", "地铁 · 号段重叠", "同一条线路、相同字母前缀的号段不该重叠——点上面的按钮进「管理地铁号段与车型」，删掉或改掉其中一条。", overlaps)
    )
    items.append(
        HealthItem(
            "metro_gap",
            "地铁 · 号段空档",
            "两个号段之间还有编号没收录（地铁车辆分批采购，本来就可能不连续）。"
            "确认确实有车 → 点上面的按钮进「管理地铁号段与车型」补上这段；确实没有 → 忽略即可。",
            gaps,
        )
    )
    return items


def metro_ranges_for_line(line: str, ranges: list[MetroRange]) -> list[MetroRange]:
    return sorted((item for item in ranges if item.line == line), key=lambda item: _metro_sort_key(item.start))


def find_metro_models(number: str, ranges: list[MetroRange]) -> list[MetroRange]:
    number = number.strip().upper()
    prefix, value = _metro_sort_key(number)
    matches = [
        item
        for item in ranges
        if _metro_sort_key(item.start)[0] == prefix and _metro_sort_key(item.start)[1] <= value <= _metro_sort_key(item.end)[1]
    ]
    return sorted(matches, key=lambda item: (item.line, _metro_sort_key(item.start)))


def validate_metro_range(line: str, start: str, end: str, model: str, ranges: list[MetroRange], exclude_index: int | None = None) -> MetroRange:
    line, start, end, model = line.strip(), start.strip().upper(), end.strip().upper(), model.strip()
    if not line or not model:
        raise MetroError("请输入线路名称和车型。")
    if not _metro_bounds_compatible(start, end):
        raise MetroError("号段起始和结束编号的字母前缀必须相同。")
    start_key, end_key = _metro_sort_key(start), _metro_sort_key(end)
    if start_key[1] > end_key[1]:
        raise MetroError("号段起始不能大于号段结束。")
    for index, current in enumerate(ranges):
        if index == exclude_index or current.line != line:
            continue
        current_start, current_end = _metro_sort_key(current.start), _metro_sort_key(current.end)
        if start_key[0] == current_start[0] and start_key[1] <= current_end[1] and end_key[1] >= current_start[1]:
            raise MetroError(f"号段与 {current.start}-{current.end} 重复，无法保存。")
    return MetroRange(line, start, end, model)


# --------------------------------------------------------------------------- #
# Theme and shared user-interface helpers
# --------------------------------------------------------------------------- #


def enable_dpi_awareness() -> None:
    """Ask Windows for real pixels so text stays crisp on scaled displays."""
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class Theme:
    """Colours, metrics and ttk styles shared by every window."""

    # 中性偏冷的底色 + 一层层变亮的白，让卡片和背景拉开层次
    background = "#F0F3F8"
    surface = "#FFFFFF"
    surface_alt = "#F5F8FC"
    header = "#123A66"
    header_soft = "#A9C6E4"
    header_accent = "#2E7BD6"
    primary = "#1668C4"
    primary_active = "#12559F"
    primary_soft = "#E8F1FC"
    primary_disabled = "#A7C0DC"
    accent = "#C92936"
    accent_active = "#A6202C"
    accent_soft = "#FBEBED"
    text = "#16212E"
    muted = "#6B7B8F"
    border = "#DCE3EC"
    border_strong = "#C9D4E2"
    stripe = "#F7FAFD"
    hover = "#EEF5FE"
    selection = "#DCE9FA"
    heading_bg = "#E9EFF7"
    heading_bg_active = "#DFE8F4"
    ok = "#1E7F4F"
    success = ok
    success_soft = "#E7F4EC"
    danger = accent
    warning = "#B26A00"
    warning_soft = "#FBF2E2"
    divider = "#E7ECF3"
    mono_family = "Consolas"
    row_height = 26
    heading_height = 30

    # 统一的间距尺度（px 会被 DPI 缩放）
    SPACE = (2, 4, 6, 8, 10, 12, 14, 16, 20, 24)

    def __init__(self, root: tk.Misc) -> None:
        self.root = root
        try:
            self.scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        except tk.TclError:
            self.scale = 1.0
        self.family = self._pick_family()
        self.style = ttk.Style(root)
        self._configure_fonts()
        self._configure_styles()

    # -- metrics --------------------------------------------------------- #
    def px(self, value: float) -> int:
        return int(round(value * self.scale))

    def space(self, step: int = 4) -> int:
        """间距尺度：space(1)=2px … space(9)=24px。"""
        index = max(0, min(len(self.SPACE) - 1, step))
        return self.px(self.SPACE[index])

    def font(self, size: int, weight: str = "normal") -> tuple[str, int] | tuple[str, int, str]:
        return (self.family, size) if weight == "normal" else (self.family, size, weight)

    def _pick_family(self) -> str:
        try:
            families = set(tkfont.families(self.root))
        except tk.TclError:
            return "TkDefaultFont"
        for candidate in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "Arial"):
            if candidate in families:
                return candidate
        return "TkDefaultFont"

    def _configure_fonts(self) -> None:
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkTooltipFont"):
            try:
                tkfont.nametofont(name).configure(family=self.family, size=10)
            except tk.TclError:
                pass
        self.root.option_add("*Font", self.font(10))

    # -- widgets --------------------------------------------------------- #
    def card(self, parent: tk.Misc, *, accent: bool = False) -> tk.Frame:
        """A flat white panel with a one-pixel border；accent=True 时顶部加一条主色横条。"""
        outer = tk.Frame(
            parent,
            background=self.surface,
            highlightthickness=1,
            highlightbackground=self.border,
            highlightcolor=self.border,
            bd=0,
        )
        if accent:
            tk.Frame(outer, background=self.primary, height=self.px(3)).pack(side="top", fill="x")
        return outer

    def section(
        self,
        parent: tk.Misc,
        text: str,
        *,
        background: str | None = None,
        pady: tuple[int, int] = (0, 6),
        grid: tuple[int, int, int] | None = None,
    ) -> tk.Frame:
        """分区标题：左侧一条主色小竖条 + 加粗标题。

        以前各窗口各写各的 Section.TLabel，这里统一成一个看得见的层次。
        `grid=(row, column, columnspan)` 时用 grid 摆放（在不允许 pack 的容器里用）。
        """
        holder = tk.Frame(parent, background=background or self.surface)
        if grid is None:
            holder.pack(anchor="w", fill="x", pady=(self.px(pady[0]), self.px(pady[1])))
        else:
            row, column, columnspan = grid
            holder.grid(row=row, column=column, columnspan=columnspan, sticky="ew",
                        pady=(self.px(pady[0]), self.px(pady[1])))
        tk.Frame(holder, background=self.primary, width=self.px(3), height=self.px(13)).pack(side="left", padx=(0, self.px(7)))
        tk.Label(
            holder,
            text=text,
            background=background or self.surface,
            foreground=self.text,
            font=self.font(10, "bold"),
        ).pack(side="left")
        return holder

    def mono_font(self, size: int = 10, weight: str = "normal") -> tuple[str, int] | tuple[str, int, str]:
        """编号/号段这类要竖着对齐的数字用等宽字体。"""
        return (self.mono_family, size) if weight == "normal" else (self.mono_family, size, weight)

    def text_box(self, parent: tk.Misc, height: int) -> tk.Text:
        return tk.Text(
            parent,
            height=height,
            wrap="word",
            font=self.font(10),
            background=self.surface,
            foreground=self.text,
            insertbackground=self.text,
            selectbackground=self.selection,
            selectforeground=self.text,
            relief="flat",
            highlightthickness=1,
            highlightbackground=self.border,
            highlightcolor=self.primary,
            padx=self.px(8),
            pady=self.px(6),
            spacing1=self.px(2),
        )

    # -- ttk styles ------------------------------------------------------ #
    def _configure_styles(self) -> None:
        style = self.style
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        pad = self.px
        style.configure(".", background=self.background, foreground=self.text, font=self.font(10), focuscolor=self.primary)
        style.configure("TFrame", background=self.background)
        style.configure("Card.TFrame", background=self.surface)
        style.configure("Header.TFrame", background=self.header)
        style.configure("TLabel", background=self.background, foreground=self.text)
        style.configure("Card.TLabel", background=self.surface, foreground=self.text)
        style.configure("HeaderTitle.TLabel", background=self.header, foreground="#FFFFFF", font=self.font(17, "bold"))
        style.configure("HeaderInfo.TLabel", background=self.header, foreground=self.header_soft, font=self.font(9))
        style.configure("Section.TLabel", background=self.surface, foreground=self.muted, font=self.font(9, "bold"))
        style.configure("CardValue.TLabel", background=self.surface, foreground=self.text, font=self.font(10))
        style.configure("CardStrong.TLabel", background=self.surface, foreground=self.text, font=self.font(12, "bold"))
        style.configure("Muted.TLabel", background=self.background, foreground=self.muted, font=self.font(9))
        style.configure("CardMuted.TLabel", background=self.surface, foreground=self.muted, font=self.font(9))
        style.configure("Result.TLabel", background=self.surface, foreground=self.primary, font=self.font(22, "bold"))
        style.configure(
            "Chip.TLabel",
            background=self.primary_soft,
            foreground=self.primary,
            font=self.font(9, "bold"),
            padding=(pad(8), pad(3)),
        )
        style.configure("DialogTitle.TLabel", background=self.background, foreground=self.text, font=self.font(13, "bold"))
        style.configure("TButton", font=self.font(10))
        style.configure("TSeparator", background=self.border)

        style.configure(
            "Primary.TButton",
            background=self.primary,
            foreground="#FFFFFF",
            font=self.font(10, "bold"),
            borderwidth=0,
            relief="flat",
            padding=(pad(18), pad(9)),
            anchor="center",
            lightcolor=self.primary,
            darkcolor=self.primary,
            bordercolor=self.primary,
            focuscolor=self.primary,
        )
        style.map(
            "Primary.TButton",
            background=[("disabled", self.primary_disabled), ("pressed", self.primary_active), ("active", self.primary_active)],
            foreground=[("disabled", "#F0F5FA")],
            lightcolor=[("pressed", self.primary_active), ("active", self.primary_active)],
            darkcolor=[("pressed", self.primary_active), ("active", self.primary_active)],
            bordercolor=[("pressed", self.primary_active), ("active", self.primary_active)],
        )

        style.configure(
            "Secondary.TButton",
            background=self.surface,
            foreground=self.text,
            font=self.font(10),
            borderwidth=1,
            relief="flat",
            padding=(pad(13), pad(9)),
            anchor="center",
            lightcolor=self.surface,
            darkcolor=self.surface,
            bordercolor=self.border_strong,
            focuscolor=self.border,
        )
        style.map(
            "Secondary.TButton",
            background=[("disabled", self.surface_alt), ("pressed", "#DFEAF7"), ("active", self.primary_soft)],
            foreground=[("disabled", "#AEBBC9")],
            bordercolor=[("active", self.primary), ("pressed", self.primary)],
            lightcolor=[("active", self.primary_soft), ("pressed", "#DFEAF7")],
            darkcolor=[("active", self.primary_soft), ("pressed", "#DFEAF7")],
        )

        style.configure(
            "Danger.TButton",
            background=self.surface,
            foreground=self.accent,
            font=self.font(10),
            borderwidth=1,
            relief="flat",
            padding=(pad(13), pad(9)),
            anchor="center",
            lightcolor=self.surface,
            darkcolor=self.surface,
            bordercolor=self.border_strong,
            focuscolor=self.border,
        )
        style.map(
            "Danger.TButton",
            background=[("disabled", self.surface_alt), ("pressed", "#F6D9DC"), ("active", self.accent_soft)],
            foreground=[("disabled", "#D6A7AC")],
            bordercolor=[("active", self.accent), ("pressed", self.accent)],
            lightcolor=[("active", self.accent_soft), ("pressed", "#F6D9DC")],
            darkcolor=[("active", self.accent_soft), ("pressed", "#F6D9DC")],
        )
        # 第四级：Ghost（无边框文字按钮），给"重置筛选"这类次要动作，不跟主操作抢注意力
        style.configure(
            "Ghost.TButton",
            background=self.surface,
            foreground=self.primary,
            font=self.font(9),
            borderwidth=0,
            relief="flat",
            padding=(pad(7), pad(3)),
            anchor="center",
        )
        style.map(
            "Ghost.TButton",
            background=[("disabled", self.surface), ("pressed", self.primary_soft), ("active", self.primary_soft)],
            foreground=[("disabled", self.muted), ("active", self.primary_active)],
        )

        style.configure(
            "TEntry",
            fieldbackground=self.surface,
            background=self.surface,
            foreground=self.text,
            insertcolor=self.text,
            bordercolor=self.border_strong,
            lightcolor=self.border_strong,
            darkcolor=self.border_strong,
            relief="flat",
            padding=pad(7),
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", self.primary)],
            lightcolor=[("focus", self.primary)],
            darkcolor=[("focus", self.primary)],
        )
        style.configure("Query.TEntry", font=self.font(15, "bold"), padding=pad(9), justify="center")
        style.configure("Search.TEntry", padding=pad(6))

        style.configure(
            "Treeview",
            background=self.surface,
            fieldbackground=self.surface,
            foreground=self.text,
            rowheight=pad(self.row_height),
            borderwidth=0,
            relief="flat",
            font=self.font(10),
        )
        style.map(
            "Treeview",
            background=[("selected", self.selection)],
            foreground=[("selected", self.text)],
        )
        style.configure(
            "Treeview.Heading",
            background=self.heading_bg,
            foreground=self.muted,
            font=self.font(9, "bold"),
            relief="flat",
            borderwidth=0,
            padding=(pad(7), pad((self.heading_height - 14) / 2)),
        )
        style.map("Treeview.Heading", background=[("active", self.heading_bg_active)], foreground=[("active", self.primary)])

        style.configure(
            "Vertical.TScrollbar",
            background="#C7D3E1",
            troughcolor=self.background,
            bordercolor=self.background,
            arrowcolor=self.muted,
            relief="flat",
            width=pad(11),
        )
        style.map("Vertical.TScrollbar", background=[("active", "#AEBFD2")])
        style.configure(
            "Horizontal.TScrollbar",
            background="#C7D3E1",
            troughcolor=self.background,
            bordercolor=self.background,
            arrowcolor=self.muted,
            relief="flat",
        )
        style.map("Horizontal.TScrollbar", background=[("active", "#AEBFD2")])

        style.configure(
            "TCombobox",
            fieldbackground=self.surface,
            background=self.surface,
            foreground=self.text,
            arrowcolor=self.muted,
            bordercolor=self.border,
            lightcolor=self.border,
            darkcolor=self.border,
            padding=pad(3),
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", self.surface)],
            bordercolor=[("focus", self.primary), ("hover", self.primary)],
            arrowcolor=[("active", self.primary)],
        )
        self.root.option_add("*TCombobox*Listbox.background", self.surface)
        self.root.option_add("*TCombobox*Listbox.foreground", self.text)
        self.root.option_add("*TCombobox*Listbox.selectBackground", self.selection)
        self.root.option_add("*TCombobox*Listbox.selectForeground", self.text)
        self.root.option_add("*TCombobox*Listbox.font", self.font(10))

        style.configure(
            "Card.TCheckbutton",
            background=self.surface,
            foreground=self.text,
            font=self.font(10),
            focuscolor=self.surface,
            padding=(0, pad(2)),
        )
        style.map(
            "Card.TCheckbutton",
            background=[("active", self.surface)],
            foreground=[("disabled", self.muted)],
            indicatorcolor=[("selected", self.primary), ("!selected", self.surface)],
        )
        style.configure(
            "Card.TRadiobutton",
            background=self.surface,
            foreground=self.text,
            font=self.font(10),
            focuscolor=self.surface,
            padding=(0, pad(2)),
        )
        style.map(
            "Card.TRadiobutton",
            background=[("active", self.surface)],
            indicatorcolor=[("selected", self.primary), ("!selected", self.surface)],
        )
        style.configure("Card.TSeparator", background=self.border)

        style.configure("TNotebook", background=self.background, borderwidth=0, tabmargins=(pad(2), pad(3), 0, 0))
        style.configure(
            "TNotebook.Tab",
            background="#DCE5F0",
            foreground=self.muted,
            font=self.font(10),
            padding=(pad(18), pad(8)),
            borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", self.surface), ("active", self.primary_soft)],
            foreground=[("selected", self.primary)],
            expand=[("selected", (0, 0, 0, 0))],
        )


def settings_path(base_directory: Path) -> Path:
    return base_directory / SETTINGS_FILENAME


def load_settings(base_directory: Path) -> dict:
    """读界面设置；文件缺失或损坏都退回空设置，不影响启动。"""
    path = settings_path(base_directory)
    if not path.is_file():
        return {}
    try:
        contents = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return contents if isinstance(contents, dict) else {}


def save_settings(base_directory: Path, settings: dict) -> None:
    """原子写界面设置（这份文件不值得进备份目录，失败也不打扰用户）。"""
    path = settings_path(base_directory)
    temporary = path.with_suffix(".tmp")
    try:
        temporary.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(path)
    except OSError:
        temporary.unlink(missing_ok=True)


def stored_geometry(base_directory: Path, key: str) -> tuple[int, int, int, int] | None:
    windows = load_settings(base_directory).get("windows")
    if not isinstance(windows, dict):
        return None
    value = windows.get(key)
    if isinstance(value, list) and len(value) == 4 and all(isinstance(item, int) for item in value):
        return value[0], value[1], value[2], value[3]
    return None


def remember_geometry(window: tk.Misc, base_directory: Path, key: str, delay_ms: int = 600) -> None:
    """窗口移动/缩放后（去抖）把位置与大小记进 界面设置.json。"""

    def store() -> None:
        try:
            if not window.winfo_exists() or window.state() != "normal":
                return
            geometry = (window.winfo_x(), window.winfo_y(), window.winfo_width(), window.winfo_height())
        except tk.TclError:
            return
        settings = load_settings(base_directory)
        windows = settings.get("windows")
        if not isinstance(windows, dict):
            windows = {}
        windows[key] = list(geometry)
        settings["windows"] = windows
        save_settings(base_directory, settings)

    trigger = debounce(window, f"geometry-{key}", delay_ms, store)
    window.bind("<Configure>", trigger, add="+")


def _lines_fingerprint(links: list[LineLink]) -> tuple:
    """线路数据的轻量指纹：用来判断"打开编辑器之后，别处有没有改过它"。"""
    return tuple(
        sorted(
            (line_key(item.line), item.model, len(item.numbers), item.first, item.last, item.status, item.source)
            for item in links
        )
    )


def compact_menu(parent: tk.Misc, text: str, items: list[tuple[str, Callable[[], None]]]) -> ttk.Menubutton:
    """把一串"数据搬运"类操作收进一个下拉按钮，别让底下一排全是按钮。

    为什么：窗口底部堆 7 个同等重要的按钮，用户根本分不清主次。主操作留在明面上，
    「复制 / 导出」这类收进菜单——要用的时候点一下就有，不用的时候不占眼睛。
    """
    button = ttk.Menubutton(parent, text=text, style="Ghost.TButton")
    menu = tk.Menu(button, tearoff=False)
    for label, command in items:
        menu.add_command(label=label, command=command)
    button.configure(menu=menu)
    return button


def debounce(widget: tk.Misc, key: str, delay_ms: int, callback: Callable[[], None]) -> Callable[..., None]:
    """返回一个去抖触发器：连续调用只在最后一次之后 delay_ms 才真正执行。"""
    state: dict[str, str | None] = {"job": None}

    def run() -> None:
        state["job"] = None
        callback()

    def trigger(*_args: object) -> None:
        pending = state["job"]
        if pending is not None:
            try:
                widget.after_cancel(pending)
            except tk.TclError:
                pass
        try:
            state["job"] = widget.after(delay_ms, run)
        except tk.TclError:
            state["job"] = None

    trigger.key = key  # type: ignore[attr-defined]
    return trigger


def center_window(
    window: tk.Misc,
    parent: tk.Misc | None,
    width: int,
    height: int,
    *,
    settings_key: str | None = None,
    base_directory: Path | None = None,
) -> None:
    """Place a dialog: 优先用记住的位置，否则贴着父窗口居中并夹到屏幕内。"""
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    if settings_key and base_directory is not None:
        stored = stored_geometry(base_directory, settings_key)
        if stored is not None:
            x, y, stored_width, stored_height = stored
            stored_width = max(240, min(stored_width, screen_width))
            stored_height = max(200, min(stored_height, screen_height))
            x = max(0, min(x, max(0, screen_width - stored_width)))
            y = max(0, min(y, max(0, screen_height - stored_height)))
            window.geometry(f"{stored_width}x{stored_height}+{x}+{y}")
            return
    if parent is not None and parent.winfo_ismapped():
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 3
    else:
        x = (screen_width - width) // 2
        y = (screen_height - height) // 3
    x = max(0, min(x, max(0, screen_width - width)))
    y = max(0, min(y, max(0, screen_height - height)))
    window.geometry(f"{width}x{height}+{x}+{y}")


class EmptyState:
    """表格没有内容时，在表格上方盖一句友好提示（ttk 没有空状态）。"""

    def __init__(self, tree: ttk.Treeview, text: str, theme: Theme) -> None:
        self.tree = tree
        self.text = text
        self.label = tk.Label(
            tree,
            text=text,
            background=theme.surface,
            foreground=theme.muted,
            font=theme.font(10),
            justify="center",
            wraplength=theme.px(360),
        )

    def update(self, text: str | None = None, has_rows: bool | None = None) -> None:
        if text is not None:
            self.text = text
            self.label.configure(text=text)
        rows = tree_rows(self.tree) if has_rows is None else ([] if not has_rows else [("x",)])
        if rows:
            self.label.place_forget()
            return
        try:
            self.label.place(relx=0.5, rely=0.5, anchor="center")
        except tk.TclError:
            pass


def heading_with_sort(tree: ttk.Treeview, columns: tuple, sort_column: str, sort_reverse: bool) -> None:
    """给表头标题加上 ▲ / ▼，让人知道现在按哪一列排序、是升还是降。"""
    for identifier, title, *_rest in columns:
        marker = ""
        if identifier == sort_column:
            marker = " ▼" if sort_reverse else " ▲"
        tree.heading(identifier, text=f"{title}{marker}")


def enable_row_hover(tree: ttk.Treeview, colour: str, tag: str = "hover") -> None:
    """鼠标划过时整行高亮（ttk 的 Treeview 没有 hover 态，只能自己维护 tag）。"""
    tree.tag_configure(tag, background=colour)
    state = {"row": ""}

    def on_motion(event: tk.Event) -> None:
        row = tree.identify_row(event.y)
        if row == state["row"]:
            return
        if state["row"]:
            try:
                tree.item(state["row"], tags=tuple(item for item in tree.item(state["row"], "tags") if item != tag))
            except tk.TclError:
                pass
        state["row"] = row
        if row:
            try:
                tree.item(row, tags=tuple(tree.item(row, "tags")) + (tag,))
            except tk.TclError:
                pass

    def on_leave(_event: tk.Event) -> None:
        if state["row"]:
            try:
                tree.item(state["row"], tags=tuple(item for item in tree.item(state["row"], "tags") if item != tag))
            except tk.TclError:
                pass
            state["row"] = ""

    tree.bind("<Motion>", on_motion, add="+")
    tree.bind("<Leave>", on_leave, add="+")


def attach_placeholder(entry: ttk.Entry, text: str, theme: Theme) -> tk.Label:
    """给输入框加一句灰色占位提示（用覆盖标签实现，不碰 textvariable）。"""

    label = tk.Label(entry, text=text, background=theme.surface, foreground=theme.muted, font=theme.font(9), bd=0)

    def refresh(*_args: object) -> None:
        try:
            empty = not entry.get()
        except tk.TclError:
            return
        if empty:
            label.place(x=theme.px(9), rely=0.5, anchor="w")
        else:
            label.place_forget()

    for sequence in ("<KeyRelease>", "<FocusIn>", "<FocusOut>", "<<Paste>>", "<<Cut>>"):
        entry.bind(sequence, refresh, add="+")
    entry.bind("<Configure>", lambda _event: refresh(), add="+")
    refresh()
    return label


def tree_rows(tree: ttk.Treeview) -> list[tuple[str, ...]]:
    return [tuple(str(value) for value in tree.item(item, "values")) for item in tree.get_children()]


def tree_headings(tree: ttk.Treeview) -> list[str]:
    return [tree.heading(column, "text") for column in tree["columns"]]


def copy_tree_rows(widget: tk.Misc, tree: ttk.Treeview) -> int:
    """把表格内容复制成 TSV（可直接粘进 Excel），返回行数。"""
    rows = tree_rows(tree)
    if not rows:
        return 0
    text = "\n".join(["\t".join(tree_headings(tree))] + ["\t".join(row) for row in rows])
    widget.clipboard_clear()
    widget.clipboard_append(text)
    widget.update()
    return len(rows)


@dataclass(frozen=True)
class TableColumn:
    """一列的配方：标题、宽度、对齐，以及排序方式。"""

    key: str
    title: str
    width: int
    anchor: str = "w"  # "w" / "center" / "e"
    sortable: bool = True
    numeric: bool = False  # 数字列：右对齐 + 排序按数值


def _default_sort_value(text: str, numeric: bool):
    if numeric:
        match = re.match(r"-?\d+", text.strip())
        return (0, int(match.group())) if match else (1, 0)
    return (0, text.strip().casefold())


def keep_row_ids(rows: list[tuple], ids: list[str]) -> Callable[[int, tuple], str]:
    """给 DataTable.set_rows 用的 iid 回调：表格会按表头重排行，行的位置会变，
    但「原来第几条」不能变——这里把原来的标识按行内容重新发给排序后的行。

    同一内容出现多次时按出现顺序依次分配，跟原列表的顺序一致。
    """
    pending: dict[tuple, list[str]] = {}
    for row, identifier in zip(rows, ids):
        pending.setdefault(tuple(str(value) for value in row), []).append(identifier)

    def make(index: int, row: tuple) -> str:
        queue = pending.get(tuple(str(value) for value in row))
        return queue.pop(0) if queue else f"row-{index}"

    return make


class DataTable:
    """一个卡片里的 Treeview：斑马纹、hover、排序指示、复制/导出、空状态一次配好。

    以前 18 个窗口各写各的，导致"排序箭头/导出/空状态"这些行为只在个别窗口有。
    """

    def __init__(
        self,
        parent: tk.Misc,
        theme: Theme,
        columns: tuple[TableColumn, ...],
        *,
        selectmode: str = "browse",
        sort_key: Callable[[str, tuple[str, ...]], object] | None = None,
        empty_text: str = "",
        on_open: Callable[[], None] | None = None,
        stripe: bool = True,
        toolbar: Callable[[tk.Frame], None] | None = None,
        sort_on_load: bool = True,
        on_delete: Callable[[], None] | None = None,
        on_search: Callable[[], None] | None = None,
    ) -> None:
        self.theme = theme
        self.columns = columns
        self.sort_key = sort_key
        self.on_open = on_open
        # sort_on_load=False：先按塞进来的顺序显示（批量查询要按用户输入的顺序复制出去）
        self._sort_on_load = sort_on_load
        self.sort_column = (columns[0].key if columns else "") if sort_on_load else ""
        self.sort_reverse = False
        self._rows: list[tuple[str | None, tuple[str, ...]]] = []
        self._iid: Callable[[int, tuple], str] | None = None
        self._stripe = stripe

        self.card = theme.card(parent)
        self.toolbar: tk.Frame | None = None
        if toolbar is not None:
            self.toolbar = ttk.Frame(self.card, style="Card.TFrame", padding=(theme.px(10), theme.px(8)))
            self.toolbar.pack(side="top", fill="x")
            toolbar(self.toolbar)
        self.tree = ttk.Treeview(
            self.card,
            columns=[column.key for column in columns],
            show="headings",
            selectmode=selectmode,
        )
        self._configure_columns()
        if stripe:
            self.tree.tag_configure("odd", background=theme.stripe)
        enable_row_hover(self.tree, theme.hover)
        scrollbar = ttk.Scrollbar(self.card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y", padx=(0, 1), pady=1)
        self.tree.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        if on_open is not None:
            self.tree.bind("<Double-Button-1>", lambda _event: on_open())
        # 右键菜单：只想复制某一列（例如车型）时不用手抄
        self.tree.bind("<Button-3>", self._popup_menu)
        # 键盘：Delete 删选中行（只在焦点在这张表时生效）、Ctrl+F 跳到搜索框
        if on_delete is not None:
            self.tree.bind("<Delete>", lambda _event: on_delete())
        if on_search is not None:
            self.tree.winfo_toplevel().bind("<Control-f>", lambda _event: on_search(), add="+")
        self.empty = EmptyState(self.tree, empty_text, theme)

    def _configure_columns(self) -> None:
        for index, column in enumerate(self.columns):
            self.tree.heading(
                column.key,
                text=column.title,
                anchor="center" if column.anchor != "w" else "w",
                command=(lambda key=column.key: self.sort_by(key)) if column.sortable else "",
            )
            self.tree.column(column.key, width=self.theme.px(column.width), anchor=column.anchor, stretch=index == 0)
        self.update_headings()

    def set_columns(self, columns: tuple[TableColumn, ...]) -> None:
        """换一组列（对账窗口的列随类别变）。

        注意别把 `sort_on_load=False` 的表又强行按第一列排序——那会把调用方排好的顺序打乱
        （对账的「型号不一致」是按"值得查"排的，被重置过一次）。
        """
        self.columns = columns
        self.sort_column = (columns[0].key if columns else "") if getattr(self, "_sort_on_load", True) else ""
        self.sort_reverse = False
        self.tree.delete(*self.tree.get_children())
        self.tree.configure(columns=[column.key for column in columns])
        self._configure_columns()
        self._rows = []

    # -- 数据 ------------------------------------------------------------- #
    def set_rows(
        self,
        rows: list[tuple],
        *,
        keep_selection: bool = True,
        iid: Callable[[int, tuple], str] | None = None,
    ) -> None:
        """塞入新数据：自动斑马纹、自动重排（如果用户点过表头）、尽量保持选中。

        iid 在**排序之前**按传入顺序算好并跟着行走：所以 `iid=lambda position, row: str(position)`
        这类"用原始下标当行 id"的窗口，即使点了表头排序，行 id 仍指向原来那条数据
        （号段管理 / 拍车日志 / 地铁号段管理的编辑与删除都靠它定位，错位就会改错行）。
        """
        text_rows = [tuple(str(value) for value in row) for row in rows]
        pairs = [
            (iid(index, text_rows[index]) if iid else None, text_rows[index])
            for index in range(len(text_rows))
        ]
        self._iid = iid
        self._original = pairs          # 原始顺序留着：点表头三下可以回到"塞进来的顺序"
        self._rows = self._sorted(pairs)
        previous = set(self.tree.selection())
        previous_first = str(self.tree.item(self.tree.selection()[0], "values")[0]) if previous else ""
        self.tree.delete(*self.tree.get_children())
        for index, (row_id, row) in enumerate(self._rows):
            tags = ("odd",) if (self._stripe and index % 2) else ()
            self.tree.insert("", "end", iid=row_id, values=row, tags=tags)
        if keep_selection and self._rows:
            children = self.tree.get_children()
            wanted = [child for child in children if child in previous]
            if not wanted and previous_first:
                wanted = [child for child in children if str(self.tree.item(child, "values")[0]) == previous_first]
            wanted = wanted or list(children[:1])
            self.tree.selection_set(wanted[0])
        self.update_headings()
        self.empty.update(has_rows=bool(self._rows))

    def _sorted(self, pairs: list[tuple[str | None, tuple[str, ...]]]) -> list[tuple[str | None, tuple[str, ...]]]:
        column = next((item for item in self.columns if item.key == self.sort_column), None)
        if column is None or not column.sortable:
            return pairs
        index = self.tree["columns"].index(self.sort_column)
        if self.sort_key is not None:
            key = lambda item: self.sort_key(self.sort_column, item[1])  # noqa: E731
        else:
            key = lambda item: _default_sort_value(item[1][index], column.numeric)  # noqa: E731
        try:
            return sorted(pairs, key=key, reverse=self.sort_reverse)
        except (TypeError, ValueError):
            return pairs

    def sort_by(self, key: str) -> None:
        """点表头：升序 → 降序 → **回到原始顺序**（行 id 已绑在数据上，直接重排即可）。

        第三下回到"塞进来的顺序"很有用：批量查询的结果要按输入顺序复制出去。
        """
        if self.sort_column != key:
            self.sort_column, self.sort_reverse = key, False
        elif not self.sort_reverse:
            self.sort_reverse = True
        else:
            self.sort_column, self.sort_reverse = "", False
        self._rows = self._sorted(getattr(self, "_original", list(self._rows)))
        previous = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        for index, (row_id, row) in enumerate(self._rows):
            tags = ("odd",) if (self._stripe and index % 2) else ()
            self.tree.insert("", "end", iid=row_id, values=row, tags=tags)
        keep = [child for child in self.tree.get_children() if child in previous]
        if keep:
            self.tree.selection_set(keep[0])
        self.update_headings()

    def update_headings(self) -> None:
        """表头上标出当前排序列与方向（▲ 升 / ▼ 降 / ⇅ 原顺序）。"""
        for column in self.columns:
            marker = ""
            if column.key == self.sort_column:
                marker = " ▼" if self.sort_reverse else " ▲"
            elif not self.sort_column and column.sortable:
                marker = " ⇅"
            self.tree.heading(column.key, text=f"{column.title}{marker}")

    # -- 取用 ------------------------------------------------------------- #
    def rows(self) -> list[tuple[str, ...]]:
        return tree_rows(self.tree)

    def selection_first(self) -> str:
        selection = self.tree.selection()
        return str(self.tree.item(selection[0], "values")[0]) if selection else ""

    def selected_values(self) -> list[tuple[str, ...]]:
        return [tuple(str(value) for value in self.tree.item(item, "values")) for item in self.tree.selection()]

    def copy(self, widget: tk.Misc | None = None) -> int:
        return copy_tree_rows(widget or self.tree, self.tree)

    def copy_column(self, key: str, *, only_selected: bool = False) -> int:
        """只复制某一列（例如"只想复制车型，不想复制别的"）。

        默认复制**整列**——刷新表格后程序会自动选中第一行，若按"选中优先"就只会复制一行，
        对"把车型列拿出来"这种事反而添乱。想只复制选中的行，走右键菜单里的那一项。
        """
        keys = [column.key for column in self.columns]
        if key not in keys:
            return 0
        index = keys.index(key)
        iids = list(self.tree.selection()) if only_selected else list(self.tree.get_children())
        values = [str(self.tree.item(item, "values")[index]) for item in iids]
        values = [value for value in values if value != ""]
        if not values:
            return 0
        self.tree.clipboard_clear()
        self.tree.clipboard_append("\n".join(values))
        return len(values)

    def _popup_menu(self, event: tk.Event) -> None:
        """右键：复制选中行 / 只复制光标所在的那一列。"""
        iid = self.tree.identify_row(event.y)
        if iid and iid not in self.tree.selection():
            self.tree.selection_set(iid)
        column_id = self.tree.identify_column(event.x)
        column = None
        if column_id.startswith("#"):
            try:
                index = int(column_id[1:]) - 1
                column = self.columns[index] if 0 <= index < len(self.columns) else None
            except ValueError:
                column = None
        menu = tk.Menu(self.tree, tearoff=0)
        selected = len(self.tree.selection())
        menu.add_command(label=f"复制选中行（{selected} 行）" if selected else "复制全部行", command=self.copy)
        if column is not None:
            if selected:
                menu.add_command(
                    label=f"只复制「{column.title}」列（选中 {selected} 行）",
                    command=lambda key=column.key: self.copy_column(key, only_selected=True),
                )
            menu.add_command(label=f"只复制「{column.title}」列（全部行）",
                             command=lambda key=column.key: self.copy_column(key))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def export(self, parent: tk.Misc, default_name: str) -> str | None:
        return export_tree_rows(parent, self.tree, default_name)


def export_tree_rows(parent: tk.Misc, tree: ttk.Treeview, default_name: str) -> str | None:
    """把表格导出成 utf-8-sig 的 CSV，返回路径（取消或失败返回 None）。"""
    rows = tree_rows(tree)
    if not rows:
        return None
    path = filedialog.asksaveasfilename(
        parent=parent,
        title="导出为 CSV",
        defaultextension=".csv",
        initialfile=default_name,
        filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
    )
    if not path:
        return None
    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(tree_headings(tree))
            writer.writerows(rows)
    except OSError as error:
        messagebox.showerror("导出失败", f"无法写入文件：{error}", parent=parent)
        return None
    return path


class WindowMixin:
    """Shared behaviour for every Toplevel dialog."""
    app: "BusLookupApp"
    theme: Theme

    def __init_subclass__(cls, **kwargs: object) -> None:
        """给每个窗口的 __init__ 套一层守卫：构造中途出错就把半成品窗口销毁掉。

        为什么需要：`super().__init__(parent)` 一执行，Tk 里就真的多了一个顶层窗口；
        之后 __init__ 里任何一行抛异常（参数不对、数据坏了、图片读不出来），
        调用方拿到异常，但那个**空窗口会一直留在屏幕上、并且永远不会被销毁**——
        反复触发就是实打实的窗口泄漏（句柄、Tk 命令、Python 对象一起涨）。
        """
        super().__init_subclass__(**kwargs)
        original = cls.__dict__.get("__init__")
        if original is None or getattr(original, "_bpt_guarded", False):
            return

        @functools.wraps(original)
        def guarded(self, *args: object, **kwargs2: object) -> None:
            try:
                original(self, *args, **kwargs2)
            except Exception:
                try:
                    if self.winfo_exists():
                        self.destroy()
                except Exception:  # noqa: BLE001 - 清理失败也别盖掉原始异常
                    pass
                raise

        guarded._bpt_guarded = True  # type: ignore[attr-defined]
        cls.__init__ = guarded  # type: ignore[assignment]

    def setup_dialog(self, parent: tk.Misc, title: str, width: int, height: int, *, resizable: bool = True) -> None:
        window = self  # type: ignore[assignment]
        window.title(title)
        window.transient(parent.winfo_toplevel())
        window.configure(background=self.theme.background)
        if resizable:
            window.minsize(self.theme.px(width * 0.75), self.theme.px(height * 0.6))
        else:
            window.resizable(False, False)
        base = application_dir()
        key = self.__class__.__name__
        # 记住计划尺寸：内容比它高时可以按内容调高（见 fit_dialog_height）
        self.planned_size = (self.theme.px(width), self.theme.px(height))
        center_window(
            window,
            parent.winfo_toplevel(),
            self.theme.px(width),
            self.theme.px(height),
            settings_key=key,
            base_directory=base,
        )
        remember_geometry(window, base, key)
        window.bind("<Escape>", lambda _event: self.close_dialog())
        window.protocol("WM_DELETE_WINDOW", self.close_dialog)
        # 内容建好之后自动兜一次底：不够高就按内容调高（并尊重屏幕上限）。
        # 放在这里而不是靠每个窗口自己记得调——否则新增窗口一漏，底部按钮又会被挤出窗口。
        window.after_idle(self.fit_dialog_height)

    def fit_dialog_height(self, minimum: int = 0) -> None:
        """按内容把窗口调高（并放宽到可调整大小），但不超过屏幕可用高度。

        窗口高度是写死的，内容一旦变多（例如 V6 给标签加了「车体形式」），
        底部那一行按钮就会被挤出窗口、用户根本点不到。这里在内容建好之后
        按 `winfo_reqheight()` 兜一次底；内容比屏幕还高时按屏幕封顶，
        超出的部分交给窗口里可伸缩的表格去吸收。
        """
        window = self  # type: ignore[assignment]
        window.update_idletasks()
        width, planned = getattr(self, "planned_size", (0, 0))
        wanted = max(window.winfo_reqheight(), planned, self.theme.px(minimum))
        screen_cap = window.winfo_screenheight() - self.theme.px(60)
        wanted = min(wanted, screen_cap)
        if wanted > planned:
            window.geometry(f"{width or window.winfo_reqwidth()}x{wanted}")
            window.resizable(True, True)

    def on_data_changed(self, kinds: set[str]) -> None:
        """数据变化时的默认反应：不认识的窗口什么也不做。"""

    def dialog_header(self, parent: tk.Misc, title: str, subtitle: str = "", *, inside_card: bool = False, wrap: int = 520) -> tk.Frame:
        """对话框顶部的统一标题：大标题 + 一行浅色说明。"""
        holder = ttk.Frame(parent, style="Card.TFrame" if inside_card else "TFrame")
        holder.pack(fill="x", anchor="w")
        ttk.Label(
            holder,
            text=title,
            style="CardStrong.TLabel" if inside_card else "DialogTitle.TLabel",
        ).pack(anchor="w")
        if subtitle:
            ttk.Label(
                holder,
                text=subtitle,
                style="CardMuted.TLabel" if inside_card else "Muted.TLabel",
                wraplength=self.theme.px(wrap),
                justify="left",
            ).pack(anchor="w", pady=(self.theme.px(3), 0))
        return holder

    def collapse_when_empty(
        self, variable: tk.StringVar, label: tk.Widget, *, before: tk.Misc | None = None, pady: int = 0
    ) -> None:
        """让"结果说明"这类标签在没有内容时收起来，别在窗口里留一条空白带。

        重排时用 pack(before=…) 保持版面顺序（不然会跑到最下面）。
        """
        def sync(*_args: object) -> None:
            if variable.get().strip():
                if not label.winfo_ismapped():
                    if before is not None:
                        label.pack(fill="x", pady=pady, before=before)
                    else:
                        label.pack(fill="x", pady=pady)
            elif label.winfo_ismapped():
                label.pack_forget()

        variable.trace_add("write", sync)
        sync()

    def release_variables(self) -> None:
        """关窗时把窗口上所有 tk.Variable 的 trace 摘掉，并丢掉引用。

        为什么必须做（实测出来的泄漏）：`var.trace_add("write", 回调)` 会在 Tcl 里注册一个命令，
        而 tkinter 把这些 Python 回调存在解释器的 `_tclCommands` 列表里——也就是说
        **Tcl 命令持有 Python 回调，回调闭包持有窗口，窗口又持有变量**，构成一个
        Python GC 看不见的环：变量永远不会被回收，`Variable.__del__` 也就永远删不掉那条 Tcl 命令。
        反复开关带搜索框的窗口，每轮稳定留下 1~6 个 Tcl 命令（CatalogWindow 6 个、SpotLogWindow 3 个…），
        是真正的线性泄漏。这里在 destroy 之前主动 trace_remove + 丢引用，就能清干净。
        """
        for key, value in list(vars(self).items()):
            if not isinstance(value, tk.Variable):
                continue
            try:
                for mode, name in [(item[0], item[1]) for item in value.trace_info()]:
                    value.trace_remove(mode, name)
            except (tk.TclError, ValueError):
                pass
            try:
                delattr(self, key)
            except AttributeError:
                pass

    def destroy(self) -> None:  # type: ignore[override]
        try:
            self.release_variables()
        except Exception:  # noqa: BLE001 - 清理失败不能挡住关窗
            pass
        super().destroy()  # type: ignore[attr-defined]

    def close_dialog(self) -> None:
        self.destroy()  # type: ignore[attr-defined]


# --------------------------------------------------------------------------- #
# Main window
# --------------------------------------------------------------------------- #


class BusLookupApp(tk.Tk):
    #: 主界面"顺手查联网"的默认值。测试里会把这个类属性改成 False，
    #: 免得整个测试套件每查一次自编号都真去联网（那会让测试变慢、还依赖外网）。
    AUTO_ONLINE_DEFAULT = True

    def __init__(self) -> None:
        enable_dpi_awareness()
        super().__init__()
        install_exception_handling(self)
        self.theme = Theme(self)
        self.title(APP_TITLE)
        self.configure(background=self.theme.background)
        self.ranges: list[BusRange] = []
        self.metro_ranges: list[MetroRange] = []
        self.profiles: dict[str, dict[str, str]] = {}
        self.tag_overrides: dict = {}
        self.lines: list[LineLink] = []
        # 线路资料（联网查的首末站/站点表；没有 线路资料.json 时是空表）
        self.line_info: dict = {}
        # 区属公交：公司 / 单车记录 / 号段 / 空号（没有 区属公交数据.json 时都是空的）
        self.district_bus: tuple[
            list[DistrictBusCompany], list[DistrictBusEntry], list[DistrictBusRange], tuple[str, ...]
        ] = ([], [], [], ())
        self.tag_revision = 0
        self.current_model: str | None = None
        # 主界面"顺手查联网"的状态：去抖定时器 / 结果队列 / 当前在查哪个号
        self._online_after: str | None = None
        self._online_queue: "queue.Queue[tuple]" = queue.Queue()
        self._online_number = ""
        self.number_var = tk.StringVar()
        self.status_var = tk.StringVar(value="就绪")
        self.hint_var = tk.StringVar(value="输入五位自编号，边输入边出结果")
        # 联网补的信息（现在在哪 / 最近跑哪条线）写进**底部状态栏**，不占窗口高度
        self.online_var = tk.StringVar(value="")
        self._base_status = ""
        self.result_var = tk.StringVar(value="正在加载数据…")
        self.subtitle_var = tk.StringVar(value="正在加载数据…")
        self._closing = False
        self.settings: dict = load_settings(application_dir())
        self.preferences: dict = self.settings.get("preferences") if isinstance(self.settings.get("preferences"), dict) else {}
        # 主界面"顺手查联网"是否开启（必须在 preferences 之后读；测试会把类属性默认值改成 False）
        self._auto_online = self.preference("auto_online_lookup", self.AUTO_ONLINE_DEFAULT)
        self.history: list[str] = load_history(application_dir())
        self.clipboard_watcher = ClipboardWatcher(self, self._on_clipboard_number)
        self.hotkey: GlobalHotkey | None = None
        self._suppress_reset = False
        self._status_job: str | None = None
        self._set_window_icon()
        self._build_interface()
        self._load_data()
        self._refresh_history_row()
        self._apply_preferences()
        self.fit_window()
        self.protocol("WM_DELETE_WINDOW", self.close_application)
        self.bind("<Control-c>", self._on_control_c)

    # -- construction ---------------------------------------------------- #
    def _set_window_icon(self) -> None:
        icon = resource_path(ICON_FILENAME)
        if icon.is_file():
            try:
                self.iconbitmap(default=str(icon))
            except tk.TclError:
                pass

    def _build_interface(self) -> None:
        theme = self.theme
        pad = theme.px
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = tk.Frame(self, background=theme.header)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        self._logo_image = self._load_logo(header)
        heading = tk.Frame(header, background=theme.header)
        heading.grid(row=0, column=1, sticky="w", padx=(pad(16), pad(18)), pady=pad(11))
        ttk.Label(heading, text=APP_TITLE, style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(heading, textvariable=self.subtitle_var, style="HeaderInfo.TLabel").pack(anchor="w", pady=(pad(3), 0))
        tk.Frame(self, background=theme.header_accent, height=pad(3)).grid(row=0, column=0, sticky="sew")

        body = ttk.Frame(self, padding=(pad(16), pad(12), pad(16), pad(6)))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)

        query_card = theme.card(body, accent=True)
        query_card.grid(row=0, column=0, sticky="ew")
        query_inner = ttk.Frame(query_card, style="Card.TFrame", padding=(pad(16), pad(12)))
        query_inner.pack(fill="both", expand=True)
        query_inner.columnconfigure(0, weight=1)

        theme.section(query_inner, "五位自编号", background=theme.surface, pady=(0, 4), grid=(0, 0, 2))
        self.entry = ttk.Entry(query_inner, textvariable=self.number_var, style="Query.TEntry")
        self.entry.grid(row=1, column=0, sticky="ew", pady=(pad(6), 0))
        self.entry.bind("<Return>", self._on_enter)
        self.button = ttk.Button(query_inner, text="查询", style="Primary.TButton", command=lambda: self.query(commit=True))
        self.button.grid(row=1, column=1, sticky="e", padx=(pad(10), 0), pady=(pad(6), 0))

        ttk.Separator(query_inner).grid(row=2, column=0, columnspan=2, sticky="ew", pady=pad(11))

        # 结果卡片：左边一条主色竖线 + 车型 + 标签 chips + 号段说明
        self.result_card = tk.Frame(query_inner, background=theme.surface_alt, highlightthickness=1, highlightbackground=theme.border, bd=0)
        self.result_card.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, pad(2)))
        self.result_bar = tk.Frame(self.result_card, background=theme.primary, width=pad(4))
        self.result_bar.pack(side="left", fill="y")
        result_inner = tk.Frame(self.result_card, background=theme.surface_alt)
        result_inner.pack(side="left", fill="both", expand=True, padx=pad(14), pady=pad(12))
        self.result_label = tk.Label(
            result_inner,
            textvariable=self.result_var,
            background=theme.surface_alt,
            foreground=theme.primary,
            font=theme.font(24, "bold"),
            justify="center",
            wraplength=pad(340),
        )
        self.result_label.pack(anchor="center")
        # 联网补的一行（现在在哪 / 最近跑哪条线）就放在车型**正下方**，一眼能看到
        self.online_label = tk.Label(
            result_inner,
            textvariable=self.online_var,
            background=theme.surface_alt,
            foreground=theme.primary,
            font=theme.font(9),
            justify="center",
            wraplength=pad(360),
        )
        self.online_label.pack(anchor="center", pady=(pad(4), 0))
        self.chips_row = tk.Frame(result_inner, background=theme.surface_alt)
        self.chips_row.pack(anchor="center", pady=(pad(6), 0))
        self.hint_label = tk.Label(
            result_inner,
            textvariable=self.hint_var,
            background=theme.surface_alt,
            foreground=theme.muted,
            font=theme.font(9),
            justify="center",
            wraplength=pad(340),
        )
        self.hint_label.pack(anchor="center", pady=(pad(6), 0))

        actions = ttk.Frame(query_inner, style="Card.TFrame")
        actions.grid(row=5, column=0, columnspan=2, pady=(pad(12), 0))
        self.copy_button = ttk.Button(actions, text="复制车型", style="Secondary.TButton", command=self.copy_result)
        self.copy_button.pack(side="left", padx=(0, pad(8)))
        self.detail_button = ttk.Button(actions, text="查看车型详情", style="Secondary.TButton", command=self.open_current_model_detail)
        self.detail_button.pack(side="left", padx=(0, pad(8)))
        self.spot_button = ttk.Button(actions, text="记一笔", style="Secondary.TButton", command=self.record_spot)
        self.spot_button.pack(side="left")
        # 连续录入：查完自动清空，接着录下一辆（注意不要放进 history_row，那里每次刷新会清空子控件）
        self.continuous_var = tk.BooleanVar(value=self.preference("continuous_entry"))
        self.continuous_check = ttk.Checkbutton(
            actions,
            text="连续录入",
            variable=self.continuous_var,
            style="Card.TCheckbutton",
            command=self.toggle_continuous,
        )
        self.continuous_check.pack(side="right")
        self.copy_button.state(["disabled"])
        self.detail_button.state(["disabled"])
        self.spot_button.state(["disabled"])

        self.history_row = ttk.Frame(query_inner, style="Card.TFrame")
        self.history_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(pad(10), 0))

        # 数据文件加载失败时才显示的自救按钮
        self.repair_row = ttk.Frame(query_inner, style="Card.TFrame")
        self.repair_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(pad(10), 0))
        self.repair_var = tk.StringVar(value="")
        ttk.Label(self.repair_row, textvariable=self.repair_var, style="CardMuted.TLabel", wraplength=pad(400), justify="left").pack(
            anchor="w", pady=(0, pad(6))
        )
        repair_buttons = ttk.Frame(self.repair_row, style="Card.TFrame")
        repair_buttons.pack(anchor="w")
        ttk.Button(repair_buttons, text="⟳ 恢复最近的备份", style="Primary.TButton", command=self.restore_latest_ranges_backup).pack(
            side="left", padx=(0, pad(6))
        )
        ttk.Button(repair_buttons, text="忽略冲突强行加载", style="Secondary.TButton", command=self.force_load_ranges).pack(
            side="left", padx=(0, pad(6))
        )
        ttk.Button(repair_buttons, text="打开程序文件夹", style="Secondary.TButton", command=self.open_data_folder).pack(side="left")
        self.repair_row.grid_remove()

        # 一张"工具"卡：原来的"快查工具"和"数据管理"合并，少一个卡头、少两行按钮
        manage_card = theme.card(body, accent=True)
        manage_card.grid(row=1, column=0, sticky="ew", pady=(pad(10), 0))
        manage_inner = ttk.Frame(manage_card, style="Card.TFrame", padding=(pad(14), pad(9)))
        manage_inner.pack(fill="both", expand=True)
        manage_inner.columnconfigure(0, weight=1)
        manage_inner.columnconfigure(1, weight=1)
        theme.section(manage_inner, "工具", background=theme.surface, pady=(0, 6), grid=(0, 0, 2))
        # 常用的六个两列排好；只有"联网"这两项收进菜单（用得少、还比较新）
        self.catalog_button = ttk.Button(manage_inner, text="车型目录", style="Secondary.TButton", command=self.open_catalog)
        self.catalog_button.grid(row=1, column=0, sticky="ew", padx=(0, pad(6)))
        self.range_button = ttk.Button(manage_inner, text="管理号段", style="Secondary.TButton", command=self.open_range_manager)
        self.range_button.grid(row=1, column=1, sticky="ew", padx=(pad(6), 0))
        self.health_button = ttk.Button(manage_inner, text="数据体检", style="Secondary.TButton", command=self.open_health)
        self.health_button.grid(row=2, column=0, sticky="ew", padx=(0, pad(6)), pady=(pad(8), 0))
        self.metro_button = ttk.Button(manage_inner, text="地铁编号查询", style="Secondary.TButton", command=self.open_metro)
        self.metro_button.grid(row=2, column=1, sticky="ew", padx=(pad(6), 0), pady=(pad(8), 0))
        ttk.Button(manage_inner, text="我的记录（导入 / 对账）", style="Secondary.TButton", command=self.open_reconcile).grid(
            row=3, column=0, sticky="ew", padx=(0, pad(6)), pady=(pad(8), 0)
        )
        ttk.Button(manage_inner, text="线路查询（我记过的）", style="Secondary.TButton", command=self.open_line_query).grid(
            row=3, column=1, sticky="ew", padx=(pad(6), 0), pady=(pad(8), 0)
        )
        ttk.Button(manage_inner, text="区属公交查询", style="Secondary.TButton", command=self.open_district_bus).grid(
            row=4, column=0, sticky="ew", padx=(0, pad(6)), pady=(pad(8), 0)
        )
        self.online_button = ttk.Button(
            manage_inner, text="联网查询", style="Secondary.TButton", command=self.open_online_query
        )
        self.online_button.grid(row=4, column=1, sticky="ew", padx=(pad(6), 0), pady=(pad(8), 0))
        compact_menu(
            manage_inner,
            "快查工具 ▾",
            [
                ("批量查询编号", lambda: self.open_quick_lookup(0)),
                ("号段区间查询", lambda: self.open_quick_lookup(1)),
                ("拍车日志", self.open_spot_log),
            ],
        ).grid(row=5, column=0, sticky="ew", padx=(0, pad(6)), pady=(pad(6), 0))
        compact_menu(
            manage_inner,
            "联网工具 ▾",
            [
                ("线路资料（首末站 / 站点 / 实时）", self.open_line_info),
                ("联网查询（这辆车跑过哪些线）", self.open_online_query),
            ],
        ).grid(row=5, column=1, sticky="ew", padx=(pad(6), 0), pady=(pad(6), 0))

        status = ttk.Frame(self, padding=(pad(18), 0, pad(18), pad(10)))
        status.grid(row=2, column=0, sticky="ew")
        ttk.Separator(status).pack(fill="x", pady=(0, pad(6)))
        status_row = ttk.Frame(status)
        status_row.pack(fill="x")
        self.status_dot = tk.Label(status_row, text="●", background=self.theme.background, foreground=self.theme.success, font=self.theme.font(9))
        self.status_dot.pack(side="left", padx=(0, pad(6)))
        ttk.Label(status_row, textvariable=self.status_var, style="Muted.TLabel", wraplength=pad(310)).pack(side="left")
        ttk.Button(status_row, text="⚙ 设置", style="Secondary.TButton", command=self.open_settings).pack(side="right")
        ttk.Button(status_row, text="🔍 自检", style="Ghost.TButton", command=self.open_selfcheck).pack(
            side="right", padx=(0, pad(6))
        )

        # 少用的联网两项已收进卡片里的菜单；不再整块折叠（折叠后窗口收不回来、也太空，用户反馈过）
        self.update_idletasks()
        width = max(pad(468), self.winfo_reqwidth())
        # 高度：内容要多少给多少，只受"屏幕 − 120"限制（绝不裁内容）
        height = min(self.winfo_reqheight(), max(480, int(self.winfo_screenheight()) - 120))
        center_window(self, None, width, height)
        self.minsize(pad(440), min(height, pad(520)))
        self._sized = True
        # 输入即查：注册在最后，确保所有控件都已就绪
        self.number_var.trace_add("write", self._on_number_changed)
        self.entry.focus_set()

    def fit_window(self) -> None:
        """内容变高时把窗口一起长高（上限：屏幕高 − 120，且不超过 1240，避免顶满屏幕）。"""
        try:
            self.update_idletasks()
            needed = min(self.winfo_reqheight(), max(480, int(self.winfo_screenheight()) - 120))
            if needed > self.winfo_height() + 2:
                self.geometry(f"{self.winfo_width()}x{needed}")
        except tk.TclError:
            pass

    def _load_logo(self, parent: tk.Misc) -> ImageTk.PhotoImage | None:
        icon = resource_path(ICON_FILENAME)
        if not icon.is_file():
            return None
        try:
            with Image.open(icon) as image:
                logo = image.convert("RGBA")
            size = self.theme.px(40)
            logo.thumbnail((size, size), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(logo)
        except (OSError, UnidentifiedImageError):
            return None
        label = tk.Label(parent, image=photo, background=self.theme.header, bd=0)
        label.grid(row=0, column=0, padx=(self.theme.px(16), 0), pady=self.theme.px(10))
        return photo

    # -- data ------------------------------------------------------------ #
    def _load_data(self) -> None:
        base = application_dir()
        try:
            self.ranges = load_ranges(base / DATA_FILENAME)
        except DataError as error:
            self.current_model = None
            self.result_var.set("数据文件不可用")
            self.result_label.configure(foreground=self.theme.accent)
            self.hint_var.set(str(error))
            self.subtitle_var.set("未加载号段数据")
            self.status_var.set(f"程序文件夹：{base}")
            self.button.state(["disabled"])
            self.repair_var.set("别急，还有救：\n· 恢复最近的备份（推荐，会把坏文件也留一份）\n· 忽略冲突强行加载（先凑合用，之后在「管理号段」里保存一次就修好了）\n· 打开程序文件夹，自己用 Excel 改")
            self.repair_row.grid()
            self.fit_window()
            return
        self.repair_row.grid_remove()
        self.result_var.set("—")
        self.subtitle_var.set(f"{len(self.ranges)} 个号段 · {len(model_names(self.ranges))} 种车型")
        notes: list[str] = []
        self.set_tag_overrides(load_tag_overrides(base), notify=False)
        line_notes = self.reload_lines(base)
        if line_notes:
            notes.append(f"线路数据：{line_notes[0]}")
        self.reload_district_bus(base)  # 预留：没有文件就是两个空表
        self.line_info = load_line_info(base)  # 线路资料（联网查的，缺文件就是空表）
        try:
            self.profiles = load_profiles(base)
        except ProfileError as error:
            self.profiles = {}
            notes.append(str(error))
        try:
            self.metro_ranges = load_metro_ranges(base)
            if not metro_data_path(base).exists():
                notes.append(f"地铁数据已载入，但无法写入 {METRO_DATA_FILENAME}（程序文件夹不可写）")
        except MetroError as error:
            self.metro_ranges = []
            self.metro_button.state(["disabled"])
            notes.append(str(error))
        self.status_var.set("；".join(notes) if notes else "就绪")

    def reload_lines(self, base: Path | None = None) -> list[str]:
        """读线路数据（默认空文件也能跑；坏条目只报告不打不开程序）。"""
        base = base or application_dir()
        self.lines, notes = load_lines_lenient(base)
        return notes

    def reload_district_bus(self, base: Path | None = None) -> None:
        """读区属公交数据（没有文件就是空表）：公司 / 单车记录 / 号段 / 空号。"""
        self.district_bus = load_district_bus(base or application_dir())

    # -- 数据文件坏了怎么救 ----------------------------------------------- #
    def ranges_backups(self) -> list[tuple[str, Path, str, float, int]]:
        """与号段表有关的备份（新的在前）。"""
        return [entry for entry in list_backups(application_dir()) if entry[2] == DATA_FILENAME]

    def restore_latest_ranges_backup(self) -> None:
        backups = self.ranges_backups()
        if not backups:
            messagebox.showinfo(
                "恢复备份",
                f"还没有 {DATA_FILENAME} 的备份。\n可以去「数据体检 → 备份与恢复」看看别的备份，\n或者手动改回 Excel。",
                parent=self,
            )
            return
        label, path, original, mtime, _size = backups[0]
        when = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
        if not messagebox.askyesno("恢复备份", f"用 {when} 的备份覆盖当前的「{label}」吗？\n（当前这个坏文件也会先备份一份）", parent=self):
            return
        try:
            restore_backup(application_dir(), path, original)
        except DataError as error:
            messagebox.showerror("恢复失败", str(error), parent=self)
            return
        self.reload_data()
        self.status_var.set(f"已恢复到 {when} 的号段表")

    def force_load_ranges(self) -> None:
        """跳过坏行、丢掉重叠行，先把能用的数据加载起来。"""
        base = application_dir()
        try:
            ranges, notes = load_ranges_lenient(base / DATA_FILENAME)
        except DataError as error:
            messagebox.showerror("还是读不出来", str(error), parent=self)
            return
        detail = "\n".join(f"· {note}" for note in notes[:8]) + ("\n…" if len(notes) > 8 else "")
        if not messagebox.askyesno(
            "强行加载",
            f"会用 {len(ranges)} 个号段继续工作（跳过了 {len(notes)} 处问题）。\n\n{detail}\n\n"
            f"注意：这只是在内存里凑合先用，文件里的坏行还在。\n"
            f"在「管理号段」里保存一次（或先恢复备份）才会真正修好文件。",
            parent=self,
        ):
            return
        self.ranges = ranges
        self.repair_row.grid_remove()
        self.button.state(["!disabled"])
        self.result_var.set("—")
        self.subtitle_var.set(f"{len(ranges)} 个号段 · {len(model_names(ranges))} 种车型（已忽略 {len(notes)} 处问题）")
        self.status_var.set("已忽略冲突强行加载；记得在「管理号段」里保存一次来修好文件")
        self.notify_data_changed({"all"})
        self.fit_window()

    def open_data_folder(self) -> None:
        try:
            os.startfile(application_dir())  # noqa: S606 - 打开资源管理器
        except OSError as error:
            messagebox.showerror("打开失败", str(error), parent=self)

    # -- actions --------------------------------------------------------- #
    def _on_enter(self, _event: tk.Event) -> None:
        self.query(commit=True)

    def _on_number_changed(self, *_args: object) -> None:
        """边输入边查询：凑满五位立刻出结果，不满足则回到待输入状态。"""
        if getattr(self, "_suppress_reset", False):
            return  # 连续录入清空输入框时不要擦掉刚才的结果
        if validate_number(self.number_var.get()) is not None:
            self.query()
        else:
            self._reset_result()

    def _reset_result(self) -> None:
        self.current_model = None
        self.copy_button.state(["disabled"])
        self.detail_button.state(["disabled"])
        self.spot_button.state(["disabled"])
        self._paint_result(self.theme.muted, self.theme.border)
        self._set_chips([])
        self.result_var.set("—")
        self.hint_var.set("输入五位自编号，边输入边出结果")
        self._base_status = ""
        self.set_status("")

    def _paint_result(self, text_colour: str, bar_colour: str) -> None:
        self.result_label.configure(foreground=text_colour)
        self.result_bar.configure(background=bar_colour)

    def open_quick_lookup(self, tab: int = 0) -> None:
        if not self.ranges:
            messagebox.showerror("快查工具", "请先修复 Excel 数据文件后再使用快查工具。", parent=self)
            return
        QuickLookupWindow(self, tab)

    def open_selfcheck(self) -> None:
        """自检：查数据文件、标签、写入权限、磁盘和（可选）联网。"""
        if getattr(self, "_selfcheck_window", None) is not None:
            try:
                self._selfcheck_window.lift()
                self._selfcheck_window.focus_force()
                return
            except tk.TclError:
                self._selfcheck_window = None
        self._selfcheck_window = SelfCheckWindow(self)
    def open_settings(self) -> None:
        SettingsWindow(self)

    def open_spot_log(self) -> None:
        SpotLogWindow(self)

    def _on_control_c(self, _event: tk.Event) -> str | None:
        focused = self.focus_get()
        if isinstance(focused, (tk.Entry, tk.Text, ttk.Entry)):
            return None
        self.copy_result()
        return "break"

    def query(self, commit: bool = False) -> None:
        """commit=True 表示这次是「正式查询」（回车 / 按钮 / 热键 / 历史），会记历史并朗读。"""
        value = self.number_var.get().strip()
        number = validate_number(value)
        if number is None:
            self.current_model = None
            self.copy_button.state(["disabled"])
            self.detail_button.state(["disabled"])
            self.spot_button.state(["disabled"])
            self._paint_result(self.theme.accent, self.theme.accent)
            self._set_chips([])
            self.result_var.set("格式不符")
            self.hint_var.set("请输入五位数字自编号，例如 20000")
            self._base_status = ""
            self.set_status("")
            self._clear_online()
            return
        model = find_model(number, self.ranges)
        if model is None:
            self.current_model = None
            self.copy_button.state(["disabled"])
            self.detail_button.state(["disabled"])
            self.spot_button.state(["disabled"])
            self._paint_result(self.theme.muted, self.theme.border_strong)
            self._set_chips([])
            self.result_var.set("未收录")
            self.hint_var.set("暂无对应车型，请核对自编号")
            self._base_status = f"自编号 {value} 未收录"
            self.set_status(self._base_status)
            self._clear_online()
            if commit:
                self.remember_query(value)
                self.prepare_next_entry()
            return
        self.current_model = model
        self.copy_button.state(["!disabled"])
        self.detail_button.state(["!disabled"])
        self.spot_button.state(["!disabled"])
        tags = model_tags(model, self.tag_overrides)
        self._paint_result(self.theme.primary, self.theme.primary)
        self._set_chips([tags.manufacturer, tags.fuel, tags.length, tags.body])
        self.result_var.set(model)
        self.hint_var.set(self._range_hint(number, model))
        self._base_status = f"自编号 {value} → {model}"
        self.set_status(self._base_status)
        self.flash_result()
        self.fit_window()
        self.schedule_online_lookup(value)  # 顺手查"现在在哪 / 最新跑哪条线"（后台，不挡输入）
        if commit:
            self.remember_query(value, model)
            self.prepare_next_entry()

    # -- 联网补两行：现在在哪 / 最新在哪条线运营 --------------------------- #
    def schedule_online_lookup(self, number: str) -> None:
        """安排一次后台联网查询（去抖：连着敲数字只会查最后一次）。

        为什么要去抖 + 后台：主界面是"边输入边出结果"，每敲一位都会走 query()，
        联网一次要 0.3~4 秒，绝不能同步做、也不能每敲一位就发一次。
        """
        self._cancel_online_lookup()
        if not self._auto_online:
            self.online_var.set("")
            return
        settings = online_settings()
        if not settings.usable:
            self.online_var.set("")
            return
        self.online_var.set("")   # 新的一次查询：先清掉上一次的旧信息
        self._online_number = number
        self._online_after = self.after(700, lambda: self._start_online_lookup(number))

    def _cancel_online_lookup(self) -> None:
        if self._online_after is not None:
            try:
                self.after_cancel(self._online_after)
            except (tk.TclError, ValueError):
                pass
            self._online_after = None

    def _clear_online(self) -> None:
        self._cancel_online_lookup()
        self._online_number = ""
        self.online_var.set("")

    def _start_online_lookup(self, number: str) -> None:
        self._online_after = None
        if number != self.number_var.get().strip():
            return  # 用户已经改了输入，这次结果没用了
        self.online_var.set("正在联网查这辆车现在在哪…")

        def work() -> None:
            service = shared_service(online_settings(), application_dir())
            track: dict = {}
            latest = ""
            problem = ""
            try:
                track, _stamp = service.track_latest(number)
            except OnlineError as error:
                problem = short_error(error)
            except Exception as error:  # noqa: BLE001 - 兜底
                problem = type(error).__name__
            try:
                history, checked = service.bus_history(number)
                links = lines_from_history(number, history, lambda item: model_lookup(self.ranges, item), checked=checked)
                newest = latest_link(links)
                if newest is not None:
                    latest = f"{newest.line}（{newest.first}–{newest.last}）"
            except OnlineError as error:
                problem = problem or short_error(error)
            except Exception as error:  # noqa: BLE001 - 兜底
                problem = problem or type(error).__name__
            self._online_queue.put((number, track, latest, problem))

        threading.Thread(target=work, name="bpt-main-online", daemon=True).start()
        self.after(150, self._poll_online)

    def _poll_online(self) -> None:
        try:
            number, track, latest, problem = self._online_queue.get_nowait()
        except queue.Empty:
            if self._online_number:      # 还在等这次查询就别停
                self.after(150, self._poll_online)
            return
        if number != self.number_var.get().strip():
            return  # 期间用户改了输入
        self.online_var.set(self._compose_online(track, latest, problem))

    @staticmethod
    def _compose_online(track: dict, latest: str, problem: str) -> str:
        """车型下面那一行：现在在哪 / 最近跑哪条线（尽量压成一行，最多两行）。"""
        parts: list[str] = []
        if track:
            where = track.get("stop_name", "") or "位置未知"
            if track.get("arrived"):
                state = "已到站"
            elif track.get("distance_m") is not None:
                state = f"距站 {track['distance_m']} 米"
            else:
                state = "在途"
            line = track.get("line_name", "") or "线路未知"
            when = str(track.get("time", ""))[-8:]
            parts.append(f"现在 {line} · {where} · {state}" + (f" · {when}" if when else ""))
        elif not problem:
            parts.append("现在没有实时记录（可能没在跑）")
        if latest:
            parts.append(f"最近跑 {latest}")
        if problem:
            parts.append(f"联网没查到（{problem}）")
        return "　｜　".join(parts)

    def flash_result(self) -> None:
        """命中时结果卡片轻轻闪一下，让眼睛知道"查到了"。"""
        try:
            self.result_card.configure(background=self.theme.primary_soft)
            self.after(180, self._restore_result_card)
        except tk.TclError:
            pass

    def _restore_result_card(self) -> None:
        try:
            if not self.winfo_exists():
                return
            self.result_card.configure(background=self.theme.surface_alt)
            for child in (self.result_card.winfo_children()):
                if isinstance(child, tk.Frame) and child is not self.result_bar:
                    child.configure(background=self.theme.surface_alt)
            self.chips_row.configure(background=self.theme.surface_alt)
            self.result_label.configure(background=self.theme.surface_alt)
            self.hint_label.configure(background=self.theme.surface_alt)
        except tk.TclError:
            pass

    def set_status(self, text: str, kind: str = "info", ttl: int = 8000) -> None:
        """状态栏：左侧一个状态点，过一会儿自动回到「就绪」。"""
        colours = {"info": self.theme.primary, "ok": self.theme.success, "warn": self.theme.warning, "error": self.theme.accent}
        try:
            self.status_dot.configure(foreground=colours.get(kind, self.theme.primary))
        except (AttributeError, tk.TclError):
            pass
        self.status_var.set(text)
        if self._status_job is not None:
            try:
                self.after_cancel(self._status_job)
            except tk.TclError:
                pass
            self._status_job = None
        if text and ttl:
            self._status_job = self.after(ttl, self._clear_status)

    def _clear_status(self) -> None:
        self._status_job = None
        try:
            if self.winfo_exists():
                self.status_var.set("就绪")
                self.status_dot.configure(foreground=self.theme.success)
        except tk.TclError:
            pass

    # -- 连续录入 --------------------------------------------------------- #
    def toggle_continuous(self) -> None:
        self.set_preference("continuous_entry", self.continuous_var.get())
        self.status_var.set("连续录入已打开：查完自动清空，接着录下一辆" if self.continuous_var.get() else "连续录入已关闭")

    def prepare_next_entry(self) -> None:
        """连续录入模式下：清空输入框、焦点回来，但把结果留在屏幕上。"""
        if not self.preference("continuous_entry"):
            return
        self.after(700, self._clear_for_next)

    def _clear_for_next(self) -> None:
        if not self.preference("continuous_entry") or not self.winfo_exists():
            return
        self._suppress_reset = True
        try:
            self.number_var.set("")
        finally:
            self._suppress_reset = False
        self.entry.focus_set()

    def _range_hint(self, number: int, model: str) -> str:
        item = find_range(number, self.ranges)
        if item is None:
            return f"自编号 {number:05d}"
        segments = ranges_for_model(model, self.ranges)
        covered = sum(entry.end - entry.start + 1 for entry in segments)
        # 措辞不用「属于」：号段只能说明"号段表里是这段"，不代表车属于谁
        hint = f"号段 {item.start}-{item.end} · 该车型 {len(segments)} 个号段 / 覆盖 {covered} 个编号"
        lines = lines_for_number(self.lines, f"{number:05d}")
        if lines:
            # 措辞刻意用「我记过」而不是「属于」：记录只能说明记录过，不代表现在还在跑
            hint += "\n我记过：" + "、".join(line_display(link, with_status=True) for link in lines[:3])
            if len(lines) > 3:
                hint += f" 等 {len(lines)} 条"
        return hint

    def _model_lines(self, model: str) -> list[LineLink]:
        return lines_for_model(self.lines, model)

    def _set_chips(self, values: list[str], tone: str = "normal") -> None:
        """结果卡片上的标签 chips（厂商 / 燃料 / 车长 / 车体形式）。"""
        for child in self.chips_row.winfo_children():
            child.destroy()
        if not values:
            return
        colours = {
            "normal": (self.theme.primary_soft, self.theme.primary),
            "muted": (self.theme.surface, self.theme.muted),
            "accent": (self.theme.accent_soft, self.theme.accent),
        }
        background, foreground = colours.get(tone, colours["normal"])
        for value in values:
            tk.Label(
                self.chips_row,
                text=value,
                background=background,
                foreground=foreground,
                font=self.theme.font(9, "bold"),
                padx=self.theme.px(9),
                pady=self.theme.px(2),
            ).pack(side="left", padx=self.theme.px(3))

    def remember_query(self, number: str, model: str | None = None) -> None:
        """记进查询历史，并按设置朗读车型。"""
        try:
            self.history = remember_number(application_dir(), number)
        except ProfileError as error:
            self.status_var.set(str(error))
        self._refresh_history_row()
        if model and self.preferences.get("speak_on_query"):
            speak_text(model)

    def _refresh_history_row(self) -> None:
        for child in self.history_row.winfo_children():
            child.destroy()
        if not self.history:
            ttk.Label(self.history_row, text="尚无查询历史", style="CardMuted.TLabel").pack(side="left")
            return
        ttk.Label(self.history_row, text="最近：", style="CardMuted.TLabel").pack(side="left", padx=(0, self.theme.px(4)))
        for number in self.history[:5]:
            link = ttk.Label(self.history_row, text=number, style="Chip.TLabel", cursor="hand2")
            link.pack(side="left", padx=(0, self.theme.px(4)))
            link.bind("<Button-1>", lambda _event, value=number: self.use_history(value))
        if len(self.history) > 5:
            more = ttk.Label(self.history_row, text="更多…", style="CardMuted.TLabel", cursor="hand2")
            more.pack(side="left")
            more.bind("<Button-1>", lambda _event: self.show_history_menu())

    def use_history(self, number: str) -> None:
        self.number_var.set(number)
        self.query(commit=True)
        self.entry.focus_set()

    def show_history_menu(self) -> None:
        if not self.history:
            return
        menu = tk.Menu(self, tearoff=False)
        for number in self.history:
            model = find_model(int(number), self.ranges) or "未收录"
            menu.add_command(label=f"{number}  {model}", command=lambda value=number: self.use_history(value))
        try:
            menu.tk_popup(self.entry.winfo_rootx(), self.entry.winfo_rooty() + self.entry.winfo_height())
        finally:
            menu.grab_release()

    def record_spot(self) -> None:
        if self.current_model is None:
            messagebox.showinfo("记一笔", "请先查询到一个车型。", parent=self)
            return
        number = self.number_var.get().strip()
        cover = str(self.profiles.get(self.current_model, {}).get("image", "") or "")
        SpotDialog(self, number=number, model=self.current_model, cover=cover)

    def _on_clipboard_number(self, number: int) -> None:
        """剪贴板里出现了新的自编号。"""
        self.deiconify()
        self.lift()
        self.number_var.set(f"{number:05d}")
        self.query(commit=True)
        self.status_var.set(f"剪贴板里发现自编号 {number:05d}")

    # -- 偏好设置 --------------------------------------------------------- #
    def preference(self, key: str, default: bool = False) -> bool:
        value = self.preferences.get(key)
        return default if value is None else bool(value)

    def set_preference(self, key: str, value: object) -> None:
        self.preferences[key] = value
        self.settings["preferences"] = self.preferences
        save_settings(application_dir(), self.settings)
        self._apply_preferences()

    def _apply_preferences(self) -> None:
        if self.preference("clipboard_watch"):
            self.clipboard_watcher.start()
        else:
            self.clipboard_watcher.stop()
        want_hotkey = self.preference("hotkey_enabled")
        if want_hotkey and (self.hotkey is None or not self.hotkey.registered):
            self.hotkey = self.hotkey or GlobalHotkey(self, self._on_hotkey)
            self.hotkey.start()
        elif not want_hotkey and self.hotkey is not None:
            self.hotkey.stop()
            self.hotkey = None

    def _on_hotkey(self) -> None:
        self.deiconify()
        self.lift()
        self.entry.focus_set()
        self.entry.selection_range(0, "end")
        self.status_var.set("热键已按下：输入自编号后会立刻查询")

    def hotkey_status(self) -> str:
        if self.hotkey is None:
            return "未启用"
        return f"已启用（{self.hotkey.description}）" if self.hotkey.registered else "启用失败：热键可能被其它程序占用"

    def copy_result(self) -> None:
        if self.current_model is None:
            return
        self.clipboard_clear()
        self.clipboard_append(self.current_model)
        self.status_var.set(f"已复制车型 {self.current_model}")
        original = "复制车型"
        self.copy_button.configure(text="已复制 ✓")
        self.after(1200, lambda: self.copy_button.configure(text=original) if self.copy_button.winfo_exists() else None)
        self.update()
        self.status_var.set(f"已复制到剪贴板：{self.current_model}")

    def open_catalog(self) -> None:
        if not self.ranges:
            messagebox.showerror("车型目录", "请先修复 Excel 数据文件后再打开车型目录。", parent=self)
            return
        CatalogWindow(self)

    def open_health(self) -> None:
        if not self.ranges:
            messagebox.showerror("数据体检", "请先修复 Excel 数据文件后再做数据体检。", parent=self)
            return
        DataHealthWindow(self)

    def reload_data(self) -> None:
        """从磁盘重新读取全部数据（恢复备份后用），并让界面回到一致状态。"""
        self._load_data()
        self._reset_result()
        self.notify_data_changed({"all"})

    def open_current_model_detail(self) -> None:
        if self.current_model is None:
            return
        ModelDetailWindow(self, self.current_model)

    def open_range_manager(self) -> None:
        if not self.ranges:
            messagebox.showerror("管理号段", "请先修复 Excel 数据文件后再管理号段。", parent=self)
            return
        RangeManagerWindow(self)

    def update_ranges(self, ranges: list[BusRange]) -> None:
        self.ranges = sorted(ranges, key=lambda item: item.start)
        self._refresh_subtitle()
        self.status_var.set(f"号段已保存，共 {len(self.ranges)} 个")

    def open_metro(self) -> None:
        if self.metro_ranges:
            MetroWindow(self)
        else:
            messagebox.showerror("地铁编号查询", f"没有可用的地铁数据，请检查 {METRO_SOURCE_FILENAME}。", parent=self)

    def open_reconcile(self) -> None:
        """我的记录：导入线路 + 对账 + 缺的编号（只读你的记录文件）。"""
        MyRecordsWindow(self)

    def open_district_bus(self) -> None:
        """区属公交：预留入口（暂时只说明要准备什么数据）。"""
        DistrictBusWindow(self)

    def open_online_query(self) -> None:
        """联网查询：只读，不写任何数据文件。"""
        OnlineQueryWindow(self)

    def open_line_info(self) -> None:
        """线路资料（联网）：首末站、用过的车、现在在跑的车。"""
        LineInfoWindow(self)

    def open_line_query(self) -> None:
        LineQueryWindow(self)

    def open_line_manager(self) -> None:
        LineManagerWindow(self)

    def update_metro_ranges(self, ranges: list[MetroRange]) -> None:
        self.metro_ranges = ranges
        self.status_var.set(f"地铁号段已保存，共 {len(ranges)} 条")

    def _refresh_subtitle(self) -> None:
        if self.ranges:
            self.subtitle_var.set(f"{len(self.ranges)} 个号段 · {len(model_names(self.ranges))} 种车型")

    def close_application(self) -> None:
        """Close child dialogs before tearing down the Tk event loop."""
        if self._closing:
            return
        dirty = self.dirty_windows()
        if dirty:
            answer = messagebox.askyesnocancel(
                "尚未保存",
                f"有 {len(dirty)} 个窗口的资料还没保存，是否现在保存？",
                parent=self,
            )
            if answer is None:
                return
            if answer:
                for window in dirty:
                    saver = getattr(window, "save", None)
                    if callable(saver):
                        saver()
                if self.dirty_windows():
                    return  # 保存没成功就留在程序里，别把改动丢掉
        self._closing = True
        for child in self.winfo_children():
            try:
                child.destroy()
            except tk.TclError:
                pass
        self.quit()
        self.destroy()

    # -- 窗口与数据变更通知 ----------------------------------------------- #
    def iter_toplevels(self) -> list[tk.Toplevel]:
        """递归收集所有已打开的子窗口（含挂在别的窗口下面的对话框）。"""
        found: list[tk.Toplevel] = []
        stack: list[tk.Misc] = list(self.winfo_children())
        while stack:
            widget = stack.pop()
            if isinstance(widget, tk.Toplevel):
                found.append(widget)
            try:
                stack.extend(widget.winfo_children())
            except tk.TclError:
                continue
        return found

    def notify_data_changed(self, kinds: set[str] | None = None) -> None:
        """数据变了以后，让所有打开着的窗口自己刷新。"""
        changed = kinds or {"all"}
        for window in self.iter_toplevels():
            hook = getattr(window, "on_data_changed", None)
            if callable(hook):
                try:
                    hook(changed)
                except tk.TclError:
                    continue

    def dirty_windows(self) -> list[tk.Toplevel]:
        """有未保存修改的窗口（目前是车型资料窗口）。"""
        result: list[tk.Toplevel] = []
        for window in self.iter_toplevels():
            checker = getattr(window, "is_dirty", None)
            if not callable(checker):
                continue
            try:
                if checker():
                    result.append(window)
            except tk.TclError:
                continue
        return result

    def set_tag_overrides(self, overrides: dict, *, notify: bool = True) -> None:
        """统一入口：写入标签覆盖并递增版本号，让各窗口的缓存失效。"""
        self.tag_overrides = overrides
        self.tag_revision += 1
        if notify:
            self.notify_data_changed({"tags"})


# --------------------------------------------------------------------------- #
# Quick lookup tools: batch numbers and number-range segmentation
# --------------------------------------------------------------------------- #


class QuickLookupWindow(WindowMixin, tk.Toplevel):
    """批量编号查询 + 号段区间查询，结果可复制或导出 CSV。"""

    def __init__(self, app: BusLookupApp, initial_tab: int = 0) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "快查工具", 760, 640)
        pad = self.theme.px
        self.missing_numbers: list[int] = []
        self.missing_segments: list[tuple[int, int]] = []
        self.last_model = ""

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(12)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        self.notebook.add(self._build_batch_tab(), text="批量编号")
        self.notebook.add(self._build_range_tab(), text="号段区间")
        self.notebook.select(max(0, min(initial_tab, 1)))

        footer = ttk.Frame(frame)
        footer.grid(row=1, column=0, sticky="ew", pady=(pad(8), 0))
        self.feedback_var = tk.StringVar(value="")
        ttk.Label(footer, textvariable=self.feedback_var, style="Muted.TLabel", wraplength=pad(520)).pack(side="left")

        self.batch_empty.update()
        self.range_empty.update()
        self.batch_text.focus_set()

    # -- shared building blocks ------------------------------------------ #
    def _rows_of(self, tree: ttk.Treeview) -> list[tuple[str, ...]]:
        return tree_rows(tree)

    def _copy_rows(self, tree: ttk.Treeview) -> None:
        count = copy_tree_rows(self, tree)
        self.feedback_var.set(f"已复制 {count} 行到剪贴板（可直接粘贴到 Excel）。" if count else "没有可复制的结果，请先查询。")

    def _export_rows(self, tree: ttk.Treeview, default_name: str) -> None:
        if not tree_rows(tree):
            self.feedback_var.set("没有可导出的结果，请先查询。")
            return
        path = export_tree_rows(self, tree, default_name)
        if path:
            self.feedback_var.set(f"已导出 {len(tree_rows(tree))} 行到 {path}")

    # -- batch tab ------------------------------------------------------- #
    def _build_batch_tab(self) -> ttk.Frame:
        pad = self.theme.px
        tab = ttk.Frame(self.notebook, padding=(pad(14), pad(12)))
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)

        ttk.Label(tab, text="每行一个自编号，也可以用空格、逗号分隔；重复的会自动去重。"
                           "结果按你输入的顺序排（点表头可升/降序，再点一下回到输入顺序）。",
                  style="Muted.TLabel", wraplength=self.theme.px(680), justify="left").grid(
            row=0, column=0, sticky="w", pady=(0, pad(6))
        )
        self.batch_text = self.theme.text_box(tab, 6)
        self.batch_text.grid(row=1, column=0, sticky="ew")
        self.batch_text.bind("<Control-Return>", lambda _event: (self.run_batch(), "break")[1])

        controls = ttk.Frame(tab)
        controls.grid(row=2, column=0, sticky="ew", pady=(pad(8), pad(10)))
        ttk.Button(controls, text="查询", style="Primary.TButton", command=self.run_batch).pack(side="left")
        ttk.Button(controls, text="清空", style="Secondary.TButton", command=self.clear_batch).pack(side="left", padx=pad(8))
        self.batch_fill_button = ttk.Button(controls, text="补录未收录编号", style="Secondary.TButton", command=self.fill_missing)
        self.batch_fill_button.pack(side="left")
        self.batch_fill_button.state(["disabled"])
        ttk.Button(controls, text="⇩ 导出 CSV", style="Secondary.TButton",
                   command=lambda: self._export_rows(self.batch_tree, "批量查询结果.csv")).pack(side="right")
        ttk.Button(controls, text="复制车型", style="Secondary.TButton",
                   command=lambda: self.copy_batch_models()).pack(side="right", padx=pad(8))
        ttk.Button(controls, text="复制选中行", style="Secondary.TButton",
                   command=lambda: self._copy_rows(self.batch_tree)).pack(side="right")

        self.batch_table = DataTable(
            tab,
            self.theme,
            (
                TableColumn("number", "自编号", 110, "center"),
                TableColumn("model", "车型", 220, "w"),
                TableColumn("range", "所属号段", 170, "center"),
            ),
            sort_on_load=False,  # 按输入顺序显示：好对着自己的清单复制
            empty_text="粘贴自编号后点「查询」",
        )
        self.batch_tree_card = self.batch_table.card
        self.batch_tree_card.grid(row=3, column=0, sticky="nsew")
        self.batch_tree = self.batch_table.tree
        self.batch_empty = self.batch_table.empty
        self.batch_tree.tag_configure("missing", foreground=self.theme.accent)
        self.batch_empty.update()
        self.batch_summary_var = tk.StringVar(value="等待查询")
        ttk.Label(tab, textvariable=self.batch_summary_var, style="Muted.TLabel").grid(row=5, column=0, sticky="w", pady=(pad(6), 0))
        return tab

    def clear_batch(self) -> None:
        self.batch_text.delete("1.0", "end")
        self.batch_table.set_rows([], keep_selection=False)
        self.batch_summary_var.set("等待查询")
        self.feedback_var.set("")
        self.batch_empty.update("粘贴自编号后点「查询」")
        self.batch_text.focus_set()

    def run_batch(self) -> None:
        numbers, ignored = parse_number_list(self.batch_text.get("1.0", "end-1c"))
        self.missing_numbers = []
        if not numbers and not ignored:
            self.batch_table.set_rows([], keep_selection=False)
            self.batch_summary_var.set("请先粘贴或输入自编号。")
            self.batch_empty.update("粘贴自编号后点「查询」")
            return
        hits = 0
        last_model = ""
        rows: list[tuple[str, str, str]] = []
        for number in numbers:
            item = find_range(number, self.app.ranges)
            if item is None:
                rows.append((f"{number:05d}", "未收录", "—"))
                self.missing_numbers.append(number)
            else:
                hits += 1
                last_model = item.model
                rows.append((f"{number:05d}", item.model, f"{item.start}-{item.end}"))
        self.batch_table.set_rows(rows, keep_selection=False)
        for item_id in self.batch_tree.get_children():
            if str(self.batch_tree.item(item_id, "values")[1]) == "未收录":
                self.batch_tree.item(item_id, tags=tuple(self.batch_tree.item(item_id, "tags")) + ("missing",))
        parts = [f"共 {len(numbers)} 条编号", f"命中 {hits}", f"未收录 {len(numbers) - hits}"]
        if ignored:
            parts.append(f"忽略 {len(ignored)} 条无效内容（{'、'.join(ignored[:5])}{'…' if len(ignored) > 5 else ''}）")
        self.batch_summary_var.set(" · ".join(parts))
        self.feedback_var.set("")
        self.last_model = last_model
        self._sync_fill_button()
        self.batch_empty.update()

    def copy_batch_models(self) -> None:
        """只把车型复制出来（对着清单核对时常用）。"""
        count = self.batch_table.copy_column("model")
        if not count:
            self.batch_summary_var.set("没有可复制的车型：先查询一下。")
            return
        self.batch_summary_var.set(f"已复制 {count} 个车型到剪贴板（只复制车型这一列）")

    def _sync_fill_button(self) -> None:
        if self.missing_numbers:
            self.batch_fill_button.state(["!disabled"])
        else:
            self.batch_fill_button.state(["disabled"])

    def fill_missing(self) -> None:
        """把未收录的自编号合并成区间后批量补录。"""
        intervals = group_missing_numbers(self.missing_numbers)
        if not intervals:
            messagebox.showinfo("补录号段", "当前没有未收录的自编号。", parent=self)
            return
        BulkRangeWindow(self, intervals, self.last_model)

    def _fill_selected_segment(self, _event: tk.Event | None = None) -> None:
        """双击区间查询里的「未收录」段 → 预填新增号段。"""
        selection = self.range_tree.selection()
        if not selection:
            return
        iid = str(selection[0])
        if not iid.startswith("missing-"):
            return
        index = int(iid.split("-", 1)[1])
        if 0 <= index < len(self.missing_segments):
            start, end = self.missing_segments[index]
            RangeEditorWindow(self, None, (start, end, self.last_model))

    def refresh(self) -> None:
        """数据变化后重跑当前标签页的查询。"""
        if self.notebook.index("current") == 1:
            self.run_range()
        elif self.batch_summary_var.get() != "等待查询":
            self.run_batch()

    # -- range tab ------------------------------------------------------- #
    def _build_range_tab(self) -> ttk.Frame:
        pad = self.theme.px
        tab = ttk.Frame(self.notebook, padding=(pad(14), pad(12)))
        tab.columnconfigure(5, weight=1)
        tab.rowconfigure(2, weight=1)

        ttk.Label(tab, text="输入一段自编号区间，自动按车型切分成号段（未收录的部分也会列出）。", style="Muted.TLabel").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, pad(6))
        )
        self.range_start_var = tk.StringVar()
        self.range_end_var = tk.StringVar()
        ttk.Label(tab, text="起始").grid(row=1, column=0, sticky="w")
        start_entry = ttk.Entry(tab, textvariable=self.range_start_var, width=10, justify="center", style="Query.TEntry")
        start_entry.grid(row=1, column=1, sticky="w", padx=(pad(6), pad(14)))
        ttk.Label(tab, text="结束").grid(row=1, column=2, sticky="w")
        end_entry = ttk.Entry(tab, textvariable=self.range_end_var, width=10, justify="center", style="Query.TEntry")
        end_entry.grid(row=1, column=3, sticky="w", padx=(pad(6), pad(14)))
        for entry in (start_entry, end_entry):
            entry.bind("<Return>", lambda _event: self.run_range())
        ttk.Button(tab, text="查询", style="Primary.TButton", command=self.run_range).grid(row=1, column=4, sticky="w")

        controls = ttk.Frame(tab)
        controls.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(pad(10), pad(10)))
        ttk.Button(controls, text="⇩ 导出 CSV", style="Secondary.TButton",
                   command=lambda: self._export_rows(self.range_tree, "号段区间结果.csv")).pack(side="right")
        ttk.Button(controls, text="复制选中行", style="Secondary.TButton",
                   command=lambda: self._copy_rows(self.range_tree)).pack(side="right", padx=pad(8))

        self.range_table = DataTable(
            tab,
            self.theme,
            (
                TableColumn("segment", "号段", 190, "center"),
                TableColumn("count", "数量", 80, "center", numeric=True),
                TableColumn("model", "车型", 240, "w"),
            ),
            empty_text="输入一段区间后点「查询」",
        )
        self.range_tree_card = self.range_table.card
        self.range_tree_card.grid(row=2, column=0, columnspan=6, sticky="nsew")
        self.range_tree = self.range_table.tree
        self.range_empty = self.range_table.empty
        self.range_tree.tag_configure("missing", foreground=self.theme.accent)
        self.range_tree.bind("<Double-Button-1>", self._fill_selected_segment)
        self.range_summary_var = tk.StringVar(value="等待查询")
        ttk.Label(tab, textvariable=self.range_summary_var, style="Muted.TLabel").grid(row=4, column=0, columnspan=6, sticky="w", pady=(pad(6), 0))
        return tab

    def run_range(self) -> None:
        start = validate_number(self.range_start_var.get())
        end = validate_number(self.range_end_var.get())
        if start is None or end is None:
            messagebox.showerror("号段区间查询", "起始和结束都必须是五位数字。", parent=self)
            return
        if start > end:
            messagebox.showerror("号段区间查询", "起始编号不能大于结束编号。", parent=self)
            return
        segments = segment_lookup(start, end, self.app.ranges)
        self.missing_numbers = []
        self.missing_segments = []
        covered = 0
        rows: list[tuple[str, str, str]] = []
        ids: list[str] = []
        for index, (segment_start, segment_end, model) in enumerate(segments):
            count = segment_end - segment_start + 1
            if model is None:
                rows.append((f"{segment_start}-{segment_end}", str(count), "未收录"))
                self.missing_numbers.extend(range(segment_start, segment_end + 1))
                self.missing_segments.append((segment_start, segment_end))
                ids.append(f"missing-{len(self.missing_segments) - 1}")
                continue
            covered += count
            rows.append((f"{segment_start}-{segment_end}", str(count), model))
            ids.append(f"segment-{index}")
        self.range_table.set_rows(rows, keep_selection=False, iid=keep_row_ids(rows, ids))
        for item_id in self.range_tree.get_children():
            if str(self.range_tree.item(item_id, "values")[2]) == "未收录":
                self.range_tree.item(item_id, tags=tuple(self.range_tree.item(item_id, "tags")) + ("missing",))
        total = end - start + 1
        self.range_summary_var.set(
            f"区间 {start}-{end} 共 {total} 个编号 · 切分 {len(segments)} 段 · 已收录 {covered} · 未收录 {total - covered}"
        )
        self.feedback_var.set("")
        self._sync_fill_button()
        self.range_empty.update()


# --------------------------------------------------------------------------- #
# Bus model catalog and detail windows
# --------------------------------------------------------------------------- #


class CatalogWindow(WindowMixin, tk.Toplevel):
    """车型目录：关键字 + 厂商 / 燃料 / 车长筛选，表头可点击排序。"""

    COLUMNS = (
        ("model", "车型", 200, "w"),
        ("manufacturer", "厂商", 105, "w"),
        ("fuel", "燃料", 85, "center"),
        ("length", "车长", 75, "center"),
        ("body", "车体形式", 80, "center"),
        ("ranges", "号段", 55, "e"),
        ("lines", "线路", 55, "e"),
        ("info", "资料", 75, "center"),
    )

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "车型目录", 860, 620)
        pad = self.theme.px
        self.filter_var = tk.StringVar()
        self.manufacturer_var = tk.StringVar(value=ANY_TAG)
        self.fuel_var = tk.StringVar(value=ANY_TAG)
        self.length_var = tk.StringVar(value=ANY_TAG)
        self.body_var = tk.StringVar(value=ANY_TAG)
        self.line_value_var = tk.StringVar(value=ANY_TAG)  # 「线路」筛选：某条线用过哪些车型
        self.count_var = tk.StringVar()
        self.shown_models: list[str] = []
        self._tag_cache: dict[str, ModelTags] = {}
        self._tag_cache_key: tuple = ()
        self._counts: dict[str, int] = {}
        self._image_names: set[str] = set()
        self._alias_hits: dict[str, str] = {}
        self._debounced_refresh = debounce(self, "catalog-search", 150, self._refresh)

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        ttk.Label(frame, text="车型目录", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(frame, textvariable=self.count_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")

        search_row = ttk.Frame(frame)
        search_row.grid(row=1, column=0, sticky="ew", pady=(pad(8), pad(6)))
        search_row.columnconfigure(1, weight=1)
        ttk.Label(search_row, text="搜索").grid(row=0, column=0, padx=(0, pad(8)))
        search = ttk.Entry(search_row, textvariable=self.filter_var, style="Search.TEntry")
        search.grid(row=0, column=1, sticky="ew")
        attach_placeholder(search, "代号 / 厂商 / 俗称 / 拼音首字母", self.theme)
        self.bind("<Control-f>", lambda _event: (search.focus_set(), "break")[1])

        filter_row = ttk.Frame(frame)
        filter_row.grid(row=2, column=0, sticky="ew", pady=(0, pad(10)))
        self.filter_combos: dict[str, ttk.Combobox] = {}
        self._build_filter(filter_row, "厂商", self.manufacturer_var, [tag.manufacturer for tag in self._tags().values()])
        self._build_filter(filter_row, "燃料", self.fuel_var, [tag.fuel for tag in self._tags().values()])
        self._build_filter(filter_row, "车长", self.length_var, [tag.length for tag in self._tags().values()])
        self._build_filter(filter_row, "车体", self.body_var, [tag.body for tag in self._tags().values()])
        self._build_filter(
            filter_row, "线路", self.line_value_var, [link.line for link in self.app.lines]
        )  # 「某条线用过哪些车型」：选项来自你自己的线路数据
        ttk.Button(filter_row, text="重置", style="Secondary.TButton", command=self._reset_filters).pack(side="left", padx=(pad(4), 0))

        self.table = DataTable(
            frame,
            self.theme,
            self._columns(),
            selectmode="extended",
            sort_key=self._sort_key,
            empty_text="没有车型",
        )
        self.table.card.grid(row=3, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.empty_state = self.table.empty
        self.tree.bind("<Double-Button-1>", self._open_selected)
        self.tree.bind("<Return>", self._open_selected)

        footer = ttk.Frame(frame)
        footer.grid(row=4, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Label(
            footer,
            text="双击查看资料 · Ctrl / Shift 可多选后批量设置标签 · 点表头排序 · 标签可在此或 车型标签.json 修正",
            style="Muted.TLabel",
        ).pack(side="left")
        ttk.Button(footer, text="查看车型详情", style="Primary.TButton", command=self._open_selected).pack(side="right")
        ttk.Button(footer, text="批量设置标签", style="Secondary.TButton", command=self.open_bulk_tags).pack(
            side="right", padx=pad(8)
        )

        for variable in (self.filter_var, self.manufacturer_var, self.fuel_var, self.length_var, self.body_var, self.line_value_var):
            variable.trace_add("write", self._debounced_refresh)
        self._refresh()
        search.focus_set()

    def on_data_changed(self, kinds: set[str]) -> None:
        if kinds & {"all", "tags", "ranges", "profiles", "lines"}:
            self._refresh()

    def _tags(self) -> dict[str, ModelTags]:
        names = tuple(model_names(self.app.ranges))
        key = (names, self.app.tag_revision)
        if key != self._tag_cache_key:
            self._tag_cache = {name: model_tags(name, self.app.tag_overrides) for name in names}
            self._tag_cache_key = key
        return self._tag_cache

    def _tag_options(self, values: list[str]) -> list[str]:
        return [ANY_TAG] + sorted(set(values), key=lambda value: (value == UNKNOWN_TAG, value))

    def _build_filter(self, parent: tk.Misc, label: str, variable: tk.StringVar, values: list[str]) -> None:
        pad = self.theme.px
        holder = ttk.Frame(parent)
        holder.pack(side="left", padx=(0, pad(10)))
        ttk.Label(holder, text=label).pack(side="left", padx=(0, pad(4)))
        combo = ttk.Combobox(holder, textvariable=variable, values=self._tag_options(values), state="readonly", width=10)
        combo.pack(side="left")
        self.filter_combos[label] = combo

    def _reset_filters(self) -> None:
        self.filter_var.set("")
        self.manufacturer_var.set(ANY_TAG)
        self.fuel_var.set(ANY_TAG)
        self.length_var.set(ANY_TAG)
        self.body_var.set(ANY_TAG)
        self.line_value_var.set(ANY_TAG)

    @classmethod
    def _columns(cls) -> tuple[TableColumn, ...]:
        return tuple(
            TableColumn(key, title, width, anchor, numeric=key == "ranges")
            for key, title, width, anchor in cls.COLUMNS
        )

    def _sort_key(self, column: str, row: tuple[str, ...]) -> object:
        """表头排序：厂商 / 燃料要按标签、车长按数字、车体按固定次序。"""
        label, manufacturer, fuel, length, body, counts, line_count, info = row  # noqa: F841
        name = label.casefold()
        if column == "manufacturer":
            return (manufacturer.casefold(), name)
        if column == "fuel":
            return (fuel.casefold(), name)
        if column == "length":
            digits = re.sub(r"\D", "", length)
            return (int(digits) if digits else 999, name)
        if column == "body":
            order = {"铰接": 0, "双层": 1, "单机": 2}
            return (order.get(body, 9), name)
        if column == "ranges":
            return (int(counts) if counts.isdigit() else 0, name)
        if column == "lines":
            return (int(line_count) if line_count.isdigit() else 0, name)
        if column == "info":
            return (0 if info == "已有资料" else 1, name)
        return (name,)

    def _refresh(self, *_args: object) -> None:
        query = self.filter_var.get().strip().casefold()
        tags = self._tags()
        # 一次算好每车型号段数与现有图片名，避免逐行扫描 / 逐行 stat
        self._counts = Counter(item.model for item in self.app.ranges)
        self._image_names = existing_image_names(application_dir())
        # 标签可能刚被改过：刷新下拉里的可选项，并在当前选择失效时回到「全部」
        for label, field, variable in (
            ("厂商", "manufacturer", self.manufacturer_var),
            ("燃料", "fuel", self.fuel_var),
            ("车长", "length", self.length_var),
            ("车体", "body", self.body_var),
        ):
            options = self._tag_options([getattr(tag, field) for tag in tags.values()])
            combo = self.filter_combos.get(label)
            if combo is not None:
                combo.configure(values=options)
            if variable.get() not in options:
                variable.set(ANY_TAG)
        manufacturer = self.manufacturer_var.get()
        fuel = self.fuel_var.get()
        length = self.length_var.get()
        body = self.body_var.get()
        # 线路数据 → 每个车型跑过哪些线路（只算一次，别逐行扫）
        lines_by_model: dict[str, list[str]] = {}
        for link in self.app.lines:
            if link.model:
                lines_by_model.setdefault(link.model, []).append(link.line)
        # 「线路」筛选的下拉要跟着线路数据变
        line_names = [ANY_TAG] + sorted({link.line for link in self.app.lines}, key=lambda value: (len(value), value))
        combo = self.filter_combos.get("线路")
        if combo is not None:
            combo.configure(values=line_names)
        line_filter = self.line_value_var.get()
        if line_filter not in line_names:
            line_filter = ANY_TAG
            self.line_value_var.set(ANY_TAG)
        names = model_names(self.app.ranges)
        aliases = {name: model_aliases(name, self.app.tag_overrides) for name in names}
        self._alias_hits = {}
        for name in names:
            for alias in aliases[name]:
                if query and query in alias.casefold():
                    self._alias_hits[name] = alias
                    break
        self.shown_models = [
            name
            for name in names
            if (not query or search_hit(query, name, tags[name].manufacturer, aliases[name]) or name in self._alias_hits)
            and (manufacturer == ANY_TAG or tags[name].manufacturer == manufacturer)
            and (fuel == ANY_TAG or tags[name].fuel == fuel)
            and (length == ANY_TAG or tags[name].length == length)
            and (body == ANY_TAG or tags[name].body == body)
            and (line_filter == ANY_TAG or line_filter in lines_by_model.get(name, ()))
        ]

        rows: list[tuple] = []
        ids: list[str] = []
        for model in self.shown_models:
            profile = self.app.profiles.get(model, {})
            tag = tags[model]
            label = model + (f"（{self._alias_hits[model]}）" if model in self._alias_hits else "")
            rows.append(
                (
                    label,
                    tag.manufacturer,
                    tag.fuel,
                    tag.length,
                    tag.body,
                    self._counts.get(model, 0),
                    len(lines_by_model.get(model, ())),
                    "已有资料" if not profile_is_empty(profile, None, self._image_names) else "暂无资料",
                )
            )
            ids.append(model)
        self.table.set_rows(rows, keep_selection=False, iid=keep_row_ids(rows, ids))
        total = len(names)
        self.count_var.set(f"共 {total} 种" if len(self.shown_models) == total else f"匹配 {len(self.shown_models)} / {total} 种")
        self.empty_state.update(
            "没有匹配的车型\n试试清空搜索词或把筛选改回「全部」" if not self.shown_models else None,
            has_rows=bool(self.shown_models),
        )

    def _open_selected(self, _event: tk.Event | None = None) -> None:
        selection = self.tree.selection()
        if selection:
            ModelDetailWindow(self, selection[0])

    def open_bulk_tags(self) -> None:
        models = list(self.tree.selection())
        if not models:
            messagebox.showinfo("批量设置标签", "请先在列表里选中一个或多个车型（按住 Ctrl / Shift 可多选）。", parent=self)
            return
        BulkTagWindow(self, models)


class BulkTagWindow(WindowMixin, tk.Toplevel):
    """给选中的一批车型一次性设置厂商 / 燃料 / 车长。"""

    FIELDS = (("manufacturer", "厂商"), ("fuel", "燃料"), ("length", "车长"), ("body", "车体形式"))

    def __init__(self, catalog: CatalogWindow, models: list[str]) -> None:
        super().__init__(catalog)
        self.catalog = catalog
        self.app = catalog.app
        self.theme = self.app.theme
        self.models = list(models)
        prefixes = {manufacturer_key(name) for name in self.models}
        self.prefix = prefixes.pop() if len(prefixes) == 1 and prefixes else ""
        self.setup_dialog(catalog, f"批量设置标签（{len(self.models)} 个车型）", 540, 430, resizable=False)
        pad = self.theme.px

        suggestions = tag_suggestions(self.app.ranges, self.app.tag_overrides)
        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(14), pady=pad(14))
        body = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(2, weight=1)

        ttk.Label(body, text=f"已选 {len(self.models)} 个车型", style="CardValue.TLabel", font=self.theme.font(12, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        preview = "、".join(self.models[:4]) + ("…" if len(self.models) > 4 else "")
        ttk.Label(body, text=preview, style="CardMuted.TLabel", wraplength=pad(440)).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(pad(2), pad(10))
        )
        ttk.Label(body, text="勾选要修改的项目，只改勾上的；留空表示恢复自动推断。", style="CardMuted.TLabel").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, pad(8))
        )

        self.use: dict[str, tk.BooleanVar] = {}
        self.vars: dict[str, tk.StringVar] = {}
        self.combos: dict[str, ttk.Combobox] = {}
        for row, (field, label) in enumerate(self.FIELDS, start=3):
            use = tk.BooleanVar(value=False)
            self.use[field] = use
            ttk.Checkbutton(body, variable=use, style="Card.TCheckbutton", command=lambda name=field: self._sync(name)).grid(
                row=row, column=0, sticky="w", pady=pad(4)
            )
            ttk.Label(body, text=label, style="CardMuted.TLabel").grid(row=row, column=1, sticky="w", padx=(0, pad(8)))
            variable = tk.StringVar()
            self.vars[field] = variable
            combo = ttk.Combobox(body, textvariable=variable, values=suggestions[field], state="disabled")
            combo.grid(row=row, column=2, sticky="ew", pady=pad(4))
            self.combos[field] = combo

        next_row = 3 + len(self.FIELDS)
        self.prefix_var = tk.BooleanVar(value=False)
        if self.prefix:
            ttk.Checkbutton(
                body,
                text=f"同时把「{self.prefix}」前缀的所有车型厂商都设为上面的值",
                variable=self.prefix_var,
                style="Card.TCheckbutton",
            ).grid(row=next_row, column=0, columnspan=3, sticky="w", pady=(pad(10), 0))

        actions = ttk.Frame(body, style="Card.TFrame")
        actions.grid(row=next_row + 1, column=0, columnspan=3, sticky="ew", pady=(pad(14), 0))
        ttk.Button(actions, text="应用", style="Primary.TButton", command=self.apply).pack(side="right")
        ttk.Button(actions, text="取消", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))

    def _sync(self, field: str) -> None:
        self.combos[field].configure(state="normal" if self.use[field].get() else "disabled")

    def apply(self) -> None:
        values = {field: self.vars[field].get().strip() for field, _label in self.FIELDS if self.use[field].get()}
        if not values:
            messagebox.showinfo("批量设置标签", "请先勾选要修改的项目。", parent=self)
            return
        overrides = clean_tag_overrides(self.app.tag_overrides if isinstance(self.app.tag_overrides, dict) else {})
        for model in self.models:
            apply_model_tag_values(overrides, model, values)
        if self.prefix and self.prefix_var.get():
            apply_manufacturer_prefix(overrides, self.prefix, values.get("manufacturer", ""))
        cleaned = clean_tag_overrides(overrides)
        try:
            save_tag_overrides(application_dir(), cleaned)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.set_tag_overrides(cleaned)
        self.app.status_var.set(f"已批量更新 {len(self.models)} 个车型的标签")
        self.destroy()


class ModelDetailWindow(WindowMixin, tk.Toplevel):
    """Model profile editor that only touches disk when the user saves."""

    def __init__(self, parent: tk.Misc, model: str) -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)
        self.theme = self.app.theme
        self.catalog = parent if isinstance(parent, CatalogWindow) else None
        self.model = model
        self.saved_profile = dict(self.app.profiles.get(model, {}))
        self.saved_images: list[dict[str, str]] = profile_images(self.saved_profile)
        self.pending_images: list[dict[str, str]] = [dict(item) for item in self.saved_images]
        self.pending_sources: dict[str, Path | Image.Image] = {}
        self.removed_images: list[str] = []
        self.photo: ImageTk.PhotoImage | None = None
        pad = self.theme.px

        self.setup_dialog(parent, f"车型信息 - {model}", 660, 640)
        self.minsize(pad(520), pad(480))

        footer = ttk.Frame(self, padding=(pad(16), 0, pad(16), pad(14)))
        footer.pack(side="bottom", fill="x")
        ttk.Separator(footer).pack(fill="x", pady=(0, pad(10)))
        self.dirty_var = tk.StringVar(value="")
        ttk.Label(footer, textvariable=self.dirty_var, style="Section.TLabel", background=self.theme.background).pack(side="left")
        ttk.Button(footer, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")
        self.revert_button = ttk.Button(footer, text="还原", style="Secondary.TButton", command=self.revert)
        self.revert_button.pack(side="right", padx=pad(8))
        self.save_button = ttk.Button(footer, text="保存", style="Primary.TButton", command=self.save)
        self.save_button.pack(side="right")
        ttk.Button(footer, text="⇩ 导出卡片", style="Secondary.TButton", command=self.export_card).pack(side="left", padx=pad(10))

        content = ttk.Frame(self)
        content.pack(side="top", fill="both", expand=True)
        canvas = tk.Canvas(content, highlightthickness=0, background=self.theme.background, bd=0)
        scrollbar = ttk.Scrollbar(content, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        frame = ttk.Frame(canvas, padding=(pad(16), pad(14)))
        window = canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        self.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-1 * (event.delta // 120), "units"))
        self.canvas = canvas

        ttk.Label(frame, text=model, style="Result.TLabel", background=self.theme.background, foreground=self.theme.text, wraplength=pad(560)).pack(
            anchor="w"
        )
        known = ranges_for_model(model, self.app.ranges)
        total = sum(item.end - item.start + 1 for item in known)
        self.tag_var = tk.StringVar()
        tag_row = ttk.Frame(frame)
        tag_row.pack(anchor="w", fill="x", pady=(pad(4), 0))
        ttk.Label(tag_row, textvariable=self.tag_var, style="Muted.TLabel").pack(side="left")
        ttk.Button(tag_row, text="编辑标签", style="Secondary.TButton", command=lambda: TagEditorWindow(self)).pack(
            side="left", padx=(pad(10), 0)
        )
        self.refresh_tags()
        ttk.Label(
            frame,
            text=f"共 {len(known)} 个号段 · 覆盖 {total} 个自编号" if known else "当前 Excel 中没有该车型的号段",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(pad(2), pad(12)))

        image_card = self.theme.card(frame)
        image_card.pack(fill="x")
        image_inner = ttk.Frame(image_card, style="Card.TFrame", padding=(pad(14), pad(12)))
        image_inner.pack(fill="both", expand=True)
        ttk.Label(image_inner, text="图片（第一张为封面）", style="Section.TLabel").pack(anchor="w")
        self.image_holder = tk.Frame(image_inner, background=self.theme.surface_alt, height=pad(190))
        self.image_holder.pack(fill="x", pady=(pad(8), pad(8)))
        self.image_holder.pack_propagate(False)
        self.image_label = tk.Label(self.image_holder, background=self.theme.surface_alt, foreground=self.theme.muted, font=self.theme.font(10), text="暂无图片")
        self.image_label.pack(fill="both", expand=True)

        self.image_tree = ttk.Treeview(image_inner, columns=("caption",), show="headings", height=3, selectmode="browse")
        self.image_tree.heading("caption", text="相册", anchor="w")
        self.image_tree.column("caption", width=pad(420), anchor="w")
        self.image_tree.tag_configure("odd", background=self.theme.stripe)
        self.image_tree.pack(fill="x", pady=(0, pad(8)))
        self.image_tree.bind("<<TreeviewSelect>>", lambda _event: self._show_image())

        image_actions = ttk.Frame(image_inner, style="Card.TFrame")
        image_actions.pack(anchor="w")
        ttk.Button(image_actions, text="上传图片", style="Secondary.TButton", command=self.choose_image).pack(side="left", padx=(0, pad(6)))
        ttk.Button(image_actions, text="粘贴图片 (Ctrl+V)", style="Secondary.TButton", command=self.paste_image).pack(side="left", padx=(0, pad(6)))
        ttk.Button(image_actions, text="设为封面", style="Secondary.TButton", command=self.set_cover).pack(side="left", padx=(0, pad(6)))
        ttk.Button(image_actions, text="编辑说明", style="Secondary.TButton", command=self.edit_caption).pack(side="left", padx=(0, pad(6)))
        ttk.Button(image_actions, text="上移", style="Secondary.TButton", command=lambda: self.move_image(-1)).pack(side="left", padx=(0, pad(6)))
        ttk.Button(image_actions, text="下移", style="Secondary.TButton", command=lambda: self.move_image(1)).pack(side="left", padx=(0, pad(6)))
        ttk.Button(image_actions, text="移除选中", style="Danger.TButton", command=self.remove_image).pack(side="left")

        self.introduction = self._text_section(frame, "简介", 7)
        self.specifications = self._text_section(frame, "技术参数", 7)

        ranges_card = self.theme.card(frame)
        ranges_card.pack(fill="x", pady=(0, pad(12)))
        ranges_inner = ttk.Frame(ranges_card, style="Card.TFrame", padding=(pad(14), pad(12)))
        ranges_inner.pack(fill="both", expand=True)
        self.theme.section(ranges_inner, "已知对应号段", pady=(0, 6))
        range_text = "、".join(f"{item.start}-{item.end}" for item in known) or "暂无对应号段"
        ttk.Label(ranges_inner, text=range_text, style="CardValue.TLabel", justify="left", wraplength=pad(560)).pack(anchor="w", pady=(pad(6), 0))

        # 「我记过」的线路：数据来自 线路.json（默认空，没有就不显示这一块）
        model_lines = lines_for_model(self.app.lines, self.model)
        if model_lines:
            lines_card = self.theme.card(frame)
            lines_card.pack(fill="x", pady=(0, pad(12)))
            lines_inner = ttk.Frame(lines_card, style="Card.TFrame", padding=(pad(14), pad(12)))
            lines_inner.pack(fill="both", expand=True)
            self.theme.section(lines_inner, "记录过的线路", pady=(0, 6))
            ttk.Label(
                lines_inner,
                text="这些是你自己记过的（不代表现在还在跑）",
                style="CardMuted.TLabel",
            ).pack(anchor="w")
            for link in model_lines[:8]:
                text = f"{line_display(link)}　{LINE_STATUS_LABELS.get(link.status, link.status)}"
                if link.numbers:
                    shown = "、".join(link.numbers[:6])
                    text += f"　自编号 {shown}" + ("…" if len(link.numbers) > 6 else "")
                ttk.Label(lines_inner, text=text, style="CardValue.TLabel", wraplength=pad(560), justify="left").pack(
                    anchor="w", pady=(pad(3), 0)
                )
            if len(model_lines) > 8:
                ttk.Label(lines_inner, text=f"…还有 {len(model_lines) - 8} 条，看「线路查询」", style="CardMuted.TLabel").pack(anchor="w", pady=(pad(3), 0))

        self.introduction.insert("1.0", self.saved_profile.get("introduction", ""))
        self.specifications.insert("1.0", self.saved_profile.get("specifications", ""))
        self.introduction.edit_modified(False)
        self.specifications.edit_modified(False)
        self.introduction.bind("<<Modified>>", self._on_text_modified)
        self.specifications.bind("<<Modified>>", self._on_text_modified)
        self.bind("<Control-v>", self._on_paste)
        self._refresh_image_list(select=0 if self.pending_images else None)
        self._show_image()
        self._refresh_dirty()
        self.introduction.focus_set()

    def _text_section(self, parent: tk.Misc, title: str, height: int) -> tk.Text:
        card = self.theme.card(parent)
        card.pack(fill="both", expand=True, pady=(0, self.theme.px(12)))
        inner = ttk.Frame(card, style="Card.TFrame", padding=(self.theme.px(14), self.theme.px(12)))
        inner.pack(fill="both", expand=True)
        ttk.Label(inner, text=title, style="Section.TLabel").pack(anchor="w")
        box = self.theme.text_box(inner, height)
        box.pack(fill="both", expand=True, pady=(self.theme.px(6), 0))
        return box

    # -- state ----------------------------------------------------------- #
    def _on_text_modified(self, event: tk.Event) -> None:
        widget = event.widget
        if not isinstance(widget, tk.Text) or not widget.edit_modified():
            return
        widget.edit_modified(False)  # re-arm the flag, then re-evaluate
        self._refresh_dirty()

    def _current_profile(self) -> dict[str, object]:
        profile: dict[str, object] = {
            "introduction": self.introduction.get("1.0", "end-1c").strip(),
            "specifications": self.specifications.get("1.0", "end-1c").strip(),
        }
        set_profile_images(profile, self.pending_images)
        return profile

    def _saved_snapshot(self) -> dict[str, object]:
        profile: dict[str, object] = {
            "introduction": self.saved_profile.get("introduction", "").strip(),
            "specifications": self.saved_profile.get("specifications", "").strip(),
        }
        set_profile_images(profile, self.saved_images)
        return profile

    def is_dirty(self) -> bool:
        if self.pending_sources or self.removed_images:
            return True
        return self._current_profile() != self._saved_snapshot()

    def refresh_tags(self) -> None:
        tags = model_tags(self.model, self.app.tag_overrides)
        self.tag_var.set(f"{tags.manufacturer} · {tags.fuel} · {tags.length}")

    def on_data_changed(self, kinds: set[str]) -> None:
        if kinds & {"all", "tags", "ranges", "profiles"}:
            self.refresh_tags()

    def _refresh_dirty(self) -> None:
        dirty = self.is_dirty()
        self.save_button.state(["!disabled"] if dirty else ["disabled"])
        self.revert_button.state(["!disabled"] if dirty else ["disabled"])
        self.dirty_var.set("● 有未保存的修改" if dirty else "所有修改已保存")
        self.title(f"车型信息 - {self.model}" + (" •已修改" if dirty else ""))

    # -- images ---------------------------------------------------------- #
    def _cover_name(self) -> str:
        return self.pending_images[0]["name"] if self.pending_images else ""

    def _free_image_name(self) -> str:
        used = existing_image_names(application_dir()) | {item["name"] for item in self.pending_images}
        for index in range(1, 100):
            candidate = _safe_image_name(self.model, index)
            if candidate not in used:
                return candidate
        return _safe_image_name(self.model, 99)

    def _selected_image_index(self) -> int | None:
        selection = self.image_tree.selection()
        return int(selection[0]) if selection else None

    def _preview_name(self) -> str:
        index = self._selected_image_index()
        if index is not None and 0 <= index < len(self.pending_images):
            return self.pending_images[index]["name"]
        return self._cover_name()

    def _refresh_image_list(self, select: int | None = None) -> None:
        self.image_tree.delete(*self.image_tree.get_children())
        for index, item in enumerate(self.pending_images):
            label = f"{index + 1}. {item['name']}"
            if item.get("caption"):
                label += f"　{item['caption']}"
            if index == 0:
                label += "　（封面）"
            if item["name"] in self.pending_sources:
                label += "　[未保存]"
            self.image_tree.insert("", "end", iid=str(index), values=(label,), tags=("odd",) if index % 2 else ())
        if select is not None and 0 <= select < len(self.pending_images):
            self.image_tree.selection_set(str(select))

    def on_data_changed(self, kinds: set[str]) -> None:  # noqa: D401 - 见 WindowMixin
        if kinds & {"all", "tags", "ranges", "profiles"}:
            self.refresh_tags()

    def _on_paste(self, _event: tk.Event) -> str | None:
        focused = self.focus_get()
        if isinstance(focused, (tk.Entry, tk.Text, ttk.Entry)):
            return None  # 文本框里 Ctrl+V 仍然是粘贴文字
        self.paste_image()
        return "break"

    def paste_image(self) -> None:
        image = clipboard_image()
        if image is None:
            messagebox.showinfo("粘贴图片", "剪贴板里没有图片。先截图或复制一张图片再试。", parent=self)
            return
        self._add_image(image)

    def choose_image(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self,
            title="选择车型图片",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"), ("所有文件", "*.*")],
        )
        if not filename:
            return
        self._add_image(Path(filename))

    def _add_image(self, source: Path | Image.Image) -> None:
        if isinstance(source, Path):
            try:
                load_image(source)
            except ProfileError as error:
                messagebox.showerror("添加图片失败", str(error), parent=self)
                return
        name = self._free_image_name()
        self.pending_images.append({"name": name, "caption": ""})
        self.pending_sources[name] = source
        self._refresh_image_list(select=len(self.pending_images) - 1)
        self._show_image()
        self._refresh_dirty()

    def remove_image(self) -> None:
        index = self._selected_image_index()
        if index is None:
            if not self.pending_images:
                return
            index = 0
        item = self.pending_images.pop(index)
        self.pending_sources.pop(item["name"], None)
        if item["name"] in {saved["name"] for saved in self.saved_images}:
            self.removed_images.append(item["name"])
        self._refresh_image_list(select=min(index, len(self.pending_images) - 1) if self.pending_images else None)
        self._show_image()
        self._refresh_dirty()

    def set_cover(self) -> None:
        index = self._selected_image_index()
        if index is None or index == 0:
            return
        self.pending_images.insert(0, self.pending_images.pop(index))
        self._refresh_image_list(select=0)
        self._show_image()
        self._refresh_dirty()

    def edit_caption(self) -> None:
        """给选中的图片写一句说明（前脸、内饰、铭牌…）。"""
        index = self._selected_image_index()
        if index is None:
            if not self.pending_images:
                return
            index = 0
        current = self.pending_images[index].get("caption", "")
        caption = simpledialog.askstring("图片说明", f"{self.pending_images[index]['name']}\n\n说明（例如：前脸 / 内饰 / 铭牌）：", initialvalue=current, parent=self)
        if caption is None:
            return
        self.pending_images[index]["caption"] = caption.strip()
        self._refresh_image_list(select=index)
        self._refresh_dirty()

    def move_image(self, offset: int) -> None:
        index = self._selected_image_index()
        if index is None:
            return
        target = index + offset
        if not 0 <= target < len(self.pending_images):
            return
        self.pending_images[index], self.pending_images[target] = self.pending_images[target], self.pending_images[index]
        self._refresh_image_list(select=target)
        self._show_image()
        self._refresh_dirty()

    def _show_image(self) -> None:
        name = self._preview_name()
        source = self.pending_sources.get(name)
        if isinstance(source, Image.Image):
            self._show_preview(source)
            return
        if isinstance(source, Path):
            try:
                self._show_preview(load_image(source))
                return
            except ProfileError:
                pass
        path = stored_image_path(application_dir(), name)
        if not name or not path.is_file():
            self.photo = None
            self.image_label.configure(image="", text="暂无图片")
            return
        try:
            with Image.open(path) as image:
                self._show_preview(image.convert("RGB"))
        except (OSError, UnidentifiedImageError):
            self.photo = None
            self.image_label.configure(image="", text="图片无法读取，请重新上传")

    def _show_preview(self, image: Image.Image) -> None:
        preview = image.copy()
        preview.thumbnail((self.theme.px(430), self.theme.px(180)), Image.Resampling.LANCZOS)
        self.photo = ImageTk.PhotoImage(preview)
        self.image_label.configure(image=self.photo, text="")

    # -- persistence ----------------------------------------------------- #
    def save(self) -> None:
        base = application_dir()
        for name, source in list(self.pending_sources.items()):
            if name not in {item["name"] for item in self.pending_images}:
                continue  # 加进来又被移除了
            try:
                import_image(base, self.model, source, filename=name, delete_previous=False)
            except ProfileError as error:
                messagebox.showerror("保存失败", str(error), parent=self)
                return
        profile = self._current_profile()
        if not (profile["introduction"] or profile["specifications"] or profile["image"]):
            self.app.profiles.pop(self.model, None)
        else:
            self.app.profiles[self.model] = profile
        try:
            save_profiles(base, self.app.profiles)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        kept = {item["name"] for item in self.pending_images}
        for name in self.removed_images:
            if name in kept:
                continue
            try:
                delete_image(base, name)
            except ProfileError as error:
                messagebox.showwarning("旧图片未删除", str(error), parent=self)
        self.pending_sources.clear()
        self.removed_images = []
        self.saved_images = [dict(item) for item in self.pending_images]
        self.saved_profile = dict(profile)
        self._refresh_image_list(select=0 if self.pending_images else None)
        self._refresh_dirty()
        self._show_image()
        self.app.notify_data_changed({"profiles"})
        self.app.status_var.set(f"已保存 {self.model} 的车型资料")

    def revert(self) -> None:
        self.pending_images = [dict(item) for item in self.saved_images]
        self.pending_sources.clear()
        self.removed_images = []
        self.introduction.delete("1.0", "end")
        self.introduction.insert("1.0", self.saved_profile.get("introduction", ""))
        self.specifications.delete("1.0", "end")
        self.specifications.insert("1.0", self.saved_profile.get("specifications", ""))
        self._refresh_image_list(select=0 if self.pending_images else None)
        self._show_image()
        self._refresh_dirty()

    def export_card(self) -> None:
        """导出一个车型的单页卡片（可打印、可发出去）。"""
        if self.is_dirty() and messagebox.askyesno("导出卡片", "资料还没保存，卡片会用保存前的内容。要现在保存吗？", parent=self):
            self.save()
        path = filedialog.asksaveasfilename(
            parent=self,
            title="导出车型卡片",
            defaultextension=".html",
            initialfile=f"{self.model}-车型卡片.html",
            filetypes=[("网页文件", "*.html"), ("所有文件", "*.*")],
        )
        if not path:
            return
        profile = dict(self.app.profiles.get(self.model, {}))
        try:
            written, size = export_model_card(
                Path(path), self.model, self.app.ranges, self.app.tag_overrides, profile, application_dir()
            )
        except DataError as error:
            messagebox.showerror("导出失败", str(error), parent=self)
            return
        self.app.status_var.set(f"已导出车型卡片 {written.name}（{size / 1024:.0f} KB）")
        if messagebox.askyesno("导出完成", f"已导出：\n{written}\n\n大小 {size / 1024:.0f} KB。\n现在打开看看吗？", parent=self):
            try:
                os.startfile(written)  # noqa: S606 - 交给默认浏览器
            except OSError as error:
                messagebox.showerror("打不开", str(error), parent=self)

    def close_dialog(self) -> None:
        if self.is_dirty():
            answer = messagebox.askyesnocancel("尚未保存", "车型资料已修改，是否现在保存？", parent=self)
            if answer is None:
                return
            if answer:
                self.save()
                if self.is_dirty():
                    return
        self.destroy()


# --------------------------------------------------------------------------- #
# Tag editor for newly imported models
# --------------------------------------------------------------------------- #


class TagEditorWindow(WindowMixin, tk.Toplevel):
    """按车型修正厂商 / 燃料 / 车长，结果写进 车型标签.json。

    留空或填成自动推断值，就等于「不覆盖」，让程序继续按型号代号推断。
    """

    FIELDS = (("manufacturer", "厂商"), ("fuel", "燃料"), ("length", "车长"), ("body", "车体形式"))

    def __init__(self, detail: ModelDetailWindow) -> None:
        super().__init__(detail)
        self.detail = detail
        self.app = detail.app
        self.theme = self.app.theme
        self.model = detail.model
        self.auto = model_tags(self.model)
        current = model_tags(self.model, self.app.tag_overrides)
        self.prefix = manufacturer_key(self.model)
        self.setup_dialog(detail, f"标签设置 - {self.model}", 520, 480)
        pad = self.theme.px

        suggestions = self._suggestions()
        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(14), pady=pad(14))

        # 按钮固定在卡片底部先占位：内容变高也不会把「保存」挤出窗口
        footer = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), 0, pad(16), pad(14)))
        footer.pack(side="bottom", fill="x")
        ttk.Separator(footer, style="Card.TSeparator").pack(fill="x", pady=(0, pad(10)))
        button_row = ttk.Frame(footer, style="Card.TFrame")
        button_row.pack(fill="x")
        ttk.Button(button_row, text="保存", style="Primary.TButton", command=self.save).pack(side="right")
        ttk.Button(button_row, text="取消", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))
        ttk.Button(button_row, text="恢复自动推断", style="Danger.TButton", command=self.reset_to_auto).pack(side="left")

        body = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        ttk.Label(body, text=self.model, style="CardValue.TLabel", font=self.theme.font(12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(body, text=f"自动推断：{self.auto.manufacturer} · {self.auto.fuel} · {self.auto.length}", style="CardMuted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(pad(2), pad(10))
        )

        self.vars: dict[str, tk.StringVar] = {}
        for row, (field, label) in enumerate(self.FIELDS, start=2):
            ttk.Label(body, text=label, style="CardMuted.TLabel").grid(row=row, column=0, sticky="w", pady=pad(5))
            variable = tk.StringVar(value=getattr(current, field))
            self.vars[field] = variable
            combo = ttk.Combobox(body, textvariable=variable, values=suggestions[field], state="normal")
            combo.grid(row=row, column=1, sticky="ew", pady=pad(5), padx=(pad(10), 0))
            combo.bind("<Return>", lambda _event: self.save())

        self.prefix_var = tk.BooleanVar(value=False)
        next_row = 2 + len(self.FIELDS)
        if self.prefix:
            names = self.app.tag_overrides.get("manufacturers") if isinstance(self.app.tag_overrides, dict) else None
            existing = names.get(self.prefix) if isinstance(names, dict) else None
            self.prefix_var.set(bool(existing))
            check = ttk.Checkbutton(
                body,
                text=f"把「{self.prefix}」前缀的所有车型厂商都设为上面的值",
                variable=self.prefix_var,
                style="Card.TCheckbutton",
            )
            check.grid(row=next_row, column=0, columnspan=2, sticky="w", pady=(pad(10), 0))
            next_row += 1

        ttk.Label(body, text="别名 / 俗称", style="CardMuted.TLabel").grid(row=next_row, column=0, sticky="w", pady=pad(5))
        self.alias_var = tk.StringVar(value="、".join(model_aliases(self.model, self.app.tag_overrides)))
        ttk.Entry(body, textvariable=self.alias_var).grid(row=next_row, column=1, sticky="ew", pady=pad(5), padx=(pad(10), 0))
        ttk.Label(body, text="用顿号或逗号分隔，例如：大福后、气尼；填了以后在车型目录里搜俗称也能找到。", style="CardMuted.TLabel", wraplength=pad(400)).grid(
            row=next_row + 1, column=0, columnspan=2, sticky="w", pady=(0, pad(6))
        )

        ttk.Label(
            body,
            text=f"保存到 {TAG_FILENAME}；留空或填回自动推断值等于取消覆盖。",
            style="CardMuted.TLabel",
            wraplength=pad(400),
        ).grid(row=next_row + 2, column=0, columnspan=2, sticky="w", pady=(pad(4), 0))

        # 按内容定高：字段多了（比如加了「车体形式」）也不会把内容挤掉
        self.fit_dialog_height()

    def _suggestions(self) -> dict[str, list[str]]:
        return tag_suggestions(self.app.ranges, self.app.tag_overrides)

    # -- persistence ----------------------------------------------------- #
    def _mutate(self, mutate) -> bool:
        overrides = clean_tag_overrides(self.app.tag_overrides if isinstance(self.app.tag_overrides, dict) else {})
        mutate(overrides)
        cleaned = clean_tag_overrides(overrides)
        try:
            save_tag_overrides(application_dir(), cleaned)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return False
        self.app.set_tag_overrides(cleaned)
        self.detail.refresh_tags()
        return True

    def save(self) -> None:
        values = {field: self.vars[field].get().strip() for field, _label in self.FIELDS}
        use_prefix = self.prefix_var.get() and bool(self.prefix)
        aliases = [name.strip() for name in re.split(r"[、,，;；/]+", self.alias_var.get()) if name.strip()]

        def mutate(overrides: dict) -> None:
            apply_model_tag_values(overrides, self.model, values)
            if use_prefix:
                apply_manufacturer_prefix(overrides, self.prefix, values["manufacturer"])
            table = overrides.get("aliases")
            if not isinstance(table, dict):
                table = {}
            if aliases:
                table[self.model] = aliases
                overrides["aliases"] = table
            else:
                table.pop(self.model, None)
                if not table:
                    overrides.pop("aliases", None)
                else:
                    overrides["aliases"] = table

        if not self._mutate(mutate):
            return
        self.app.status_var.set(f"已更新 {self.model} 的标签")
        self.destroy()

    def reset_to_auto(self) -> None:
        def mutate(overrides: dict) -> None:
            models = overrides.get("models")
            if isinstance(models, dict):
                models.pop(self.model, None)

        if not self._mutate(mutate):
            return
        self.app.status_var.set(f"已恢复 {self.model} 的自动推断标签")
        self.destroy()


# --------------------------------------------------------------------------- #
# Backup browser / restore
# --------------------------------------------------------------------------- #


class BackupWindow(WindowMixin, tk.Toplevel):
    """列出自动备份并支持一键恢复（恢复前也会先备份当前状态）。"""

    COLUMNS = (
        ("source", "数据", 150, "w"),
        ("time", "备份时间", 170, "center"),
        ("size", "大小", 90, "center"),
        ("file", "文件", 300, "w"),
    )

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)
        self.theme = self.app.theme
        self.setup_dialog(parent, "备份与恢复", 820, 560)
        pad = self.theme.px
        self.entries: list[tuple[str, Path, str, float, int]] = []

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text="备份与恢复", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.count_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.count_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")
        ttk.Label(
            frame,
            text=f"每次保存前都会自动备份到 {BACKUP_DIRECTORY_NAME} 文件夹（每份数据保留最近 {BACKUP_KEEP} 次）；恢复前也会先把当前状态备份一遍。",
            style="Muted.TLabel",
            wraplength=pad(760),
        ).grid(row=1, column=0, sticky="w", pady=(pad(4), pad(10)))

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(key, title, width, anchor, numeric=key == "size")
                for key, title, width, anchor in self.COLUMNS
            ),
            empty_text="还没有备份（首次保存后才会产生）",
        )
        self.table.card.grid(row=2, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.tree.bind("<Double-Button-1>", lambda _event: self.restore())

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(actions, text="⟳ 恢复选中", style="Primary.TButton", command=self.restore).pack(side="left")
        ttk.Button(actions, text="刷新", style="Secondary.TButton", command=self.refresh).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="打开备份文件夹", style="Secondary.TButton", command=self.open_folder).pack(side="left")
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")

        self.refresh()

    def refresh(self) -> None:
        self.entries = list_backups(application_dir())
        rows = [
            (label, datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"), f"{size / 1024:.1f} KB", path.name)
            for label, path, _original, mtime, size in self.entries
        ]
        # iid 仍是 self.entries 里的下标：表头排序后「恢复选中」也不会拿错备份
        self.table.set_rows(rows, keep_selection=False, iid=keep_row_ids(rows, [str(index) for index in range(len(rows))]))
        self.count_var.set(f"共 {len(self.entries)} 份备份" if self.entries else "还没有备份（首次保存后才会产生）")

    def open_folder(self) -> None:
        directory = backup_directory(application_dir())
        try:
            directory.mkdir(exist_ok=True)
            import os

            os.startfile(directory)  # noqa: S606 - 打开资源管理器
        except OSError as error:
            messagebox.showerror("打开失败", str(error), parent=self)

    def restore(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("恢复备份", "请先选中一份备份。", parent=self)
            return
        label, path, original, mtime, _size = self.entries[int(selection[0])]
        when = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
        if not messagebox.askyesno(
            "恢复备份",
            f"将用这份备份覆盖当前的「{label}」（{original}）：\n\n备份时间：{when}\n文件：{path.name}\n\n当前状态会先自动备份一份，确定继续吗？",
            parent=self,
        ):
            return
        try:
            restore_backup(application_dir(), path, original)
        except DataError as error:
            messagebox.showerror("恢复失败", str(error), parent=self)
            return
        self.app.reload_data()
        self.refresh()
        self.app.status_var.set(f"已从备份恢复「{label}」（{when}）")
        messagebox.showinfo("恢复完成", f"已恢复「{label}」。\n\n如果界面上的数字没变，关掉重开一次即可。", parent=self)


# --------------------------------------------------------------------------- #
# Data health check / bulk data exchange
# --------------------------------------------------------------------------- #


class LineEditorWindow(WindowMixin, tk.Toplevel):
    """新增/修改一条线路记录。"""

    def __init__(self, parent: tk.Misc, link: LineLink | None = None) -> None:
        super().__init__(parent)
        self.manager = parent
        self.app: BusLookupApp = getattr(parent, "app", parent)  # type: ignore[attr-defined]
        self.theme = self.app.theme
        self.link = link
        self.setup_dialog(parent, "修改线路记录" if link else "新增线路记录", 460, 340, resizable=False)
        pad = self.theme.px

        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(16), pady=pad(16))
        header = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14), pad(16), 0))
        header.pack(fill="x")
        self.dialog_header(
            header,
            "修改线路记录" if link else "新增线路记录",
            "只能说「我记过」：记录能证明记录过，不代表现在还在跑；现在还在跑的请在状态下选「现在还在跑」。",
            inside_card=True,
            wrap=400,
        )
        body = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(12)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        self.line_var = tk.StringVar(value=link.line if link else "")
        self.model_var = tk.StringVar(value=link.model if link else "")
        self.numbers_var = tk.StringVar(value="、".join(link.numbers) if link else "")
        self.first_var = tk.StringVar(value=link.first if link else "")
        self.last_var = tk.StringVar(value=link.last if link else "")
        self.status_var = tk.StringVar(value=LINE_STATUS_LABELS.get(link.status if link else "recorded", "我记过"))

        rows = (
            ("线路", self.line_var, None),
            ("车型", self.model_var, None),
            ("自编号", self.numbers_var, None),
            ("首次记录", self.first_var, "年.月，例如 2024.9"),
            ("末次记录", self.last_var, "年.月"),
        )
        for index, (label, variable, hint) in enumerate(rows):
            ttk.Label(body, text=label).grid(row=index, column=0, sticky="w", pady=pad(4), padx=(0, pad(8)))
            entry = ttk.Entry(body, textvariable=variable, style="Search.TEntry")
            entry.grid(row=index, column=1, sticky="ew", pady=pad(4))
            if hint:
                ttk.Label(body, text=hint, style="CardMuted.TLabel").grid(row=index, column=2, sticky="w", padx=pad(6))
        ttk.Label(body, text="状态").grid(row=len(rows), column=0, sticky="w", pady=pad(4), padx=(0, pad(8)))
        ttk.Combobox(
            body,
            textvariable=self.status_var,
            values=[LINE_STATUS_LABELS[key] for key in LINE_STATUSES],
            state="readonly",
        ).grid(row=len(rows), column=1, sticky="w", pady=pad(4))

        footer = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), 0, pad(16), pad(14)))
        footer.pack(side="bottom", fill="x")
        ttk.Separator(footer, style="Card.TSeparator").pack(fill="x", pady=(0, pad(10)))
        button_row = ttk.Frame(footer, style="Card.TFrame")
        button_row.pack(fill="x")
        ttk.Button(button_row, text="保存", style="Primary.TButton", command=self.save).pack(side="right")
        ttk.Button(button_row, text="取消", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))
        self.fit_dialog_height()

    def save(self) -> None:
        line = self.line_var.get().strip()
        if not line:
            messagebox.showinfo("线路记录", "线路名不能为空。", parent=self)
            return
        status = next((key for key, label in LINE_STATUS_LABELS.items() if label == self.status_var.get()), "recorded")
        numbers = tuple(part.strip() for part in re.split(r"[、,，\s]+", self.numbers_var.get()) if part.strip())
        link = LineLink(
            line=line,
            model=self.model_var.get().strip(),
            numbers=numbers,
            first=self.first_var.get().strip(),
            last=self.last_var.get().strip(),
            status=status,
            source="manual",
            note=self.link.note if self.link else "",
        )
        self.manager.upsert(self.link, link)
        self.close_dialog()


class LineManagerWindow(WindowMixin, tk.Toplevel):
    """线路管理：手工维护「我记过 线路 ↔ 车型/自编号」，也可以从 Excel 导入。"""

    COLUMNS = (
        ("line", "线路", 150, "w", False),
        ("model", "车型", 180, "w", False),
        ("numbers", "自编号", 70, "e", True),
        ("span", "记录时间", 130, "w", False),
        ("status", "状态", 90, "center", False),
        ("source", "来源", 70, "center", False),
    )

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "线路管理", 900, 600)
        pad = self.theme.px
        self.links: list[LineLink] = list(app.lines)
        self._opened_fingerprint = _lines_fingerprint(self.links)  # 用来发现"别处改过数据"
        self.dirty = False

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        self.dialog_header(
            frame,
            "线路管理",
            "这里的数据是你自己的记录（存在 线路.json，默认是空的，发布给别人时不会带出去）。"
            "要给别人用也不需要它：没有线路数据时程序照常工作，只是不显示「我记过」那部分。",
            wrap=820,
        )

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor, numeric=numeric)
                for identifier, title, width, anchor, numeric in self.COLUMNS
            ),
            empty_text="还没有线路记录\n点「新增」手工加一条，或者「从外部表格导入」",
            on_open=self.edit_selected,
        )
        self.table.card.pack(fill="both", expand=True, pady=(pad(10), 0))
        self.tree = self.table.tree

        self.summary_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel", wraplength=pad(820)).pack(fill="x", pady=(pad(6), 0))

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(pad(8), 0))
        ttk.Button(actions, text="新增", style="Primary.TButton", command=self.add_link).pack(side="left")
        ttk.Button(actions, text="修改选中", style="Secondary.TButton", command=self.edit_selected).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="✕ 删除选中", style="Danger.TButton", command=self.delete_selected).pack(side="left")
        ttk.Button(actions, text="从外部表格导入…", style="Secondary.TButton", command=self.import_from_file).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="保存", style="Primary.TButton", command=self.save).pack(side="right")
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))

        self.refresh()
        self.fit_dialog_height()

    def _rows(self) -> list[tuple]:
        return [
            (
                link.line,
                link.model,
                len(link.numbers),
                line_display(link),
                LINE_STATUS_LABELS.get(link.status, link.status),
                "导入" if link.source == "import" else "手工",
            )
            for link in self.links
        ]

    def refresh(self) -> None:
        self.table.set_rows(self._rows(), keep_selection=False, iid=lambda position, row: str(position))
        families = line_family_suggestions(self.links)
        hint = ""
        if families:
            key, names = families[0]
            hint = f"　·　疑似同号家族：{'、'.join(names[:3])}（要不要合并由你定，程序不自动合并）"
        self.summary_var.set(
            f"共 {len(self.links)} 条记录 · {len(all_line_names(self.links))} 条线路"
            f" · {len({link.model for link in self.links if link.model})} 种车型{hint}"
        )

    def upsert(self, old: LineLink | None, new: LineLink) -> None:
        if old is None:
            self.links.append(new)
        else:
            index = self.links.index(old)
            self.links[index] = new
        self.dirty = True
        self.refresh()

    def add_link(self) -> None:
        LineEditorWindow(self, None)

    def edit_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("线路管理", "请先选中一条记录。", parent=self)
            return
        LineEditorWindow(self, self.links[int(selection[0])])

    def delete_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("线路管理", "请先选中要删除的记录。", parent=self)
            return
        for item in sorted((int(value) for value in selection), reverse=True):
            self.links.pop(item)
        self.dirty = True
        self.refresh()

    def import_from_file(self) -> None:
        """从外部表格导入（通用：任何人的表都行，列名认「线路/车型/自编号」）。"""
        path = filedialog.askopenfilename(
            parent=self,
            title="选择要导入的表格",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            records = load_ride_records(Path(path))
        except DataError as error:
            messagebox.showerror("导入失败", str(error), parent=self)
            return
        if not records:
            messagebox.showinfo("导入", "这个文件里没找到「线路 + 车型」这样的列，导入不了。", parent=self)
            return
        incoming = build_links_from_records(records, source="import")
        plan = merge_lines(self.links, incoming)
        if not messagebox.askyesno(
            "从表格导入线路记录",
            f"从 {Path(path).name} 读出 {len(records)} 条记录，整理成 {len(incoming)} 条线路记录：\n"
            f"新增 {len(plan.added)} · 合并更新 {len(plan.updated)}\n\n"
            "同一「线路 + 车型」会合并（时间取并集、自编号取并集）。确定导入吗？",
            parent=self,
        ):
            return
        self.links = apply_line_plan(self.links, plan)
        self.dirty = True
        self.refresh()
        self.app.status_var.set(f"已导入 {len(plan.added)} 条新线路记录（还要点「保存」才写进文件）")

    def save(self) -> None:
        # 这个窗口是拿"打开时那一份"在编辑：如果这期间别处改过线路数据
        # （典型场景：在「线路资料」里点了「写入线路数据」，或「我的记录」里导入了新的），
        # 直接保存会把那些改动整份覆盖掉——所以先问一句。
        live = list(self.app.lines)
        if _lines_fingerprint(live) != self._opened_fingerprint:
            if not messagebox.askyesno(
                "线路数据在别处被改过",
                f"打开这个窗口之后，线路数据被别的操作改过（现在有 {len(live)} 条）。\n\n"
                f"继续保存会用这里的 {len(self.links)} 条**整份覆盖**它，"
                "别处新增的（例如联网写入的）会丢。\n\n"
                "想保留两边：先「取消」，再重新打开这个窗口编辑。\n\n确定要覆盖吗？",
                parent=self,
            ):
                return
        try:
            save_lines(application_dir(), self.links)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.lines = list(self.links)
        self._opened_fingerprint = _lines_fingerprint(self.links)
        self.dirty = False
        self.app.notify_data_changed({"lines"})
        self.app.status_var.set(f"已保存 {len(self.links)} 条线路记录")
        self.close_dialog()


class LineQueryWindow(WindowMixin, tk.Toplevel):
    """线路查询：左边选线路，右边看这条线记录过的车型与自编号。"""

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "线路查询", 900, 580)
        pad = self.theme.px
        self.links = list(app.lines)
        self._catalog_index: dict[str, list[tuple[str, str]]] | None = None

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        self.dialog_header(
            frame,
            "线路查询",
            "数据来自你自己的线路记录 + 联网查过的（分别标出来）：只显示已经记过的线路与车型，时间范围也是记录范围。",
            wrap=820,
        )

        self.filter_var = tk.StringVar()
        search_row = ttk.Frame(frame)
        search_row.pack(fill="x", pady=(pad(10), pad(8)))
        ttk.Label(search_row, text="搜索线路").pack(side="left", padx=(0, pad(6)))
        search = ttk.Entry(search_row, textvariable=self.filter_var, style="Search.TEntry")
        search.pack(side="left", fill="x", expand=True)
        attach_placeholder(search, "线路名，例如 886", self.theme)
        ttk.Button(search_row, text="查询", style="Primary.TButton", command=self.refresh_lines).pack(side="left", padx=pad(6))
        ttk.Button(search_row, text="显示全部", style="Ghost.TButton", command=self.clear_filter).pack(side="left")
        ttk.Label(search_row, text="来源").pack(side="left", padx=(pad(10), pad(4)))
        self.source_var = tk.StringVar(value="全部来源")
        self.source_combo = ttk.Combobox(
            search_row, textvariable=self.source_var, state="readonly", width=12,
            values=("全部来源", "我记过", "网上查的", "表格导入"),
        )
        self.source_combo.pack(side="left")
        self.source_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_lines())
        self.filter_var.trace_add("write", lambda *_args: self.refresh_lines())
        search.bind("<Return>", lambda _event: self.refresh_lines())

        body = ttk.Frame(frame)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        self.line_table = DataTable(
            body,
            self.theme,
            (
                TableColumn("line", "线路", 110, "w"),
                TableColumn("stops", "首末站（本地目录）", 220, "w"),
                TableColumn("models", "车型", 60, "e", numeric=True),
                TableColumn("numbers", "自编号", 70, "e", numeric=True),
            ),
            empty_text="还没有线路记录",
        )
        self.line_table.card.grid(row=0, column=0, sticky="nsew", padx=(0, pad(8)))
        self.line_table.tree.bind("<<TreeviewSelect>>", lambda _event: self.show_line())

        self.detail_table = DataTable(
            body,
            self.theme,
            (
                TableColumn("model", "车型", 170, "w"),
                TableColumn("numbers", "自编号", 240, "w"),
                TableColumn("span", "记录时间", 150, "w"),
                TableColumn("status", "状态", 84, "center"),
            ),
            empty_text="左边选一条线路",
        )
        self.detail_table.card.grid(row=0, column=1, sticky="nsew")
        self.detail_table.tree.bind("<Double-Button-1>", self._open_model)

        footer = ttk.Frame(frame)
        footer.pack(fill="x", pady=(pad(8), 0))
        self.summary_var = tk.StringVar()
        ttk.Label(footer, textvariable=self.summary_var, style="Muted.TLabel").pack(side="left")
        ttk.Button(footer, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")

        self.refresh_lines()

    def on_data_changed(self, kinds: set[str]) -> None:
        """线路数据在别处变了（导入 / 联网写入 / 删除）就重新取一份。

        这个窗口构造时把 app.lines 复制了一份，不刷新的话会一直显示旧记录
        （开着它的时候在「线路资料」里点一次「写入线路数据」，它还是老样子）。
        """
        if kinds & {"all", "lines"}:
            self._catalog_index = None  # 别处可能刚更新过线路目录
            self.links = list(self.app.lines)
            self.refresh_lines()

    def clear_filter(self) -> None:
        self.filter_var.set("")
        self.source_var.set("全部来源")

    def _filtered_links(self) -> list[LineLink]:
        """按来源筛选（我记过 / 网上查的 / 表格导入）。"""
        wanted = SOURCE_LABELS.get(
            {"我记过": "manual", "网上查的": "online", "表格导入": "import"}.get(self.source_var.get(), ""),
        )
        if not wanted:
            return list(self.links)
        return [link for link in self.links if source_label(link.source) == wanted]

    def refresh_lines(self) -> None:
        query = self.filter_var.get().strip().casefold()
        links = self._filtered_links()
        # 先按线路归并一次。以前是「每条线路都重扫一遍全部记录」= O(线路 × 记录)，
        # 1000 条记录时每次刷新要 500ms 左右（搜索框每敲一个字都卡一下）；
        # 归并成一次 O(记录) 之后就无所谓数据多大了。
        index: dict[str, dict[str, object]] = {}
        for link in links:
            key = line_key(link.line)
            bucket = index.get(key)
            if bucket is None:
                bucket = {"models": set(), "numbers": set()}
                index[key] = bucket
            if link.model:
                bucket["models"].add(link.model)  # type: ignore[union-attr]
            bucket["numbers"].update(link.numbers)  # type: ignore[union-attr]
        names = [name for name in all_line_names(links) if not query or query in name.casefold()]
        catalog = self._catalog()
        rows = []
        known = 0
        for name in names:
            bucket = index.get(line_key(name)) or {"models": (), "numbers": ()}
            stops = self._stops_text(catalog.get(line_key(name)))
            if stops:
                known += 1
            rows.append((name, stops, len(bucket["models"]), len(bucket["numbers"])))  # type: ignore[arg-type]
        self.line_table.set_rows(rows, keep_selection=True, iid=lambda position, row: row[0])
        # 选中优先保留；一个都没选中（刚打开、或筛选后原选择不在了）就自动选第一条，
        # 否则右边的明细会一直空着，得手点一下才出内容。
        if not self.line_table.selection_first():
            children = self.line_table.tree.get_children()
            if children:
                self.line_table.tree.selection_set(children[0])
        suffix = "" if self.source_var.get() == "全部来源" else f"（来源：{self.source_var.get()}）"
        if catalog:
            suffix += f"　其中 {known} 条能对上首末站"
        else:
            suffix += "　（还没有本地线路目录：在「线路资料（联网）」点一次「更新线路目录」，这里就能离线显示首末站）"
        self.summary_var.set(f"共 {len(names)} 条线路 · 记录总数 {len(links)}{suffix}")
        self.show_line()

    # -- 本地线路目录（离线看服务端首末站） ------------------------------- #
    def _catalog(self) -> dict[str, list[tuple[str, str]]]:
        if self._catalog_index is None:
            payload = load_route_list(application_dir())
            self._catalog_index = route_list_numbers(payload.get("routes") or {})
        return self._catalog_index

    @staticmethod
    def _stops_text(pairs) -> str:
        """把目录里的首末站写成一行：「A ↔ B」（正好是上下行）或「A → B　C → D」。"""
        if not pairs:
            return "—"
        first_dir = tuple(pairs[0])
        if len(pairs) == 1:
            return f"{first_dir[0]} → {first_dir[1]}"
        second_dir = tuple(pairs[1])
        if first_dir[0] == second_dir[1] and first_dir[1] == second_dir[0]:
            return f"{first_dir[0]} ↔ {first_dir[1]}"
        return "　".join(f"{first} → {last}" for first, last in pairs[:2])

    def show_line(self) -> None:
        line = self.line_table.selection_first()
        if not line:
            self.detail_table.set_rows([], keep_selection=False)
            return
        key = line_key(line)
        links = [link for link in self._filtered_links() if line_key(link.line) == key]
        links.sort(key=lambda link: (LINE_STATUS_ORDER.get(link.status, 9), link.model))
        rows = [
            (
                link.model or "—",
                "、".join(link.numbers) or "—",
                line_display(link, with_source=True),
                LINE_STATUS_LABELS.get(link.status, link.status),
                source_label(link.source),
            )
            for link in links
        ]
        self.detail_table.set_columns(
            (
                TableColumn("model", "车型", 170, "w"),
                TableColumn("numbers", "自编号", 240, "w"),
                TableColumn("span", "记录时间", 150, "w"),
                TableColumn("status", "状态", 84, "center"),
                TableColumn("source", "来源", 90, "center"),
            )
        )
        self.detail_table.set_rows(rows, keep_selection=False)

    def _open_model(self, _event: tk.Event | None = None) -> None:
        selection = self.detail_table.tree.selection()
        if not selection:
            return
        model = str(self.detail_table.tree.item(selection[0], "values")[0])
        if model in set(model_names(self.app.ranges)):
            ModelDetailWindow(self, model)


class SelfCheckWindow(WindowMixin, tk.Toplevel):
    """自检窗口：把"这个程序现在能不能正常用"一条条列出来，失败项直接给处理建议。"""

    def __init__(self, app: "BusLookupApp") -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "自检", 760, 540)
        pad = self.theme.px
        self._queue: "queue.Queue" = queue.Queue()
        self._poll_after: str | None = None
        self._running = False
        self.items: list = []

        frame = ttk.Frame(self, padding=(pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(
            frame, text="自检", style="HeaderTitle.TLabel",
            background=self.theme.background, foreground=self.theme.text,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            frame, text="查号段表、标签、数据文件、写入权限和磁盘；结果可以复制发给别人",
            style="Muted.TLabel", background=self.theme.background,
        ).grid(row=0, column=0, sticky="w", pady=(pad(26), 0))
        self.state_var = tk.StringVar(value="正在检查…")
        ttk.Label(
            frame, textvariable=self.state_var, style="Muted.TLabel", background=self.theme.background
        ).grid(row=1, column=0, sticky="w", pady=(pad(6), pad(8)))

        holder = ttk.Frame(frame)
        holder.grid(row=2, column=0, sticky="nsew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.table = ttk.Treeview(
            holder, columns=("item", "result", "detail"), show="headings", height=9, selectmode="browse"
        )
        for column, title, width, anchor in (
            ("item", "项目", pad(130), "w"),
            ("result", "结果", pad(56), "center"),
            ("detail", "说明", pad(400), "w"),
        ):
            self.table.heading(column, text=title)
            self.table.column(column, width=width, anchor=anchor, stretch=column != "item")
        scroll = ttk.Scrollbar(holder, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.table.tag_configure("ok", foreground=self.theme.text)
        self.table.tag_configure("warn", foreground=getattr(self.theme, "warning", "#B26A00"))
        self.table.tag_configure("fail", foreground=self.theme.danger)

        self.hint_var = tk.StringVar(value="")
        ttk.Label(
            frame, textvariable=self.hint_var, style="Muted.TLabel", background=self.theme.background,
            wraplength=pad(660), justify="left",
        ).grid(row=3, column=0, sticky="w", pady=(pad(8), 0))

        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(buttons, text="重新检查", style="Secondary.TButton", command=self.run_checks).pack(side="left")
        ttk.Button(
            buttons, text="顺便测联网", style="Secondary.TButton", command=lambda: self.run_checks(deep=True)
        ).pack(side="left", padx=(pad(6), 0))
        ttk.Button(buttons, text="复制结果", style="Secondary.TButton", command=self.copy_result).pack(
            side="left", padx=(pad(6), 0)
        )
        ttk.Button(buttons, text="关闭", style="Ghost.TButton", command=self.close_dialog).pack(side="right")

        self.run_checks()

    def run_checks(self, deep: bool = False) -> None:
        """在工作线程里跑，结果回主线程渲染（不在线程里碰控件）。"""
        if self._running:
            return
        self._running = True
        self.state_var.set("正在测联网（要几秒）…" if deep else "正在检查…")
        self.hint_var.set("")
        self.table.delete(*self.table.get_children())
        base = application_dir()

        def work() -> None:
            import bpt_selfcheck as check

            try:
                items = check.run_self_checks(base, {}, deep=deep, tag_checker=_selfcheck_tag_item)
            except Exception as error:  # noqa: BLE001 - 自检自己出错也要显示出来
                items = [
                    check.CheckItem("self", "自检本身", False, f"{type(error).__name__}: {error}", "把这句发给维护者")
                ]
            self._queue.put(items)

        threading.Thread(target=work, name="bpt-selfcheck", daemon=True).start()
        self._poll_after = self.after(120, self._poll)

    def _poll(self) -> None:
        try:
            items = self._queue.get_nowait()
        except queue.Empty:
            self._poll_after = self.after(120, self._poll)
            return
        self._running = False
        self.items = list(items)
        self.render(self.items)

    def render(self, items: list) -> None:
        for item in items:
            tag = "ok" if item.ok else ("warn" if not item.blocking else "fail")
            self.table.insert("", "end", values=(item.title, item.mark, item.detail), tags=(tag,))
        failed = [item for item in items if not item.ok and item.blocking]
        notes = [item for item in items if not item.ok and not item.blocking]
        if failed:
            self.state_var.set(f"有 {len(failed)} 项需要处理")
        elif notes:
            self.state_var.set(f"能正常用；另有 {len(notes)} 项是提示")
        else:
            self.state_var.set("全部通过")
        self.hint_var.set("　".join(f"{item.title}：{item.hint}" for item in items if not item.ok and item.hint))

    def copy_result(self) -> None:
        import bpt_selfcheck as check

        text = "\n".join(check.format_report(self.items)) if self.items else "还没有结果"
        self.clipboard_clear()
        self.clipboard_append(text)
        self.state_var.set("结果已复制到剪贴板")


class DataHealthWindow(WindowMixin, tk.Toplevel):
    """数据体检：待办清单 + 与 Excel 交换标签列。"""

    # 哪些待办类别可以在界面里直接「确认」（把程序当前的判断明确记下来，之后不再提示）
    CONFIRMABLE = {
        "body_guess": ("body", "确认这几台（保留现在的判断）"),
        "manufacturer": ("manufacturer", "确认这几台的厂商名"),
    }

    # 每类待办对应的「怎么解决」按钮（点开就能去改，而不是只告诉你问题）
    ACTIONS: dict[str, tuple[str, str]] = {
        "fuel": ("打开燃料待确认", "open_fuel_confirm"),
        "fuel_guess": ("打开燃料待确认", "open_fuel_confirm"),
        "fuel_manual": ("打开燃料待确认", "open_fuel_confirm"),
        "gap": ("管理号段（补空档）", "open_range_manager"),
        "image": ("打开车型目录", "open_catalog"),
        "profile": ("半自动建档", "open_drafts"),
        "line_model": ("线路管理", "open_line_manager"),
        "line_number": ("线路管理", "open_line_manager"),
        "line_stops": ("线路资料（联网）", "open_line_info"),
        "curated": ("线路管理", "open_line_manager"),
        "metro_source": ("地铁数据工具（写回 / 重新导入）", "open_metro"),
        "metro_gap": ("管理地铁号段与车型", "open_metro_manager"),
        "metro_overlap": ("管理地铁号段与车型", "open_metro_manager"),
        "metro_model": ("管理地铁号段与车型", "open_metro_manager"),
    }

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "数据体检", 820, 640)
        pad = self.theme.px
        self.items: list[HealthItem] = []

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text="数据体检", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.summary_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")

        self.category_table = DataTable(
            frame,
            self.theme,
            (
                TableColumn("title", "待办类别", 560, "w"),
                TableColumn("count", "数量", 90, "e", numeric=True),
            ),
            empty_text="没有待办",
            on_open=self._show_selected,
        )
        self.category_table.card.grid(row=1, column=0, sticky="ew", pady=(pad(8), pad(10)))
        self.category_tree = self.category_table.tree
        self.category_tree.configure(height=4, selectmode="browse")  # 矮一点：给下面的明细和按钮留出高度
        self.category_tree.bind("<<TreeviewSelect>>", self._show_selected)

        self.detail_table = DataTable(
            frame,
            self.theme,
            (
                TableColumn("subject", "对象", 240, "w"),
                TableColumn("detail", "说明", 480, "w"),
            ),
            selectmode="extended",
            empty_text="这一类目前没有待办项",
            on_open=self._open_subject,
        )
        self.detail_table.card.grid(row=2, column=0, sticky="nsew")
        self.detail_tree = self.detail_table.tree
        self.detail_tree.configure(height=6)
        self.detail_tree.bind("<Return>", self._open_subject)

        self.hint_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.hint_var, style="Muted.TLabel", wraplength=pad(760)).grid(
            row=3, column=0, sticky="w", pady=(pad(8), 0)
        )

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, sticky="ew", pady=(pad(8), 0))
        # 主操作行：只放「跟当前待办有关」的按钮，其余收进「更多工具」
        primary = ttk.Frame(actions)
        primary.pack(fill="x")
        self.action_button = ttk.Button(
            primary, text="怎么解决", style="Primary.TButton", command=self._run_action
        )
        self.confirm_button = ttk.Button(
            primary, text="确认选中项", style="Primary.TButton", command=self.confirm_selected
        )
        self.tools_button = ttk.Button(primary, text="更多工具 ▾", style="Secondary.TButton", command=self.toggle_tools)
        self.bulk_button = ttk.Button(primary, text="批量设置标签", style="Secondary.TButton", command=self.bulk_tag_selected)
        ttk.Button(primary, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")
        self.tools_button.pack(side="right", padx=pad(8))
        self.bulk_button.pack(side="right", padx=pad(8))
        # 两个按当前待办出现的按钮先不显示，选中待办时再 pack 出来
        self.action_button.pack_forget()
        self.confirm_button.pack_forget()

        # 「更多工具」：平时收起，点开才显示（按钮太多反而不知道点哪个）
        self.tools_frame = ttk.Frame(actions)
        self.tools_shown = False
        tools_row1 = ttk.Frame(self.tools_frame)
        tools_row1.pack(anchor="w")
        ttk.Label(tools_row1, text="标签：", style="Muted.TLabel").pack(side="left", padx=(0, pad(4)))
        for text, command in (
            ("燃料待确认", lambda: TagConfirmWindow(self.app)),
            ("标签规则", self.open_rules),
            ("撤销上次标签修改", self.undo_last_tag_change),
            ("备份与恢复", lambda: BackupWindow(self)),
        ):
            ttk.Button(tools_row1, text=text, style="Ghost.TButton", command=command).pack(side="left", padx=pad(4))
        tools_row2 = ttk.Frame(self.tools_frame)
        tools_row2.pack(anchor="w", pady=(pad(4), 0))
        ttk.Label(tools_row2, text="导入导出：", style="Muted.TLabel").pack(side="left", padx=(0, pad(4)))
        for text, command in (
            ("增量导入 Excel", self.import_ranges),
            ("半自动建档", self.open_drafts),
            ("统计面板", self.open_stats),
            ("⇩ 导出待办 CSV", self.export_todos),
            ("⇩ 导出网页版", self.export_web),
            ("⇩ 导出带标签 Excel", self.export_workbook),
            ("从 Excel 读回标签", self.import_workbook),
        ):
            ttk.Button(tools_row2, text=text, style="Ghost.TButton", command=command).pack(side="left", padx=pad(4))

        self.refresh()
        self.fit_dialog_height()

    def toggle_tools(self) -> None:
        """展开 / 收起「更多工具」。"""
        if self.tools_shown:
            self.tools_frame.pack_forget()
            self.tools_button.configure(text="更多工具 ▾")
        else:
            self.tools_frame.pack(fill="x", pady=(self.theme.px(8), 0))
            self.tools_button.configure(text="收起工具 ▴")
        self.tools_shown = not self.tools_shown

    def on_data_changed(self, kinds: set[str]) -> None:
        self.refresh()

    def _run_action(self) -> None:
        """点「怎么解决」：按当前类别去打开对应的工具窗口。"""
        category = self.category_tree.selection()
        spec = self.ACTIONS.get(category[0] if category else "")
        if spec is None:
            return
        method = getattr(self, spec[1], None)
        if method is None:
            return
        method()

    # -- 各类待办的解决入口 ----------------------------------------------- #
    def open_fuel_confirm(self) -> None:
        TagConfirmWindow(self.app)

    def open_range_manager(self) -> None:
        RangeManagerWindow(self.app)

    def open_catalog(self) -> None:
        CatalogWindow(self.app)

    def open_drafts(self) -> None:
        DraftProfilesWindow(self.app)

    def open_line_manager(self) -> None:
        LineManagerWindow(self.app)

    def open_metro(self) -> None:
        if self.app.metro_ranges:
            MetroWindow(self.app)
        else:
            messagebox.showerror("地铁编号查询", f"没有可用的地铁数据，请检查 {METRO_SOURCE_FILENAME}。", parent=self)

    def open_metro_manager(self) -> None:
        if not self.app.metro_ranges:
            messagebox.showerror("地铁号段", f"没有可用的地铁数据，请检查 {METRO_SOURCE_FILENAME}。", parent=self)
            return
        MetroRangeManagerWindow(MetroWindow(self.app))

    # -- 批量处理 --------------------------------------------------------- #
    def selected_models(self) -> list[str]:
        """明细里选中的、且确实是车型的那些行。"""
        known = set(model_names(self.app.ranges))
        models: list[str] = []
        for item in self.detail_tree.selection():
            subject = str(self.detail_tree.item(item, "values")[0])
            if subject in known and subject not in models:
                models.append(subject)
        return models

    def bulk_tag_selected(self) -> None:
        models = self.selected_models()
        if not models:
            messagebox.showinfo("批量设置标签", "请先在下面的明细里选中一个或多个车型（Ctrl / Shift 可多选）。", parent=self)
            return
        BulkTagWindow(self, models)

    def undo_last_tag_change(self) -> None:
        backups = [entry for entry in list_backups(application_dir()) if entry[0].startswith("车型标签")]
        if not backups:
            messagebox.showinfo("撤销", "还没有「车型标签」的备份可以撤销。", parent=self)
            return
        _label, path, original, mtime, _size = backups[0]
        when = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
        if not messagebox.askyesno(
            "撤销上次标签修改",
            f"将把 车型标签.json 恢复到 {when} 的版本。\n（当前状态也会先备份一份）\n\n确定继续吗？",
            parent=self,
        ):
            return
        try:
            restore_backup(application_dir(), path, original)
        except DataError as error:
            messagebox.showerror("撤销失败", str(error), parent=self)
            return
        self.app.reload_data()
        self.app.status_var.set(f"已撤销到 {when} 的标签版本")

    def confirm_selected(self) -> None:
        """把选中的这几台「确认」下来：当前判断是什么就明确记成什么，之后体检不再提示。"""
        category = self.category_tree.selection()
        spec = self.CONFIRMABLE.get(category[0] if category else "")
        if spec is None:
            messagebox.showinfo("确认", "当前这一类的待办不需要逐条确认。", parent=self)
            return
        field, _label = spec
        models = self.selected_models()
        if not models:
            messagebox.showinfo("确认", "请先在下面的明细里选中要确认的车型（Ctrl / Shift 可多选）。", parent=self)
            return
        overrides = clean_tag_overrides(self.app.tag_overrides if isinstance(self.app.tag_overrides, dict) else {})
        confirmed = 0
        for model in models:
            # 用 force=True：即使跟自动推断一样也记下来，这样"猜对了"也能被确认掉
            value = getattr(model_tags(model, overrides), field)
            if value:
                apply_model_tag_values(overrides, model, {field: value}, force=True)
                confirmed += 1
        try:
            save_tag_overrides(application_dir(), overrides)
        except ProfileError as error:
            messagebox.showerror("确认失败", str(error), parent=self)
            return
        self.app.set_tag_overrides(overrides)
        self.refresh()
        self.app.status_var.set(f"已确认 {confirmed} 个车型的{field}（可撤销）")

    # -- rendering ------------------------------------------------------- #
    def refresh(self) -> None:
        self.items = health_report(
            application_dir(),
            self.app.ranges,
            self.app.profiles,
            self.app.tag_overrides,
            self.app.metro_ranges,
            self.app.lines,
            self.app.line_info,
        )
        # iid 用待办类别的 key，选中哪一类就靠它
        self.category_table.set_rows(
            [(item.title, len(item.rows)) for item in self.items],
            keep_selection=False,
            iid=lambda position, row: self.items[position].key,
        )
        pending = sum(len(item.rows) for item in self.items)
        self.summary_var.set(f"共 {len(self.items)} 类待办 · {pending} 条")
        first = next((item.key for item in self.items if item.rows), None)
        if first:
            self.category_tree.selection_set(first)
            self.category_tree.focus(first)
            self._show_selected()
        self.category_empty = self.category_table.empty
        self.category_empty.update("数据很干净，暂时没有待办项" if not pending else None, has_rows=bool(pending))

    def _show_selected(self, _event: tk.Event | None = None) -> None:
        selection = self.category_tree.selection()
        if not selection:
            return
        item = next((entry for entry in self.items if entry.key == selection[0]), None)
        if item is None:
            self.detail_table.set_rows([], keep_selection=False)
            return
        self.hint_var.set(item.hint)
        # 两个按钮只在当前待办用得上时才出现（用户反馈：按钮太多不知道点哪个）
        gap = self.theme.px(8)
        self.action_button.pack_forget()
        self.confirm_button.pack_forget()
        spec = self.CONFIRMABLE.get(item.key)
        if spec is not None:
            self.confirm_button.configure(text=spec[1])
            self.confirm_button.pack(side="left", padx=(gap, 0))
        action = self.ACTIONS.get(item.key)
        if action is not None:
            self.action_button.configure(text=action[0])
            self.action_button.pack(
                side="left", padx=(gap, 0), before=self.confirm_button if spec is not None else self.bulk_button
            )
        if self.tools_shown and not self.tools_frame.winfo_ismapped():
            self.tools_frame.pack(fill="x", pady=(self.theme.px(8), 0))
        self.detail_table.set_rows([tuple(row) for row in item.rows], keep_selection=False)
        self.detail_empty = self.detail_table.empty
        self.detail_empty.update("这一类目前没有待办项", has_rows=bool(item.rows))

    def _open_subject(self, _event: tk.Event | None = None) -> None:
        selection = self.detail_tree.selection()
        if not selection:
            return
        category = self.category_tree.selection()
        key = category[0] if category else ""
        if key == "gap":
            # 「号段空档」的对象是 20000-20212 这样的区间：直接预填新增号段
            for item in selection:
                subject = str(self.detail_tree.item(item, "values")[0])
                match = re.fullmatch(r"(\d{5})-(\d{5})", subject.strip())
                if match:
                    RangeEditorWindow(self, None, (int(match.group(1)), int(match.group(2)), ""))
            return
        if key in ("metro_gap", "metro_overlap", "metro_model"):
            # 地铁号段带字母前缀（T401 这种），不好自动算区间；直接开到管理窗口自己加
            self.open_metro_manager()
            return
        if key == "metro_source":
            self.open_metro()
            return
        subject = str(self.detail_tree.item(selection[0], "values")[0])
        if subject in set(model_names(self.app.ranges)):
            ModelDetailWindow(self, subject)

    # -- exchange -------------------------------------------------------- #
    def open_rules(self) -> None:
        TagRulesWindow(self)

    def import_ranges(self) -> None:
        """把外部 Excel 的号段增量合并进来。"""
        path = filedialog.askopenfilename(
            parent=self,
            title="选择要导入的号段表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            incoming = load_ranges(Path(path))
        except DataError as error:
            messagebox.showerror("无法读取", str(error), parent=self)
            return
        if not incoming:
            messagebox.showinfo("增量导入", "这个文件里没有号段数据。", parent=self)
            return
        if not merge_ranges(self.app.ranges, incoming).total:
            messagebox.showinfo("增量导入", "导入表与当前数据完全一致，没有需要合并的内容。", parent=self)
            return
        MergeWindow(self, incoming, Path(path).name)

    def open_drafts(self) -> None:
        DraftProfilesWindow(self)

    def open_stats(self) -> None:
        StatsWindow(self)

    def export_web(self) -> None:
        """导出单文件网页版，方便发给别人或手机上打开。"""
        default = f"北京公交车型查询器-{datetime.now().strftime('%Y%m%d')}.html"
        path = filedialog.asksaveasfilename(
            parent=self,
            title="导出网页版",
            defaultextension=".html",
            initialfile=default,
            filetypes=[("网页文件", "*.html"), ("所有文件", "*.*")],
        )
        if not path:
            return
        notices = [f"{item.title}：{len(item.rows)} 条" for item in self.items if item.rows]
        try:
            written, size = export_web_page(
                Path(path),
                self.app.ranges,
                self.app.tag_overrides,
                self.app.profiles,
                self.app.metro_ranges,
                application_dir(),
                embed_images=True,
                notices=notices,
                spots=load_spots(application_dir()),
            )
        except DataError as error:
            messagebox.showerror("导出失败", str(error), parent=self)
            return
        self.app.status_var.set(f"已导出网页版 {written.name}（{size / 1024:.0f} KB）")
        messagebox.showinfo(
            "导出完成",
            f"已导出：\n{written}\n\n大小 {size / 1024:.0f} KB。\n"
            "这个文件是自包含的（数据、样式、脚本都在里面），双击即可打开，\n"
            "也可以直接发到手机上离线查看。",
            parent=self,
        )

    def export_todos(self) -> None:
        rows = [(item.title, subject, detail) for item in self.items for subject, detail in item.rows]
        if not rows:
            messagebox.showinfo("导出待办", "当前没有待办项。", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self,
            title="导出待办清单",
            defaultextension=".csv",
            initialfile="数据体检待办.csv",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(["类别", "对象", "说明"])
                writer.writerows(rows)
        except OSError as error:
            messagebox.showerror("导出失败", f"无法写入文件：{error}", parent=self)
            return
        self.app.status_var.set(f"已导出 {len(rows)} 条待办到 {path}")

    def export_workbook(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self,
            title="导出带标签的 Excel",
            defaultextension=".xlsx",
            initialfile="BPT_BusDigitalNumber-带标签.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            count = write_tagged_workbook(Path(path), self.app.ranges, self.app.tag_overrides)
        except DataError as error:
            messagebox.showerror("导出失败", str(error), parent=self)
            return
        messagebox.showinfo(
            "导出完成",
            f"已导出 {count} 行，含「厂商 / 燃料 / 车长」三列。\n\n在 Excel 里改完这三列后，用「从 Excel 读回标签」导回程序即可。",
            parent=self,
        )

    def import_workbook(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="选择含有标签列的 Excel",
            initialdir=str(application_dir()),
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        columns = load_tag_columns(Path(path))
        if not columns:
            messagebox.showinfo("没有标签列", "这个文件里没有找到「车型」和「厂商 / 燃料 / 车长」列。", parent=self)
            return
        merged, ignored = merge_tag_columns(self.app.tag_overrides, columns, set(model_names(self.app.ranges)))
        applied = len(columns) - len(ignored)
        if not applied:
            messagebox.showinfo("没有可用的标签", "文件里的车型都不在当前号段数据里。", parent=self)
            return
        detail = f"将把 {applied} 个车型的标签写进 {TAG_FILENAME}。"
        if ignored:
            detail += f"\n另有 {len(ignored)} 个车型不在当前数据里，会被忽略。"
        if not messagebox.askyesno("读回标签", detail, parent=self):
            return
        try:
            save_tag_overrides(application_dir(), merged)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.set_tag_overrides(merged)
        self.refresh()
        messagebox.showinfo("完成", f"已更新 {applied} 个车型的标签。", parent=self)


# --------------------------------------------------------------------------- #
# 设置窗口
# --------------------------------------------------------------------------- #


class SettingsWindow(WindowMixin, tk.Toplevel):
    """拍车模式开关与界面偏好。"""

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "设置", 600, 480, resizable=False)
        pad = self.theme.px

        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(16), pady=pad(16))
        body = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)

        header = ttk.Frame(body, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        self.dialog_header(header, "拍车模式", "复制自编号就自动查、念出车型、一键叫出输入框——在外面看车时用。", inside_card=True)

        self.clipboard_var = tk.BooleanVar(value=app.preference("clipboard_watch"))
        ttk.Checkbutton(
            body,
            text="复制自编号后自动查询（监听剪贴板）",
            variable=self.clipboard_var,
            style="Card.TCheckbutton",
            command=lambda: self.toggle("clipboard_watch", self.clipboard_var.get()),
        ).grid(row=1, column=0, sticky="w", pady=pad(4))
        ttk.Label(body, text="打开后每 0.4 秒看一次剪贴板；发现新的五位自编号就自动查询并把窗口提到最前。", style="CardMuted.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, pad(10))
        )

        self.speak_var = tk.BooleanVar(value=app.preference("speak_on_query"))
        ttk.Checkbutton(
            body,
            text="查询后朗读车型",
            variable=self.speak_var,
            style="Card.TCheckbutton",
            command=lambda: self.toggle("speak_on_query", self.speak_var.get()),
        ).grid(row=3, column=0, sticky="w", pady=pad(4))
        ttk.Label(body, text="用 Windows 自带语音念出车型代号（首次会稍慢，失败时静默跳过）。", style="CardMuted.TLabel").grid(
            row=4, column=0, sticky="w", pady=(0, pad(10))
        )

        self.hotkey_var = tk.BooleanVar(value=app.preference("hotkey_enabled"))
        ttk.Checkbutton(
            body,
            text="启用全局热键 Ctrl+Alt+B（任何时候按一下就把输入框叫到前面）",
            variable=self.hotkey_var,
            style="Card.TCheckbutton",
            command=lambda: self.toggle("hotkey_enabled", self.hotkey_var.get()),
        ).grid(row=5, column=0, sticky="w", pady=pad(4))
        self.hotkey_status_var = tk.StringVar(value=app.hotkey_status())
        ttk.Label(body, textvariable=self.hotkey_status_var, style="CardMuted.TLabel").grid(
            row=6, column=0, sticky="w", pady=(0, pad(12))
        )

        ttk.Separator(body).grid(row=7, column=0, sticky="ew", pady=pad(6))
        ttk.Label(body, text="界面", style="Section.TLabel").grid(row=8, column=0, sticky="w", pady=(0, pad(6)))
        ttk.Button(body, text="重置窗口位置（下次打开窗口时重新居中）", style="Secondary.TButton", command=self.reset_geometry).grid(
            row=9, column=0, sticky="w"
        )
        ttk.Button(body, text="打开程序文件夹", style="Secondary.TButton", command=self.open_folder).grid(row=10, column=0, sticky="w", pady=(pad(8), 0))

        # 联网（可选，V5）：开关 + 参数 + 缓存/日志维护
        current = online_settings()
        ttk.Separator(body).grid(row=11, column=0, sticky="ew", pady=pad(8))
        ttk.Label(body, text="联网（可选）", style="Section.TLabel").grid(row=12, column=0, sticky="w", pady=(0, pad(6)))
        self.network_var = tk.BooleanVar(value=current.enabled)
        ttk.Checkbutton(
            body, text="允许联网查询（关掉后所有联网按钮变灰，本地功能不受影响）",
            variable=self.network_var, style="Card.TCheckbutton",
        ).grid(row=13, column=0, sticky="w")
        self.auto_lookup_var = tk.BooleanVar(value=self.app.preference("auto_online_lookup", True))
        ttk.Checkbutton(
            body, text="主界面自动显示这辆车「现在在哪 / 最新在哪条线运营」（输入 5 位自编后联网）",
            variable=self.auto_lookup_var, style="Card.TCheckbutton",
            command=lambda: self._toggle_auto_lookup(),
        ).grid(row=14, column=0, sticky="w", pady=(pad(2), 0))
        ttk.Label(
            body,
            text=f"服务：{current.base_url}（第三方数据，仅供参考）",
            style="CardMuted.TLabel", wraplength=pad(520), justify="left",
        ).grid(row=14, column=0, sticky="w", pady=(pad(4), pad(6)))

        numbers = ttk.Frame(body, style="Card.TFrame")
        numbers.grid(row=15, column=0, sticky="w")
        self.timeout_var = tk.StringVar(value=str(current.timeout_sec))
        self.cache_days_var = tk.StringVar(value=str(current.cache_days))
        self.interval_var = tk.StringVar(value=str(current.min_interval_sec))
        for index, (label, variable, width) in enumerate(
            (("单次超时(秒)", self.timeout_var, 5), ("缓存天数", self.cache_days_var, 5), ("最小间隔(秒)", self.interval_var, 5))
        ):
            ttk.Label(numbers, text=label, style="CardMuted.TLabel").grid(row=0, column=index * 2, sticky="w", padx=(0 if index == 0 else pad(10), pad(4)))
            ttk.Entry(numbers, textvariable=variable, width=width).grid(row=0, column=index * 2 + 1, sticky="w")

        ttk.Button(body, text="保存联网设置", style="Secondary.TButton", command=self.save_network).grid(
            row=16, column=0, sticky="w", pady=(pad(6), 0)
        )
        self.network_status_var = tk.StringVar(value="")
        ttk.Label(body, textvariable=self.network_status_var, style="CardMuted.TLabel", wraplength=pad(520)).grid(
            row=17, column=0, sticky="w", pady=(pad(4), 0)
        )

        maintenance = ttk.Frame(body, style="Card.TFrame")
        maintenance.grid(row=18, column=0, sticky="w", pady=(pad(8), 0))
        ttk.Button(maintenance, text="清空联网缓存", style="Ghost.TButton", command=self.clear_cache_action).pack(side="left")
        ttk.Button(maintenance, text="清空联网日志", style="Ghost.TButton", command=self.clear_log_action).pack(side="left", padx=pad(6))
        ttk.Button(maintenance, text="看最近的联网错误", style="Ghost.TButton", command=self.show_errors_action).pack(side="left")

        actions = ttk.Frame(body, style="Card.TFrame")
        actions.grid(row=19, column=0, sticky="e", pady=(pad(12), 0))
        ttk.Button(actions, text="关闭", style="Primary.TButton", command=self.close_dialog).pack(side="right")

    # -- 联网设置 --------------------------------------------------------- #
    def _toggle_auto_lookup(self) -> None:
        """主界面自动联网开关：既存偏好，也立刻作用到正在跑的主窗口。"""
        value = self.auto_lookup_var.get()
        self.app.set_preference("auto_online_lookup", value)
        self.app._auto_online = value
        if not value:
            self.app.online_var.set("")
        self.network_status_var.set("主界面自动联网查询：" + ("开" if value else "关"))

    def save_network(self) -> None:
        """保存联网设置，并让正在用的会话立刻按新设置走（不用重启）。"""
        settings = dict(self.app.settings)
        payload = dict(settings.get("network") or {})
        payload["enabled"] = bool(self.network_var.get())
        for key, variable in (
            ("timeout_sec", self.timeout_var),
            ("cache_days", self.cache_days_var),
            ("min_interval_sec", self.interval_var),
        ):
            try:
                payload[key] = float(variable.get().strip())
            except (TypeError, ValueError):
                payload.pop(key, None)  # 填了看不懂的值就回到默认，不报错卡住
        settings["network"] = payload
        self.app.settings = settings
        try:
            save_settings(application_dir(), settings)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        fresh = OnlineSettings.from_dict(payload)
        shared_service(fresh, application_dir())  # 立刻生效
        self.network_status_var.set(
            f"已保存：联网{'开' if fresh.enabled else '关'}　超时 {fresh.timeout_sec}s　"
            f"缓存 {fresh.cache_days} 天　最小间隔 {fresh.min_interval_sec}s"
        )
        self.app.status_var.set("联网设置已保存")

    def clear_cache_action(self) -> None:
        service = shared_service(online_settings(), application_dir())
        count = service.clear_cache()
        self.network_status_var.set(f"已清空 {count} 条联网缓存" if count else "联网缓存本来就是空的")

    def clear_log_action(self) -> None:
        ok = clear_online_log(application_dir())
        self.network_status_var.set("联网日志已清空" if ok else "日志删不掉（可能被别的程序占着）")

    def show_errors_action(self) -> None:
        errors = recent_online_errors(application_dir())
        if not errors:
            self.network_status_var.set("最近没有联网错误")
            return
        messagebox.showinfo("最近的联网错误", "\n".join(errors[-8:]), parent=self)

    def toggle(self, key: str, value: bool) -> None:
        self.app.set_preference(key, value)
        self.hotkey_status_var.set(self.app.hotkey_status())

    def reset_geometry(self) -> None:
        settings = self.app.settings
        settings.pop("windows", None)
        self.app.settings = settings
        save_settings(application_dir(), settings)
        messagebox.showinfo("已重置", "窗口位置已清除，下次打开各个窗口时会重新居中。", parent=self)

    def open_folder(self) -> None:
        try:
            os.startfile(application_dir())  # noqa: S606 - 打开资源管理器
        except OSError as error:
            messagebox.showerror("打开失败", str(error), parent=self)


class SpotDialog(WindowMixin, tk.Toplevel):
    """记一笔 / 编辑一条拍车记录（可以带上自己拍的照片）。"""

    def __init__(
        self,
        parent: tk.Misc,
        *,
        index: int | None = None,
        number: str = "",
        model: str = "",
        note: str = "",
        photo: str = "",
        cover: str = "",
        time_text: str = "",
    ) -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)
        self.theme = self.app.theme
        self.log_window = parent
        self.index = index
        self.number = number
        self.model = model
        self.cover = cover
        self.time_text = time_text
        self.photo_name = photo
        self.pending_photo: Path | Image.Image | None = None
        self.drop_photo = ""
        self.preview_photo: ImageTk.PhotoImage | None = None
        self.setup_dialog(parent, "记一笔" if index is None else "编辑拍车记录", 520, 520, resizable=False)
        pad = self.theme.px

        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(16), pady=pad(16))
        body = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)

        ttk.Label(body, text=f"{number}　{model}", style="CardStrong.TLabel").grid(row=0, column=0, sticky="w")
        if time_text:
            ttk.Label(body, text=f"拍摄时间：{time_text}", style="CardMuted.TLabel").grid(row=1, column=0, sticky="w", pady=(pad(2), 0))

        ttk.Label(body, text="备注", style="CardMuted.TLabel").grid(row=2, column=0, sticky="w", pady=(pad(12), pad(4)))
        self.note_var = tk.StringVar(value=note)
        note_entry = ttk.Entry(body, textvariable=self.note_var)
        note_entry.grid(row=3, column=0, sticky="ew")
        note_entry.bind("<Return>", lambda _event: self.save())

        ttk.Label(body, text="照片", style="CardMuted.TLabel").grid(row=4, column=0, sticky="w", pady=(pad(12), pad(4)))
        self.holder = tk.Frame(body, background=self.theme.surface_alt, height=pad(220))
        self.holder.grid(row=5, column=0, sticky="ew")
        self.holder.grid_propagate(False)
        self.preview = tk.Label(self.holder, background=self.theme.surface_alt, foreground=self.theme.muted, font=self.theme.font(10), text="还没有照片")
        self.preview.pack(fill="both", expand=True)

        photo_actions = ttk.Frame(body, style="Card.TFrame")
        photo_actions.grid(row=6, column=0, sticky="w", pady=(pad(8), 0))
        ttk.Button(photo_actions, text="选照片", style="Secondary.TButton", command=self.choose_photo).pack(side="left", padx=(0, pad(6)))
        ttk.Button(photo_actions, text="粘贴照片 (Ctrl+V)", style="Secondary.TButton", command=self.paste_photo).pack(side="left", padx=(0, pad(6)))
        ttk.Button(photo_actions, text="移除照片", style="Danger.TButton", command=self.remove_photo).pack(side="left")

        actions = ttk.Frame(body, style="Card.TFrame")
        actions.grid(row=7, column=0, sticky="ew", pady=(pad(16), 0))
        ttk.Button(actions, text="保存", style="Primary.TButton", command=self.save).pack(side="right")
        ttk.Button(actions, text="取消", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))

        self.bind("<Control-v>", self._on_paste)
        self._show_photo()
        note_entry.focus_set()

    # -- 照片 ------------------------------------------------------------- #
    def _current_photo(self) -> str:
        return self.photo_name

    def _show_photo(self) -> None:
        if isinstance(self.pending_photo, Image.Image):
            self._show_preview(self.pending_photo)
            return
        if isinstance(self.pending_photo, Path):
            try:
                self._show_preview(load_image(self.pending_photo))
                return
            except ProfileError:
                pass
        name = self._current_photo()
        if name:
            path = stored_photo_path(application_dir(), name)
            if path.is_file():
                try:
                    with Image.open(path) as image:
                        self._show_preview(image.convert("RGB"))
                    return
                except (OSError, UnidentifiedImageError):
                    pass
            self.preview.configure(image="", text="照片文件找不到了")
            return
        self.preview_photo = None
        self.preview.configure(image="", text="还没有照片")

    def _show_preview(self, image: Image.Image) -> None:
        preview = image.copy()
        preview.thumbnail((self.theme.px(420), self.theme.px(210)), Image.Resampling.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(preview)
        self.preview.configure(image=self.preview_photo, text="")

    def choose_photo(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self,
            title="选择照片",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"), ("所有文件", "*.*")],
        )
        if not filename:
            return
        source = Path(filename)
        try:
            load_image(source)
        except ProfileError as error:
            messagebox.showerror("照片打不开", str(error), parent=self)
            return
        self.pending_photo = source
        self._show_photo()

    def paste_photo(self) -> None:
        image = clipboard_image()
        if image is None:
            messagebox.showinfo("粘贴照片", "剪贴板里没有图片。先截图或复制一张图片再试。", parent=self)
            return
        self.pending_photo = image
        self._show_photo()

    def _on_paste(self, _event: tk.Event) -> str | None:
        if isinstance(self.focus_get(), (tk.Entry, ttk.Entry)):
            return None
        self.paste_photo()
        return "break"

    def remove_photo(self) -> None:
        if self._current_photo() and self.pending_photo is None:
            self.drop_photo = self._current_photo()
        self.pending_photo = None
        self.photo_name = ""
        self._show_photo()

    # -- 保存 ------------------------------------------------------------- #
    def save(self) -> None:
        base = application_dir()
        name = self._current_photo()
        if self.pending_photo is not None:
            try:
                name = save_spot_photo(base, self.pending_photo)
            except ProfileError as error:
                messagebox.showerror("保存失败", str(error), parent=self)
                return
        note = self.note_var.get().strip()
        try:
            if self.index is None:
                entry = make_spot(self.number, self.model, note, self.cover, name)
                append_spot(base, entry)
                message = f"已记入拍车日志：{entry.number} {entry.model}"
            else:
                entry = SpotEntry(time=self.time_text, number=self.number, model=self.model, note=note, image=self.cover, photo=name)
                replace_spot(base, self.index, entry)
                message = f"已更新 {entry.number} 的记录"
        except (ProfileError, IndexError) as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        if self.drop_photo and self.drop_photo != name:
            try:
                delete_spot_photo(base, self.drop_photo)
            except ProfileError as error:
                messagebox.showwarning("旧照片未删除", str(error), parent=self)
        self.app.status_var.set(message)
        refresher = getattr(self.log_window, "reload", None)
        if callable(refresher):
            refresher()
        self.destroy()


class OnlineQueryWindow(WindowMixin, tk.Toplevel):
    """联网查询：查「这辆车跑过哪些线」和「这条线是哪两个方向」。

    三条硬规矩：
    - **只读**：这个窗口不写任何数据文件（写进线路数据是 V2 的事）；
    - **后台查询**：联网在工作线程里跑，界面不卡；服务端第一次握手可能要二十秒；
    - **离线不影响本地**：连不上就在窗口里说一句人话，其它功能照常。
    """

    COLUMNS_BUS = (
        TableColumn("line", "线路", 150, "w"),
        TableColumn("model", "车型", 190, "w"),
        TableColumn("spans", "时间段数", 90, "e", numeric=True),
        TableColumn("first", "最早", 100, "center"),
        TableColumn("last", "最近", 100, "center"),
        TableColumn("status", "状态", 110, "center"),
    )
    COLUMNS_ROUTE = (
        TableColumn("direction", "方向", 90, "center"),
        TableColumn("first_stop", "首站", 260, "w"),
        TableColumn("last_stop", "末站", 260, "w"),
    )

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.settings = online_settings()
        # 全程序共用一个会话与缓存：服务端建会话要 26 秒、每次查询 6.4 秒，不能每开一个窗口重付
        self.service = shared_service(self.settings, application_dir())
        self.busy = False
        self.mode = "bus"
        self.last_links: list[LineLink] = []
        self.last_checked = ""
        self.setup_dialog(app, "联网查询", 940, 640)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        self.dialog_header(
            frame,
            "联网查询",
            "数据来自第三方实时服务（bjbus.jinzihao.me），不是本程序的数据，仅供参考；"
            "只在你点查询时才联网，查不到不影响本地功能。",
            wrap=840,
        )

        search = ttk.Frame(frame)
        search.pack(fill="x", pady=(pad(10), pad(6)))
        ttk.Label(search, text="查询").pack(side="left", padx=(0, pad(6)))
        self.query_var = tk.StringVar()
        entry = ttk.Entry(search, textvariable=self.query_var, style="Search.TEntry")
        entry.pack(side="left", fill="x", expand=True)
        attach_placeholder(entry, "自编（如 31688）或线路号（如 886）", self.theme)
        entry.bind("<Return>", lambda _event: self.query("bus"))
        self.bus_button = ttk.Button(search, text="查这辆车", style="Primary.TButton", command=lambda: self.query("bus"))
        self.bus_button.pack(side="left", padx=pad(6))
        self.route_button = ttk.Button(search, text="查这条线", style="Primary.TButton", command=lambda: self.query("route"))
        self.route_button.pack(side="left")

        self.answer_var = tk.StringVar(value="")
        self.answer_label = ttk.Label(
            frame, textvariable=self.answer_var, style="CardStrong.TLabel", wraplength=pad(860), justify="left"
        )
        self.answer_label.pack(fill="x", pady=(0, pad(4)))
        self.detail_var = tk.StringVar(value="")
        self.detail_label = ttk.Label(
            frame, textvariable=self.detail_var, style="Muted.TLabel", wraplength=pad(860), justify="left"
        )
        self.detail_label.pack(fill="x", pady=(0, pad(8)))

        # 底部只放"这次查询能干什么"的两件事 + 一个搬运菜单 + 关闭；
        # 「清空缓存」这类维护操作收进 设置 → 联网（那里本来就有），别在这儿抢位置。
        actions = ttk.Frame(frame)
        actions.pack(side="bottom", fill="x", pady=(pad(10), 0))
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")
        compact_menu(
            actions,
            "复制 / 导出 ▾",
            [
                ("只复制「车型」列", lambda: self._copy_column("model", "车型")),
                ("复制选中行", self._copy_rows),
                ("导出 CSV", lambda: self.table.export(self, "联网查询结果.csv")),
            ],
        ).pack(side="right", padx=pad(8))
        self.write_button = ttk.Button(
            actions, text="写入线路数据", style="Primary.TButton", command=self.write_links
        )
        self.write_button.pack(side="right", padx=pad(8))
        self.write_button.state(["disabled"])
        self.where_button = ttk.Button(
            actions, text="现在在哪（实时）", style="Secondary.TButton", command=self.where_now
        )
        self.where_button.pack(side="right", padx=pad(8))
        self.where_button.state(["disabled"])
        self.state_var = tk.StringVar(value="")
        ttk.Label(actions, textvariable=self.state_var, style="CardMuted.TLabel").pack(side="left")

        self.table = DataTable(
            frame,
            self.theme,
            self.COLUMNS_BUS,
            sort_on_load=False,  # 保持"按最早出现时间"的顺序；想按字母排点表头
            empty_text="输入自编或线路号，点右边按钮；首次连接可能要一会儿",
        )
        self.table.card.pack(fill="both", expand=True)
        self.table.tree.configure(height=9)
        self.table.tree.bind("<Double-1>", lambda _event: self._open_selected())
        self.fit_dialog_height()

        self.refresh_state()
        entry.focus_set()

    # -- 状态 ------------------------------------------------------------- #
    def refresh_state(self) -> None:
        """联网没开 / 环境变量强制离线时，把按钮禁用并说明原因。"""
        usable = self.settings.usable
        for button in (self.bus_button, self.route_button):
            button.state(["!disabled"] if usable else ["disabled"])
        if not usable:
            reason = "程序被设为离线了（在设置里可以打开）" if offline_forced() else "设置里把联网关了"
            self.state_var.set(f"联网不可用（{reason}）；本地功能不受影响")
            return
        cached = self.service.cache_size
        mark = self.service.build_mark()
        measured = f"上次查询 {self.service.last_seconds:.1f}s　" if self.service.last_seconds else ""
        self.state_var.set(
            f"{measured}数据来自第三方服务，仅供参考；连不上就再点一次"
            + (f"　（查过的已存下 {cached} 条）" if cached else "")
            + (f"　{mark}" if mark else "")
        )

    def warm_up(self) -> None:
        """打开窗口就在后台把会话建起来，省得第一次查询还要现建连接。"""
        if not self.settings.usable or self.service.connected:
            return
        self._warm_queue: "queue.Queue[bool]" = queue.Queue()

        def work() -> None:
            self.service.warm_up()
            self._warm_queue.put(True)

        threading.Thread(target=work, name="bpt-online-warmup", daemon=True).start()
        self.after(200, self._poll_warmup)

    def _poll_warmup(self) -> None:
        try:
            self._warm_queue.get_nowait()
        except queue.Empty:
            if self.winfo_exists():
                self.after(200, self._poll_warmup)
            return
        if self.winfo_exists() and not self.busy:
            self.refresh_state()

    def where_now(self) -> None:
        """实时：这辆车现在跑到哪条线、哪一站（轨迹记录的最后一条）。"""
        text = self.query_var.get().strip()
        if not text or self.busy:
            return
        self.busy = True
        self.where_button.state(["disabled"])
        self.answer_var.set(f"查 {text} 现在在哪…")
        self.detail_var.set("查不到就说明它现在没在跑（或第三方数据里没有它）。")
        self.results: "queue.Queue[tuple]" = queue.Queue()

        def work() -> None:
            try:
                record, checked = self.service.track_latest(text)
                self.results.put(("ok", record, checked, ""))
            except OnlineError as error:
                self.results.put(("error", {}, "", error_text(error)))

        threading.Thread(target=work, name="bpt-online-where", daemon=True).start()
        self.after(120, self._poll_where)

    def _poll_where(self) -> None:
        try:
            kind, record, checked, message = self.results.get_nowait()
        except queue.Empty:
            if self.busy:
                self.after(120, self._poll_where)
            return
        self.busy = False
        self.where_button.state(["!disabled"])
        if kind == "error":
            self.answer_var.set("实时查询没成功")
            self.detail_var.set(message)
            return
        if not record:
            self.answer_var.set(f"{self.query_var.get().strip()}　现在没有在跑（或没有实时记录）")
            self.detail_var.set("只有车在跑的时候才有实时记录；隔几分钟再试。")
            return
        self.answer_var.set(f"{self.query_var.get().strip()}　{record.get('line_name', '（线路未知）')}")
        distance = record.get("distance_m")
        state = "已到站" if record.get("arrived") else (f"距站 {distance} 米" if distance is not None else "在途")
        self.detail_var.set(
            f"最后一笔：{record.get('time', '')}　在 {record.get('stop_name', '')}　{state}　"
            f"拥挤度 {record.get('occupancy', '—')}　（这一段轨迹共 {record.get('points', '0')} 个点）"
        )

    def clear_cache(self) -> None:
        count = self.service.clear_cache()
        self.state_var.set(f"已清空 {count} 条联网缓存" if count else "缓存本来就是空的")

    def write_links(self) -> None:
        """把这次查到的线路并进 线路.json（来源标成「网上查的」，可一键删）。"""
        if not self.last_links:
            messagebox.showinfo("写入线路数据", "先查一辆车，再写入。", parent=self)
            return
        plan = merge_lines(self.app.lines, self.last_links)
        if not plan.changed:
            messagebox.showinfo("写入线路数据", f"线路数据里已经有这些记录了（现在共 {len(self.app.lines)} 条）。", parent=self)
            return
        preview = "\n".join(
            f"　{link.line}　{link.first}–{link.last}　"
            f"{LINE_STATUS_LABELS.get(link.status, link.status)}　（{source_label(link.source)}）"
            for link in self.last_links[:8]
        )
        if not messagebox.askyesno(
            "写入线路数据",
            f"将把 {len(self.last_links)} 条查到记录并进现有 {len(self.app.lines)} 条：\n"
            f"新增 {len(plan.added)} · 合并 {len(plan.updated)}\n\n{preview}"
            f"{chr(10) + '　…' if len(self.last_links) > 8 else ''}\n\n"
            "来源会标成「网上查的」，查车时会显示出来；你自己记过的记录不会被覆盖。\n"
            "写入前自动备份 线路.json。确定继续吗？",
            parent=self,
        ):
            return
        merged = apply_line_plan(self.app.lines, plan)
        try:
            save_lines(application_dir(), merged)
        except ProfileError as error:
            messagebox.showerror("写入失败", str(error), parent=self)
            return
        self.app.lines = merged
        self.app.notify_data_changed({"lines"})
        self.app.status_var.set(f"已写入网上查的线路记录：线路数据现在共 {len(merged)} 条")
        self.state_var.set(f"已写入（线路数据共 {len(merged)} 条）")

    def close_dialog(self) -> None:
        """关窗口不关会话（会话是全程序共用的，留着给下次查询省 26 秒）。"""
        super().close_dialog()

    # -- 查询（后台线程） ------------------------------------------------- #
    def query(self, mode: str) -> None:
        text = self.query_var.get().strip()
        if not text:
            self.answer_var.set("请输入自编或线路号。")
            self.detail_var.set("")
            return
        if self.busy:
            return
        self.mode = mode
        self.busy = True
        for button in (self.bus_button, self.route_button):
            button.state(["disabled"])
        self.answer_var.set(f"查询中…（{text}）")
        self.detail_var.set("查询不会卡住界面；第一次连接会慢一点，之后就快了。")
        self.table.set_columns(self.COLUMNS_BUS if mode == "bus" else self.COLUMNS_ROUTE)
        self.table.set_rows([], keep_selection=False)
        self.results: "queue.Queue[tuple[list, str, str]]" = queue.Queue()
        threading.Thread(target=self._work, args=(mode, text), name="bpt-online-query", daemon=True).start()
        self.after(120, self._poll_results)
        self.warm_up()  # 顺手把会话建起来（已经在建就什么都不做）
    def _poll_results(self) -> None:
        """主线程轮询工作线程的结果。

        为什么不用 `self.after(0, …)` 直接从线程里回填：Tkinter 不是线程安全的，
        子线程里只要碰一下 Tk（连 `winfo_exists()` 都算）就会抛
        `RuntimeError: main thread is not in main loop`。所以线程只负责算，
        结果放进队列，由主线程的定时器取回来。
        """
        try:
            rows, answer, detail, links, checked = self.results.get_nowait()
        except queue.Empty:
            if self.busy:
                self.after(120, self._poll_results)
            return
        self._show(rows, answer, detail, links, checked)

    def _client(self) -> MCPClient:
        return self.service.client()

    def _work(self, mode: str, text: str) -> None:
        """在工作线程里查（结果与界面回填都通过队列交给主线程）。"""
        rows: list[tuple] = []
        answer = ""
        detail = ""
        links: list[LineLink] = []
        checked = ""
        try:
            if mode == "bus":
                history, checked = self.service.bus_history(text)     # 带缓存：命中就是 0 秒
                links = lines_from_history(
                    text,
                    history,
                    lambda number: model_lookup(self.app.ranges, number),
                    checked=checked,
                )
                if not links:
                    answer = f"{text}　没有查到运营记录"
                    detail = "没查到这辆车的线路历史（可能不在第三方数据里）。"
                else:
                    answer = f"{text}　跑过 {len(links)} 条线路"
                    spans = sum(len(history.get(link.line, ())) for link in links)
                    newest = max((link.last for link in links), key=month_key, default="")
                    detail = (
                        f"共 {spans} 个时间段；最近出现在 {newest}。"
                        "标「现在还在跑」的是最近 7 天内还出现过的。"
                        f"　{self.service.source_text()}"
                    )
                    rows = [
                        (
                            link.line,
                            link.model or "（车型未知）",
                            len(history.get(link.line, ())),
                            link.first,
                            link.last,
                            LINE_STATUS_LABELS.get(link.status, link.status),
                        )
                        for link in links
                    ]
            else:
                routes, checked = self.service.route_ids(text)
                if not routes:
                    answer = f"{text}　没有查到这条线路"
                    detail = "线路号要填对，例如 886。"
                else:
                    answer = f"{text}　{len(routes)} 个方向"
                    detail = f"首末站来自第三方数据。　{self.service.source_text()}"
                    rows = [
                        (f"方向 {index}", first_stop, last_stop)
                        for index, (_rid, first_stop, last_stop) in enumerate(routes, start=1)
                    ]
        except OnlineError as error:
            answer, detail = "联网查询没成功", error_text(error)
            self.service.close()  # 会话可能坏了：下一次查询会重新握手
        except Exception as error:  # noqa: BLE001 - 兜底：工作线程里的意外不能把界面弄死
            answer, detail = "联网查询出错", f"{type(error).__name__}: {error}"
        self.results.put((rows, answer, detail, links, checked))

    def _show(self, rows: list[tuple], answer: str, detail: str, links: list[LineLink], checked: str) -> None:
        self.busy = False
        self.answer_var.set(answer)
        self.detail_var.set(detail)
        self.table.set_rows(rows, keep_selection=False)
        self.last_links = list(links) if self.mode == "bus" else []
        self.last_checked = checked
        self.write_button.state(["!disabled"] if self.last_links else ["disabled"])
        self.where_button.state(["!disabled"] if self.mode == "bus" else ["disabled"])
        self.refresh_state()

    # -- 小工具 ----------------------------------------------------------- #
    def _copy_rows(self) -> None:
        count = self.table.copy()
        self.state_var.set(f"已复制 {count} 行" if count else "没有可复制的内容")

    def _copy_column(self, key: str, label: str) -> None:
        count = self.table.copy_column(key)
        self.state_var.set(f"已复制 {count} 个{label}" if count else f"没有可复制的{label}")

    def _open_selected(self) -> None:
        """双击：查到车型就打开车型详情（只读查询的顺手去处）。"""
        selection = self.table.tree.selection()
        if not selection or self.mode != "bus":
            return
        values = self.table.tree.item(selection[0], "values")
        model = str(values[1]) if len(values) > 1 else ""
        if model and model in set(model_names(self.app.ranges)):
            ModelDetailWindow(self, model)


class LineInfoWindow(WindowMixin, tk.Toplevel):
    """线路资料（联网）：一条线的两个方向、用过的车、现在在跑的车。

    - 首末站与站点表存进 `线路资料.json`（**内部编号不落盘**，服务端说它会变）；
    - 「用过的车」按车型归并后可一键写进 `线路.json`（来源=网上查的）；
    - 「现在在跑的车」这个接口偶尔不回包，失败时照实说一句，不弹错误框。
    """

    COLUMNS_USE = (
        TableColumn("model", "车型", 200, "w"),
        TableColumn("buses", "车辆数", 80, "e", numeric=True),
        TableColumn("first", "最早", 100, "center"),
        TableColumn("last", "最近", 100, "center"),
        TableColumn("status", "状态", 110, "center"),
    )
    COLUMNS_DIR = (
        TableColumn("direction", "方向", 90, "center"),
        TableColumn("first_stop", "首站", 260, "w"),
        TableColumn("last_stop", "末站", 260, "w"),
    )

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.settings = online_settings()
        self.service = shared_service(self.settings, application_dir())
        self.busy = False
        self.last_links: list[LineLink] = []
        self.last_checked = ""
        self.last_directions: list[tuple[str, str]] = []
        self.last_stops: list[tuple] = []
        self._catalog: dict = {}
        self._catalog_index: dict[str, list[tuple[str, str]]] | None = None
        self._catalog_date = ""
        self.setup_dialog(app, "线路资料（联网）", 980, 660)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        self.dialog_header(
            frame,
            "线路资料（联网）",
            "输入线路号点「查询」：看两个方向的首末站、这条线用过哪些车、现在有哪些车在跑。"
            "数据来自第三方实时服务，仅供参考；查到的首末站可以保存下来，以后不联网也能看。",
            wrap=880,
        )

        search = ttk.Frame(frame)
        search.pack(fill="x", pady=(pad(10), pad(6)))
        ttk.Label(search, text="线路号").pack(side="left", padx=(0, pad(6)))
        self.query_var = tk.StringVar()
        entry = ttk.Entry(search, textvariable=self.query_var, style="Search.TEntry")
        entry.pack(side="left", fill="x", expand=True)
        attach_placeholder(entry, "例如 886", self.theme)
        entry.bind("<Return>", lambda _event: self.query())
        self.query_button = ttk.Button(search, text="查询", style="Primary.TButton", command=self.query)
        self.query_button.pack(side="left", padx=pad(6))
        self.write_button = ttk.Button(search, text="写入线路数据", style="Primary.TButton", command=self.write_links)
        self.write_button.pack(side="left")
        self.write_button.state(["disabled"])
        # 「保存首末站」跟着结果走；「更新线路目录」是偶尔做一次的维护动作 → 收进菜单
        self.catalog_menu = compact_menu(
            search,
            "更多 ▾",
            [
                ("保存首末站到线路资料", self.save_info),
                ("更新线路目录", self.update_route_list),
            ],
        )
        self.catalog_menu.pack(side="left", padx=pad(6))

        self.answer_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.answer_var, style="CardStrong.TLabel", wraplength=pad(900), justify="left").pack(
            fill="x", pady=(0, pad(4))
        )
        self.detail_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.detail_var, style="Muted.TLabel", wraplength=pad(900), justify="left").pack(
            fill="x", pady=(0, pad(8))
        )

        actions = ttk.Frame(frame)
        actions.pack(side="bottom", fill="x", pady=(pad(10), 0))
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")
        compact_menu(
            actions,
            "复制 / 导出 ▾",
            [
                ("只复制「车型」列", lambda: self._copy_column("model", "车型")),
                ("导出 CSV", lambda: self.table.export(self, "线路资料.csv")),
            ],
        ).pack(side="right", padx=pad(8))
        self.state_var = tk.StringVar(value="")
        ttk.Label(actions, textvariable=self.state_var, style="CardMuted.TLabel").pack(side="left")

        self.notebook = ttk.Notebook(frame)
        self.notebook.pack(fill="both", expand=True)
        use_tab = ttk.Frame(self.notebook, padding=(pad(8), pad(8)))
        dir_tab = ttk.Frame(self.notebook, padding=(pad(8), pad(8)))
        stop_tab = ttk.Frame(self.notebook, padding=(pad(8), pad(8)))
        self.notebook.add(use_tab, text="① 用过的车（按车型）")
        self.notebook.add(dir_tab, text="② 两个方向的首末站")
        self.notebook.add(stop_tab, text="③ 站点表")
        self.table = DataTable(
            use_tab, self.theme, self.COLUMNS_USE, sort_on_load=False,
            empty_text="输入线路号后点「查询」——一条线可能有几百辆车",
        )
        self.table.card.pack(fill="both", expand=True)
        self.table.tree.configure(height=9)
        self.dir_table = DataTable(
            dir_tab, self.theme, self.COLUMNS_DIR, empty_text="还没有查到方向",
        )
        self.dir_table.card.pack(fill="both", expand=True)
        self.dir_table.tree.configure(height=6)
        self.stop_table = DataTable(
            stop_tab,
            self.theme,
            (
                TableColumn("seq", "序号", 60, "e", numeric=True),
                TableColumn("name", "站名", 200, "w"),
                TableColumn("distance", "距上一站", 110, "e"),
                TableColumn("lon", "经度", 110, "center"),
                TableColumn("lat", "纬度", 110, "center"),
            ),
            sort_on_load=False,
            empty_text="还没有查到站点（只查第一个方向的站序）",
        )
        self.stop_table.card.pack(fill="both", expand=True)
        self.stop_table.tree.configure(height=8)
        self.stop_hint_var = tk.StringVar(value="")
        ttk.Label(stop_tab, textvariable=self.stop_hint_var, style="CardMuted.TLabel", wraplength=pad(880)).pack(
            anchor="w", pady=(pad(6), 0)
        )
        ttk.Button(stop_tab, text="保存站点到线路资料", style="Secondary.TButton", command=self.save_stops).pack(
            anchor="w", pady=(pad(6), 0)
        )

        # ④ 实时（V5）：到站 + 时刻表。**实时数据不落盘**，只 30 秒内存缓存
        live_tab = ttk.Frame(self.notebook, padding=(pad(8), pad(8)))
        self.notebook.add(live_tab, text="④ 实时（到站 / 时刻表）")
        live_row = ttk.Frame(live_tab)
        live_row.pack(fill="x", pady=(0, pad(6)))
        ttk.Label(live_row, text="看哪个站").pack(side="left", padx=(0, pad(6)))
        self.stop_name_var = tk.StringVar()
        ttk.Entry(live_row, textvariable=self.stop_name_var, style="Search.TEntry").pack(
            side="left", fill="x", expand=True
        )
        self.arrivals_button = ttk.Button(live_row, text="这个站的到站", style="Secondary.TButton", command=self.load_arrivals)
        self.arrivals_button.pack(side="left", padx=pad(6))
        self.schedule_button = ttk.Button(live_row, text="每辆车几点发车", style="Secondary.TButton", command=self.load_schedules)
        self.schedule_button.pack(side="left", padx=(0, pad(6)))
        # 剩下这些问题都是"偶尔才问一次"：收进一个菜单，别让页签里堆一排按钮
        self.live_menu = compact_menu(
            live_row,
            "问别的 ▾",
            [
                ("车从哪个场站出车", self.load_depot),
                ("这条线的时刻表（概要）", self.load_timetable),
                ("「站名」到「到」那站要多久", self.load_journey),
                ("这个站有哪些线", self.load_routes_at_stop),
                ("从这里到这里有哪些直达线", self.load_routes_between),
            ],
        )
        self.live_menu.pack(side="left")
        self._live_controls = (self.arrivals_button, self.schedule_button, self.live_menu)

        # 终点站输入框（站名搜索 / 两站耗时都要用）
        search_row2 = ttk.Frame(live_tab)
        search_row2.pack(fill="x", pady=(0, pad(6)))
        ttk.Label(search_row2, text="到").pack(side="left", padx=(0, pad(4)))
        self.to_stop_var = tk.StringVar()
        ttk.Entry(search_row2, textvariable=self.to_stop_var, style="Search.TEntry").pack(side="left", fill="x", expand=True)
        self.live_hint_var = tk.StringVar(value="实时数据只在本次运行里缓存 30 秒；查不到不代表没有车。")
        ttk.Label(live_tab, textvariable=self.live_hint_var, style="CardMuted.TLabel", wraplength=pad(880)).pack(
            anchor="w", pady=(0, pad(6))
        )
        self.live_table = DataTable(
            live_tab,
            self.theme,
            (
                TableColumn("bus", "自编号", 100, "center"),
                TableColumn("time", "时间", 170, "center"),
                TableColumn("occupancy", "拥挤度", 90, "center"),
                TableColumn("extra", "说明", 320, "w"),
            ),
            sort_on_load=False,
            empty_text="「看这个站的到站」或「看时刻表」；两者都要先查过这条线路",
        )
        self.live_table.card.pack(fill="both", expand=True)
        self.live_table.tree.configure(height=8)
        self.fit_dialog_height()

        self.refresh_state()
        entry.focus_set()

    def refresh_state(self) -> None:
        usable = self.settings.usable
        self.query_button.state(["!disabled"] if usable else ["disabled"])
        if not usable:
            reason = "程序被设为离线了（在设置里可以打开）" if offline_forced() else "设置里把联网关了"
            self.state_var.set(f"联网不可用（{reason}）；本地功能不受影响")
            return
        catalog = self.route_catalog()
        extra = ""
        if catalog:
            extra = f"　线路目录 {len(catalog)} 条（{self._catalog_checked()}）"
        self.state_var.set(f"服务：{self.settings.base_url}（第三方数据，仅供参考）{extra}")

    # -- 本地线路目录（离线判断"这个线路号认不认"） ----------------------- #
    def route_catalog(self) -> dict[str, list[tuple[str, str]]]:
        """本地线路号索引；第一次用会从 线路目录.json 读。"""
        if self._catalog_index is None:
            payload = load_route_list(application_dir())
            self._catalog = payload.get("routes") or {}
            self._catalog_date = str(payload.get("checked", ""))
            self._catalog_index = route_list_numbers(self._catalog)
        return self._catalog_index

    def _catalog_checked(self) -> str:
        self.route_catalog()
        stamp = self._catalog_date
        return f"{stamp[:4]}-{stamp[4:6]}-{stamp[6:]} 更新" if len(stamp) == 8 else "日期未知"

    def update_route_list(self) -> None:
        """拉一份服务端全量线路目录存到本地（之后判断线路号就不用联网了）。"""
        if not self.settings.usable or self.busy:
            return
        self.busy = True
        self.catalog_menu.state(["disabled"])
        self.state_var.set("正在从网上更新线路目录…")
        self.results: "queue.Queue[tuple]" = queue.Queue()

        def work() -> None:
            try:
                routes, checked = self.service.route_list()
                save_route_list(application_dir(), routes, checked=checked)
                self.results.put(("ok", len(routes), checked, ""))
            except OnlineError as error:
                self.results.put(("error", 0, "", error_text(error)))
            except Exception as error:  # noqa: BLE001 - 兜底
                self.results.put(("error", 0, "", f"{type(error).__name__}: {error}"))

        threading.Thread(target=work, name="bpt-online-catalog", daemon=True).start()
        self.after(120, self._poll_catalog)

    def _poll_catalog(self) -> None:
        try:
            kind, count, checked, message = self.results.get_nowait()
        except queue.Empty:
            if self.busy:
                self.after(120, self._poll_catalog)
            return
        self.busy = False
        self.catalog_menu.state(["!disabled"])
        if kind == "error":
            self.state_var.set(f"线路目录没更新成功：{message}")
            return
        self._catalog_index = None  # 强制重读
        self.state_var.set(f"线路目录已更新：{count} 条（{checked}）—— 以后打错线路号不用等联网了")

    def _seed_from_catalog(self, text: str) -> None:
        """本地目录里有这条线就用它先把首末站填上（离线可用，不必等联网）。

        数据来源是 `线路目录.json`（用户自己拉下来的，带更新日期），所以填进去的
        首末站是有依据的；联网查到的新数据随后会覆盖它。
        """
        index = self.route_catalog()
        pairs = index.get(text)
        if not pairs:
            return
        self.last_directions = list(pairs)
        self.last_checked = self._catalog_date or today_stamp()
        self.dir_table.set_rows(
            [(f"方向 {index + 1}", first, last) for index, (first, last) in enumerate(pairs, start=1)],
            keep_selection=False,
        )
        self.state_var.set(
            f"首末站来自本地线路目录（{self._catalog_checked()}）；正在联网查更详细的数据…"
        )

    def _catalog_hint(self, text: str) -> str:
        """查之前先用本地目录给个提示：认不认、像不像打错、首末站是什么。"""
        index = self.route_catalog()
        if not index:
            return ""
        if text in index:
            pairs = index[text]
            stops = "；".join(f"{first} → {last}" for first, last in pairs[:2])
            return f"（本地目录里有这条线：{stops}）"
        similar = route_list_suggest(index, text)
        if similar:
            return f"（本地目录里没有「{text}」，你是不是要查：{'、'.join(similar)}）"
        return f"（本地目录里没有「{text}」——可能不在第三方数据里）"

    def query(self) -> None:
        text = self.query_var.get().strip()
        if not text:
            self.answer_var.set("请输入线路号。")
            return
        if self.busy:
            return
        hint = self._catalog_hint(text)
        self.busy = True
        self.query_button.state(["disabled"])
        self.answer_var.set(f"查询中…（{text}）")
        self.detail_var.set(
            (hint + "　" if hint else "") + "这条线用过的车可能很多；界面不会卡，查完自动显示。"
        )
        self.table.set_rows([], keep_selection=False)
        # 先清空、再用本地目录填首末站（顺序不能反：反了会被这句清空冲掉）
        self.dir_table.set_rows([], keep_selection=False)
        self._seed_from_catalog(text)
        self.results: "queue.Queue[tuple]" = queue.Queue()
        threading.Thread(target=self._work, args=(text,), name="bpt-online-line", daemon=True).start()
        self.after(120, self._poll)

    def _work(self, text: str) -> None:
        """工作线程：三个接口各查各的，哪个失败都不影响另外两个。"""
        lines: list[tuple] = []
        directions: list[tuple] = []
        buses_now = ""
        stops: list[tuple] = []
        stops_note = ""
        detail = ""
        links: list[LineLink] = []
        checked = ""
        failed = False
        try:
            ids, checked = self.service.route_ids(text)
            directions = [(first, last) for _rid, first, last in ids]
            history, checked2 = self.service.route_history(text)
            checked = checked2 or checked
            # 站点表（V4）：只查第一个方向的站序
            if ids:
                try:
                    stops, _stops_checked = self.service.route_stops(ids[0][0])
                    direction = directions[0] if directions else ("", "")
                    stops_note = (
                        f"{direction[0]} → {direction[1]}　共 {len(stops)} 站"
                        f"（站距是直线距离，仅供参考）"
                    )
                except OnlineError as error:
                    stops_note = f"站点表暂时查不到：{error_text(error)}"
            links = lines_from_route_history(
                text, history, lambda number: model_lookup(self.app.ranges, number), checked=checked
            )
            lines = [
                (
                    link.model or "（车型未知）",
                    len(link.numbers),
                    link.first,
                    link.last,
                    LINE_STATUS_LABELS.get(link.status, link.status),
                )
                for link in links
            ]
            # 「现在在跑的车」这个接口偶尔不回包：只影响这一块
            if ids:
                try:
                    buses, _checked = self.service.buses_on_route(ids[0][0])
                    with_model = {}
                    for bus in buses:
                        model = model_lookup(self.app.ranges, bus) or "（车型未知）"
                        with_model[model] = with_model.get(model, 0) + 1
                    buses_now = f"现在在跑 {len(buses)} 辆：" + "、".join(
                        f"{model} {count} 辆" for model, count in sorted(with_model.items(), key=lambda item: -item[1])
                    )
                except OnlineError as error:
                    buses_now = f"「现在在跑的车」暂时查不到（{error_text(error)}）"
            detail = f"首末站来自第三方数据。{self.service.source_text()}"
        except OnlineError as error:
            failed = True
            detail = error_text(error)
            self.service.close()  # 会话可能坏了：下次查询重新握手
        except Exception as error:  # noqa: BLE001 - 兜底
            failed = True
            detail = f"{type(error).__name__}: {error}"
        self.results.put((lines, directions, buses_now, detail, links, checked, stops, stops_note, failed))

    def _poll(self) -> None:
        try:
            lines, directions, buses_now, detail, links, checked, stops, stops_note, failed = self.results.get_nowait()
        except queue.Empty:
            if self.busy:
                self.after(120, self._poll)
            return
        self.busy = False
        self.query_button.state(["!disabled"] if self.settings.usable else ["disabled"])
        self.last_links = list(links)
        self.last_checked = checked or self.last_checked
        # 联网没给方向时（失败、或服务端没这条线）别把本地目录刚填好的首末站冲掉
        if directions:
            self.last_directions = list(directions)
            self.dir_table.set_rows(
                [(f"方向 {index}", first, last) for index, (first, last) in enumerate(directions, start=1)],
                keep_selection=False,
            )
        self.last_stops = list(stops)
        self.stop_hint_var.set(stops_note)
        text = self.query_var.get().strip()
        if failed:
            # 连不上 ≠ 没有这条线：标题要说清是"查询没成功"，否则用户会以为线路号错了
            self.answer_var.set(f"{text}　查询没成功（没连上联网服务）")
        elif not directions and not lines:
            self.answer_var.set(f"{text}　没有查到这条线路")
        else:
            self.answer_var.set(
                f"{text}　{len(directions)} 个方向 · "
                f"用过 {sum(len(link.numbers) for link in links)} 辆车 / {len(lines)} 种车型"
            )
        self.detail_var.set("　".join(part for part in (buses_now, detail) if part))
        self.table.set_rows(lines, keep_selection=False)
        self.stop_table.set_rows(stops, keep_selection=False)
        self.write_button.state(["!disabled"] if self.last_links else ["disabled"])
        self.refresh_state()

    # -- 落盘 ------------------------------------------------------------- #
    def save_info(self) -> None:
        """把首末站写进 线路资料.json（站点表是 V4 的事）。"""
        if not self.last_directions:
            messagebox.showinfo("线路资料", "先查一条线路。", parent=self)
            return
        line = self.query_var.get().strip()
        info = load_line_info(application_dir())
        updated = update_line_info(
            info, line, directions=self.last_directions, checked=self.last_checked or today_stamp()
        )
        try:
            save_line_info(application_dir(), updated)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.line_info = updated
        self.app.status_var.set(f"已保存 {line} 的线路资料（首末站）")
        self.state_var.set(f"已保存 {line} 的首末站到 {LINE_INFO_FILENAME}")

    def load_schedules(self) -> None:
        """这条线每辆车几点发车、几点到终点。"""
        self._route_live_call("schedules")

    def load_journey(self) -> None:
        """同一条线上，从「站名」到「到」那站，每辆车花了多久。"""
        self._route_live_call("journey")

    def load_depot(self) -> None:
        """这条线的车从哪个场站出车（找车/拍车有用）。"""
        self._route_live_call("depot")

    def _route_live_call(self, kind: str) -> None:
        if not self.last_directions:
            messagebox.showinfo("线路资料", "先查一条线路。", parent=self)
            return
        if self.busy:
            return
        line = self.query_var.get().strip()
        start = self.stop_name_var.get().strip()
        end = self.to_stop_var.get().strip()
        if kind == "journey" and (not start or not end):
            messagebox.showinfo("乘车耗时", "要填两个站：上面的「站名」是起点，「到」是终点（站名要和站牌一致）。", parent=self)
            return
        self.busy = True
        for button in self._live_controls:
            button.state(["disabled"])
        self.live_hint_var.set("正在查询…")
        self.results: "queue.Queue[tuple]" = queue.Queue()

        def work() -> None:
            try:
                ids, _checked = self.service.route_ids(line)
                if not ids:
                    self.results.put(("error", "这个线路号查不到", []))
                    return
                route_id = ids[0][0]
                if kind == "schedules":
                    rows, _stamp = self.service.schedules(route_id)
                    self.results.put((
                        "ok",
                        f"{line} 当天 {len(rows)} 个班次（发车 → 到达）",
                        [(bus, depart_time, "", f"{depart_stop} → {arrive_time} {arrive_stop}")
                         for bus, depart_time, depart_stop, arrive_time, arrive_stop in rows],
                    ))
                elif kind == "journey":
                    rows, _stamp = self.service.journey_time(route_id, start, end)
                    durations = [row[2] for row in rows if row[2]]
                    summary = f"最快 {durations[0]}" if durations else "没有班次"
                    self.results.put((
                        "ok",
                        f"{line}「{start}」→「{end}」当天 {len(rows)} 趟（{summary}）",
                        [(bus, start_time, duration, f"到 {end_time}") for bus, start_time, duration, end_time in rows],
                    ))
                else:
                    rows, _stamp = self.service.depot(route_id)
                    self.results.put((
                        "ok",
                        f"{line} 的出车记录 {len(rows)} 条（从场站出发的时间与位置）",
                        [(bus, time_text, "", f"{stop_name}　开往 {direction}") for bus, time_text, stop_name, direction in rows],
                    ))
            except OnlineError as error:
                self.results.put(("error", error_text(error), []))
            except Exception as error:  # noqa: BLE001 - 兜底
                self.results.put(("error", f"{type(error).__name__}: {error}", []))

        threading.Thread(target=work, name="bpt-online-route-live", daemon=True).start()
        self.after(120, self._poll_route_live)

    def _poll_route_live(self) -> None:
        try:
            kind, message, rows = self.results.get_nowait()
        except queue.Empty:
            if self.busy:
                self.after(120, self._poll_route_live)
            return
        self.busy = False
        for button in self._live_controls:
            button.state(["!disabled"])
        self.live_hint_var.set(message)
        self.live_table.set_rows(rows, keep_selection=False)

    def load_routes_at_stop(self) -> None:
        """这个站有哪些线（不需要先查线路）。"""
        stop = self.stop_name_var.get().strip()
        if not stop:
            messagebox.showinfo("站名搜索", "请在上面那个框里填站名（要和站牌一致，例如 德胜门西）。", parent=self)
            return
        self._search_call("stop", stop, self.to_stop_var.get().strip())

    def load_routes_between(self) -> None:
        """从这里到这里有哪些直达线路。"""
        start = self.stop_name_var.get().strip()
        end = self.to_stop_var.get().strip()
        if not start or not end:
            messagebox.showinfo("站名搜索", "起点和终点都要填（例如 德胜门西 → 昌平东关）。", parent=self)
            return
        self._search_call("between", start, end)

    def _search_call(self, kind: str, start: str, end: str) -> None:
        if self.busy:
            return
        self.busy = True
        self.live_menu.state(["disabled"])
        self.live_hint_var.set("正在查询…（站名要和站牌一致）")
        self.results: "queue.Queue[tuple]" = queue.Queue()

        def work() -> None:
            try:
                if kind == "stop":
                    rows, _stamp = self.service.routes_at_stop(start)
                    self.results.put((
                        "ok",
                        f"「{start}」有 {len(rows)} 条线路经过（实时数据，仅供参考）",
                        [(number, f"第 {stop_no} 站", "", f"{first} → {last}") for number, first, last, stop_no in rows],
                    ))
                else:
                    rows, _stamp = self.service.routes_between(start, end)
                    self.results.put((
                        "ok",
                        f"「{start}」→「{end}」有 {len(rows)} 条直达线路",
                        [(number, f"第 {from_no} 站上车", "", f"第 {to_no} 站下车　{first} → {last}")
                         for number, first, last, from_no, to_no in rows],
                    ))
            except OnlineError as error:
                self.results.put(("error", error_text(error), []))
            except Exception as error:  # noqa: BLE001 - 兜底
                self.results.put(("error", f"{type(error).__name__}: {error}", []))

        threading.Thread(target=work, name="bpt-online-search", daemon=True).start()
        self.after(120, self._poll_search)

    def _poll_search(self) -> None:
        try:
            kind, message, rows = self.results.get_nowait()
        except queue.Empty:
            if self.busy:
                self.after(120, self._poll_search)
            return
        self.busy = False
        self.live_menu.state(["!disabled"])
        self.live_hint_var.set(message)
        self.live_table.set_rows(rows, keep_selection=False)

    def load_arrivals(self) -> None:
        """实时：这条线在这个站的到站记录。"""
        self._live_call("arrivals")

    def load_timetable(self) -> None:
        """实时：这条线的时刻表概要。"""
        self._live_call("timetable")

    def _live_call(self, kind: str) -> None:
        if not self.last_directions:
            messagebox.showinfo("线路资料", "先查一条线路。", parent=self)
            return
        if self.busy:
            return
        line = self.query_var.get().strip()
        stop = self.stop_name_var.get().strip()
        if kind == "arrivals" and not stop:
            messagebox.showinfo("实时", "请填一个站名（要和站牌一致，例如 德胜门西）。", parent=self)
            return
        self.busy = True
        self.arrivals_button.state(["disabled"])
        self.live_menu.state(["disabled"])
        self.live_hint_var.set("正在查询…")
        self.results: "queue.Queue[tuple]" = queue.Queue()

        def work() -> None:
            try:
                ids, _checked = self.service.route_ids(line)
                if not ids:
                    self.results.put(("error", "这个线路号查不到", []))
                    return
                route_id = ids[0][0]
                if kind == "arrivals":
                    rows, _stamp = self.service.arrivals(route_id, stop)
                    self.results.put(
                        ("ok", f"{line} 在「{stop}」的到站（实时，仅供参考）", [(bus, time_text, occ, "") for bus, time_text, occ in rows])
                    )
                else:
                    rows, _stamp = self.service.timetable(route_id)
                    self.results.put(
                        ("ok", f"{line} 时刻表概要（{len(rows)} 趟）", [(bus, start, "", f"{start} → {end}，共 {stops} 站") for bus, start, end, stops in rows])
                    )
            except OnlineError as error:
                self.results.put(("error", error_text(error), []))
            except Exception as error:  # noqa: BLE001 - 兜底
                self.results.put(("error", f"{type(error).__name__}: {error}", []))

        threading.Thread(target=work, name="bpt-online-live", daemon=True).start()
        self.after(120, self._poll_live)

    def _poll_live(self) -> None:
        try:
            kind, message, rows = self.results.get_nowait()
        except queue.Empty:
            if self.busy:
                self.after(120, self._poll_live)
            return
        self.busy = False
        self.arrivals_button.state(["!disabled"])
        self.live_menu.state(["!disabled"])
        self.live_hint_var.set(message)
        self.live_table.set_rows(rows, keep_selection=False)

    def save_stops(self) -> None:
        """把站点表写进 线路资料.json。"""
        if not self.last_stops:
            messagebox.showinfo("线路资料", "先查一条线路，而且这条线要有站点表。", parent=self)
            return
        line = self.query_var.get().strip()
        info = load_line_info(application_dir())
        updated = update_line_info(info, line, stops=self.last_stops, checked=self.last_checked or today_stamp())
        try:
            save_line_info(application_dir(), updated)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.line_info = updated
        self.app.status_var.set(f"已保存 {line} 的 {len(self.last_stops)} 个站点")
        self.stop_hint_var.set(f"已保存 {len(self.last_stops)} 个站点到 {LINE_INFO_FILENAME}")

    def write_links(self) -> None:
        """把「用过的车」按车型写进 线路.json（来源=网上查的）。"""
        if not self.last_links:
            messagebox.showinfo("写入线路数据", "先查一条线路。", parent=self)
            return
        plan = merge_lines(self.app.lines, self.last_links)
        if not plan.changed:
            messagebox.showinfo("写入线路数据", f"线路数据里已经有这些记录了（现在共 {len(self.app.lines)} 条）。", parent=self)
            return
        preview = "\n".join(
            f"　{link.line}　{link.model or '（车型未知）'}　{len(link.numbers)} 辆　{link.first}–{link.last}"
            for link in self.last_links[:8]
        )
        if not messagebox.askyesno(
            "写入线路数据",
            f"将把 {len(self.last_links)} 条（按车型归并）并进现有 {len(self.app.lines)} 条：\n"
            f"新增 {len(plan.added)} · 合并 {len(plan.updated)}\n\n{preview}"
            f"{chr(10) + '　…' if len(self.last_links) > 8 else ''}\n\n"
            "来源标成「网上查的」，你自己记过的不会被覆盖。写入前自动备份 线路.json。确定继续吗？",
            parent=self,
        ):
            return
        merged = apply_line_plan(self.app.lines, plan)
        try:
            save_lines(application_dir(), merged)
        except ProfileError as error:
            messagebox.showerror("写入失败", str(error), parent=self)
            return
        self.app.lines = merged
        self.app.notify_data_changed({"lines"})
        self.app.status_var.set(f"已写入网上查的线路记录：线路数据现在共 {len(merged)} 条")
        self.state_var.set(f"已写入（线路数据共 {len(merged)} 条）")

    def _copy_column(self, key: str, label: str) -> None:
        count = self.table.copy_column(key)
        self.state_var.set(f"已复制 {count} 个{label}" if count else f"没有可复制的{label}")


class DistrictBusEntryWindow(WindowMixin, tk.Toplevel):
    """记一辆区属公交：自己坐一次就记一条「车牌 + 车型」（可带线路）。

    其他区基本只有车牌、没有自编，靠号段推不出来，只能这样一条条攒。
    """

    def __init__(
        self,
        parent: tk.Misc,
        *,
        entry: DistrictBusEntry | None = None,
        cars: list[str] | None = None,
        hint: str = "",
    ) -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)
        self.theme = self.app.theme
        self.entry = entry
        self.result: DistrictBusEntry | None = None
        self.setup_dialog(parent, "改一辆区属公交" if entry else "记一辆区属公交", 540, 430)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        self.dialog_header(
            frame,
            "改一辆区属公交" if entry else "记一辆区属公交",
            "填车牌和车型就行（线路、公司、备注都能空着）。同一个车牌记成不同车型时程序会提醒你，"
            "不会直接覆盖。" + (f"\n{hint}" if hint else ""),
            wrap=470,
        )

        body = ttk.Frame(frame)
        body.pack(fill="both", expand=True, pady=(pad(10), 0))
        body.columnconfigure(1, weight=1)
        self.vars: dict[str, tk.StringVar] = {}
        fields = (
            ("plate", "车牌 *", "例如 京B12345"),
            ("model", "车型 *", "例如 XMQ6900"),
            ("line", "线路", "例如 昌1路"),
            ("company", "公司 / 区县", ""),
            ("date", "记录日期", "例如 2026.9.30"),
            ("note", "备注", ""),
        )
        for row, (key, label, hint_text) in enumerate(fields):
            ttk.Label(body, text=label, style="CardMuted.TLabel").grid(row=row, column=0, sticky="w", pady=pad(4))
            variable = tk.StringVar(value=getattr(entry, key, "") if entry else "")
            self.vars[key] = variable
            if key == "model" and cars:
                widget: tk.Widget = ttk.Combobox(body, textvariable=variable, values=cars)
            else:
                widget = ttk.Entry(body, textvariable=variable)
            widget.grid(row=row, column=1, sticky="ew", padx=(pad(10), 0), pady=pad(4))
            widget.bind("<Return>", lambda _event: self.save())
            if hint_text:
                ttk.Label(body, text=hint_text, style="CardMuted.TLabel").grid(row=row, column=2, sticky="w", padx=pad(6))
            if key == "plate":
                widget.focus_set()

        actions = ttk.Frame(frame)
        actions.pack(side="bottom", fill="x", pady=(pad(12), 0))
        ttk.Button(actions, text="保存", style="Primary.TButton", command=self.save).pack(side="right")
        ttk.Button(actions, text="取消", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))
        self.fit_dialog_height()

    def save(self) -> None:
        plate = self.vars["plate"].get().strip()
        model = self.vars["model"].get().strip()
        if not plate:
            messagebox.showinfo("记一辆区属公交", "车牌不能为空。", parent=self)
            return
        if not model:
            messagebox.showinfo("记一辆区属公交", "车型不能为空——「车牌 + 车型」才是这条记录的价值。", parent=self)
            return
        self.result = DistrictBusEntry(
            company=self.vars["company"].get().strip(),
            plate=normalize_plate(plate),
            line=self.vars["line"].get().strip(),
            model=model,
            note=self.vars["note"].get().strip(),
            date=self.vars["date"].get().strip(),
        )
        self.close_dialog()


class DistrictBusWindow(WindowMixin, tk.Toplevel):
    """区属公交查询：按自编查是哪家公司、哪年哪批第几辆、什么车型。

    区属公交和公交集团不是一套编号：各区/各公司自己编，有的区只有车牌没有自编。
    编号规则挂在公司上（前缀、位数、年份/批次/序号各占哪几位），所以换一家公司
    只要在 区属公交数据.json 里加一条 companies 就能用。
    """

    COLUMNS = (
        ("company", "公司", 190, "w"),
        ("span", "号段", 200, "center"),
        ("count", "实车数", 70, "e"),
        ("batch", "年份 / 批次", 130, "center"),
        ("model", "车型", 170, "w"),
        ("skip", "未编", 120, "w"),
    )

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "区属公交查询", 980, 620)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)

        self.dialog_header(
            frame,
            "区属公交查询",
            "区属公交的编号跟公交集团不是一套：昌平有 8 位自编（前两位 59 是公司缩写，第 3-4 位购车年份，"
            "第 5 位批次，后三位是当年第几辆）；其他区目前基本只有车牌号、没有自编，所以查询框自编和车牌都能输。",
            wrap=860,
        )

        search_row = ttk.Frame(frame)
        search_row.pack(fill="x", pady=(pad(10), pad(8)))
        ttk.Label(search_row, text="自编 / 车牌").pack(side="left", padx=(0, pad(6)))
        self.number_var = tk.StringVar()
        entry = ttk.Entry(search_row, textvariable=self.number_var, style="Search.TEntry")
        entry.pack(side="left", fill="x", expand=True)
        attach_placeholder(entry, "自编（如 59181007）或车牌（如 京B12345）", self.theme)
        ttk.Button(search_row, text="查询", style="Primary.TButton", command=self.query).pack(side="left", padx=pad(6))
        entry.bind("<Return>", lambda _event: self.query())

        self.answer_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.answer_var, style="CardStrong.TLabel", wraplength=pad(880), justify="left").pack(
            fill="x", pady=(0, pad(4))
        )
        self.detail_var = tk.StringVar(value="")
        self.detail_label = ttk.Label(frame, textvariable=self.detail_var, style="Muted.TLabel", wraplength=pad(880), justify="left")
        self.detail_label.pack(fill="x", pady=(0, pad(4)))
        # 其他区（只有车牌、没有自编）的现状说明：默认收起，点「数据说明」才展开
        self.reserved_var = tk.StringVar(value="")
        self.reserved_label = ttk.Label(
            frame, textvariable=self.reserved_var, style="CardMuted.TLabel", wraplength=pad(880), justify="left"
        )
        self.explain_shown = False

        actions = ttk.Frame(frame)
        actions.pack(side="bottom", fill="x", pady=(pad(10), 0))
        # 先给右边的按钮占位（先到先得），中间的统计文字最后铺、剩多少显示多少
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")
        self.tools_button = ttk.Button(actions, text="数据说明 ▾", style="Ghost.TButton", command=self.toggle_explain)
        self.tools_button.pack(side="right", padx=pad(8))
        ttk.Button(actions, text="重新载入", style="Secondary.TButton", command=self.reload).pack(side="left")
        ttk.Button(actions, text="建立空模板", style="Ghost.TButton", command=self.create_template).pack(side="left")
        self.summary_var = tk.StringVar(value="")
        ttk.Label(actions, textvariable=self.summary_var, style="CardMuted.TLabel").pack(
            side="left", fill="x", expand=True, padx=pad(10)
        )

        self.notebook = ttk.Notebook(frame)
        self.notebook.pack(fill="both", expand=True)
        ranges_tab = ttk.Frame(self.notebook, padding=(pad(8), pad(8)))
        plates_tab = ttk.Frame(self.notebook, padding=(pad(8), pad(8)))
        self.notebook.add(ranges_tab, text="① 自编号段（昌平）")
        self.notebook.add(plates_tab, text="② 车牌记录（坐一次记一条）")

        self.table = DataTable(
            ranges_tab,
            self.theme,
            tuple(TableColumn(key, title, width, anchor, numeric=(key == "count")) for key, title, width, anchor in self.COLUMNS),
            toolbar=self._build_toolbar,
            empty_text=f"还没有 {DISTRICT_BUS_FILENAME}：可以先点「建立一个空的区属公交数据.json」再填数据",
        )
        self.table.card.pack(fill="both", expand=True)
        self.table.tree.configure(height=9)

        # ② 车牌记录：其他区没有自编，只能自己坐一次记一条「车牌 + 车型」
        self.entry_table = DataTable(
            plates_tab,
            self.theme,
            (
                TableColumn("plate", "车牌", 130, "w"),
                TableColumn("model", "车型", 190, "w"),
                TableColumn("line", "线路", 120, "w"),
                TableColumn("company", "公司 / 区县", 150, "w"),
                TableColumn("date", "记录日期", 100, "center"),
                TableColumn("note", "备注", 150, "w"),
            ),
            toolbar=self._build_plate_toolbar,
            selectmode="extended",
            on_delete=self.delete_entry,
            on_search=lambda: entry.focus_set(),
            empty_text="还没有车牌记录：点「＋ 记一辆」把坐过的车牌和车型记下来（也可以批量粘贴）",
        )
        self.entry_table.card.pack(fill="both", expand=True, pady=(pad(6), 0))
        self.entry_table.tree.configure(height=7)
        self.entry_table.tree.bind("<Double-1>", lambda _event: self.edit_entry())

        self.fit_dialog_height()
        self.refresh()

        companies, _entries, ranges, _empty = self.app.district_bus
        if companies:
            self.query("59181001")  # 有数据就顺手演示一条，省得空着让人不知道能查什么

    def _build_plate_toolbar(self, bar: ttk.Frame) -> None:
        pad = self.theme.px
        # 常用的一排；其余收进「更多 ▾」——按钮挤成一行会被窗口切掉
        row = ttk.Frame(bar)
        row.pack(fill="x")
        # 先给「更多 ▾」占位：pack 是先到先得，长文字若先铺会把按钮挤出去
        self.more_button = ttk.Button(row, text="更多 ▾", style="Ghost.TButton", command=self.toggle_plate_tools)
        self.more_button.pack(side="right")
        for text, style, command in (
            ("＋ 记一辆", "Primary.TButton", self.add_entry),
            ("＋ 同一车型批量录", "Primary.TButton", self.batch_by_model),
            ("批量设置公司/线路", "Secondary.TButton", self.batch_set_company),
            ("修改选中", "Secondary.TButton", self.edit_entry),
            ("删除选中", "Danger.TButton", self.delete_entry),
        ):
            ttk.Button(row, text=text, style=style, command=command).pack(side="left", padx=(0, pad(6)))
        # 统计文字最后铺：剩多少显示多少，不再挤按钮
        self.plate_count_var = tk.StringVar(value="")
        ttk.Label(row, textvariable=self.plate_count_var, style="CardMuted.TLabel").pack(
            side="left", fill="x", expand=True, padx=pad(8)
        )

        self.plate_tools = ttk.Frame(bar)
        self.plate_tools_shown = False
        ttk.Button(self.plate_tools, text="复制车型", style="Ghost.TButton",
                   command=lambda: self._copy_column(self.entry_table, "model", "车型")).pack(side="left", padx=(0, pad(6)))
        ttk.Button(self.plate_tools, text="复制车牌", style="Ghost.TButton",
                   command=lambda: self._copy_column(self.entry_table, "plate", "车牌")).pack(side="left", padx=(0, pad(6)))
        ttk.Button(self.plate_tools, text="批量粘贴（车牌+车型）", style="Ghost.TButton", command=self.paste_entries).pack(
            side="left", padx=(0, pad(6))
        )
        ttk.Button(self.plate_tools, text="检查重复", style="Ghost.TButton", command=self.check_duplicates).pack(
            side="left", padx=(0, pad(6))
        )
        ttk.Button(
            self.plate_tools, text="⇩ 导出 CSV", style="Ghost.TButton",
            command=lambda: self.entry_table.export(self, "区属公交车牌记录.csv"),
        ).pack(side="left")

    def toggle_plate_tools(self) -> None:
        """展开 / 收起车牌页签的次要按钮。"""
        if self.plate_tools_shown:
            self.plate_tools.pack_forget()
            self.more_button.configure(text="更多 ▾")
        else:
            self.plate_tools.pack(fill="x", pady=(self.theme.px(6), 0))
            self.more_button.configure(text="收起 ▴")
        self.plate_tools_shown = not self.plate_tools_shown

    # -- 同一车型批量录车牌 ----------------------------------------------- #
    def _copy_column(self, table: DataTable, key: str, label: str) -> None:
        """只复制某一列（选中优先，没选中就全部），并给一句反馈。"""
        count = table.copy_column(key)
        if not count:
            messagebox.showinfo("复制", f"没有可复制的{label}：先选中几行，或先把数据查出来。", parent=self)
            return
        self.app.status_var.set(f"已复制 {count} 个{label}到剪贴板")

    def batch_by_model(self) -> None:
        """选一个车型，把这一批车牌一次录进去（先预演，再说写不写）。"""
        entries = list(self.app.district_bus[1])
        window = tk.Toplevel(self)
        window.title("同一车型批量录车牌")
        window.transient(self)
        window.configure(background=self.theme.background)
        pad = self.theme.px
        ttk.Label(
            window,
            text="先填车型，再把这一批车牌贴进来（每行一个，或用空格/逗号隔开）\n例：京B12345 京B12346 京B12347",
            style="Muted.TLabel",
            justify="left",
        ).pack(anchor="w", padx=pad(14), pady=(pad(14), pad(6)))

        selection = self.entry_table.tree.selection()
        start_model = self.entry_table.tree.item(selection[0], "values")[1] if selection else ""
        form = ttk.Frame(window)
        form.pack(fill="x", padx=pad(14))
        ttk.Label(form, text="车型 *").pack(side="left")
        model_var = tk.StringVar(value=start_model)
        ttk.Combobox(form, textvariable=model_var, values=district_bus_entry_cars(entries), width=24).pack(
            side="left", padx=pad(8)
        )
        ttk.Label(form, text="线路（可空）").pack(side="left", padx=(pad(10), pad(4)))
        line_var = tk.StringVar()
        ttk.Entry(form, textvariable=line_var, width=14).pack(side="left")

        box = tk.Text(window, height=10, width=54)
        box.pack(padx=pad(14), pady=pad(8), fill="both", expand=True)
        box.focus_set()
        preview = tk.StringVar(value="")
        ttk.Label(window, textvariable=preview, style="CardMuted.TLabel", justify="left", wraplength=pad(430)).pack(
            anchor="w", padx=pad(14)
        )

        def plates_now() -> list[str]:
            return district_bus_split_plates(box.get("1.0", "end"))

        def refresh_preview(*_args) -> None:
            plan = district_bus_plan_batch(entries, plates_now(), model_var.get(), line_var.get())
            bits = [f"会新增 {len(plan.added)} 个"]
            if plan.refreshed:
                bits.append(f"已有同车型 {len(plan.refreshed)} 个（只补上空的）")
            if plan.conflicts:
                bits.append(f"⚠ 已记成别的车型 {len(plan.conflicts)} 个（不动）")
            if plan.duplicates:
                bits.append(f"这一批里重复 {len(plan.duplicates)} 个（只算一次）")
            if plan.invalid:
                bits.append(f"认不出是车牌 {len(plan.invalid)} 个")
            preview.set("　".join(bits))

        box.bind("<KeyRelease>", refresh_preview)
        model_trace = model_var.trace_add("write", lambda *_args: refresh_preview())
        refresh_preview()

        def close_window() -> None:
            """关这个内联对话框前，先把 trace 摘掉。

            不摘的话：`model_var.trace_add` 在 Tcl 里注册的命令会持有这个闭包，
            闭包又持有窗口和变量 —— Python GC 看不见这层环，于是每开一次就永久多一条 Tcl 命令
            （实测就是这个对话框让"反复开关窗口"每轮涨 1 条）。
            """
            for variable, name in ((model_var, model_trace),):
                try:
                    variable.trace_remove("write", name)
                except tk.TclError:
                    pass
            window.destroy()

        def apply() -> None:
            model = model_var.get().strip()
            if not model:
                messagebox.showinfo("批量录入", "先填车型——这一批车牌都是这个车型的。", parent=window)
                return
            updated, plan = district_bus_apply_batch(entries, plates_now(), model, line_var.get())
            close_window()
            if not plan.total_ok:
                messagebox.showinfo("批量录入", "没有可新增的车牌（可能都已经录过了）。", parent=self)
                return
            self._save_entries(updated, f"同一车型批量录入：{model} 新增 {len(plan.added)}、更新 {len(plan.refreshed)}")
            if plan.conflicts:
                messagebox.showwarning(
                    "有车牌记过别的车型",
                    "这些车牌已经记成别的车型，没有改动：\n"
                    + "\n".join(f"{plate}（原记 {old}）" for plate, old in plan.conflicts[:8]),
                    parent=self,
                )

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=pad(14), pady=pad(12))
        ttk.Button(buttons, text="录进去", style="Primary.TButton", command=apply).pack(side="right")
        ttk.Button(buttons, text="取消", style="Secondary.TButton", command=close_window).pack(side="right", padx=pad(8))
        window.bind("<Escape>", lambda _event: close_window())

    def batch_set_company(self) -> None:
        """把选中的几条记录一次设成同一个公司 / 线路（填哪项改哪项）。"""
        selection = self.entry_table.tree.selection()
        if not selection:
            messagebox.showinfo("批量设置", "请先在表里选中要改的记录（Ctrl / Shift 可多选）。", parent=self)
            return
        keys = [str(self.entry_table.tree.item(iid, "values")[0]) for iid in selection]
        entries = list(self.app.district_bus[1])
        window = tk.Toplevel(self)
        window.title(f"批量设置（{len(keys)} 条）")
        window.transient(self)
        window.configure(background=self.theme.background)
        pad = self.theme.px
        ttk.Label(
            window,
            text=f"要给选中的 {len(keys)} 条记录改什么？只填要改的项，留空的不动。\n"
                 "（公司/区县留一个空格可以清空这一项）",
            style="Muted.TLabel",
            justify="left",
        ).pack(anchor="w", padx=pad(14), pady=(pad(14), pad(8)))

        body = ttk.Frame(window)
        body.pack(fill="x", padx=pad(14))
        body.columnconfigure(1, weight=1)
        company_var = tk.StringVar()
        line_var = tk.StringVar()
        ttk.Label(body, text="公司 / 区县").grid(row=0, column=0, sticky="w", pady=pad(4))
        ttk.Combobox(body, textvariable=company_var, values=district_bus_companies_seen(entries), width=28).grid(
            row=0, column=1, sticky="ew", padx=(pad(8), 0), pady=pad(4)
        )
        ttk.Label(body, text="线路").grid(row=1, column=0, sticky="w", pady=pad(4))
        ttk.Entry(body, textvariable=line_var, width=28).grid(row=1, column=1, sticky="ew", padx=(pad(8), 0), pady=pad(4))

        def apply() -> None:
            company = company_var.get().strip() if company_var.get() else None
            line = line_var.get().strip() if line_var.get() else None
            if company is None and line is None:
                messagebox.showinfo("批量设置", "公司/区县和线路都没填，没有要改的内容。", parent=window)
                return
            updated, changed = district_bus_set_fields(entries, keys, company=company, line=line)
            window.destroy()
            if not changed:
                messagebox.showinfo("批量设置", "选中的记录本来就是这些值，没有变化。", parent=self)
                return
            what = "、".join(part for part in (f"公司={company}" if company else "", f"线路={line}" if line else "") if part)
            self._save_entries(updated, f"批量设置 {changed} 条：{what}")

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=pad(14), pady=pad(12))
        ttk.Button(buttons, text="设置", style="Primary.TButton", command=apply).pack(side="right")
        ttk.Button(buttons, text="取消", style="Secondary.TButton", command=window.destroy).pack(side="right", padx=pad(8))
        window.bind("<Escape>", lambda _event: window.destroy())

    def check_duplicates(self) -> None:
        """检查重复上传：同车牌多条 / 同车牌不同车型 / 缺字段，可一键合并。"""
        entries = list(self.app.district_bus[1])
        quality = district_bus_quality(entries)
        lines: list[str] = []
        if quality["duplicates"]:
            lines.append("【重复上传】同一个车牌有多条记录：" + "、".join(f"{key}×{count}" for key, count in quality["duplicates"][:10]))
        if quality["conflicts"]:
            lines.append(
                "【车型冲突】同一个车牌记了不同车型："
                + "、".join(f"{key}（{'/'.join(models)}）" for key, models in quality["conflicts"][:10])
            )
        if quality["no_model"]:
            lines.append("【缺车型】" + "、".join(quality["no_model"][:10]))
        if quality["no_key"]:
            lines.append("【缺车牌/自编】" + "、".join(quality["no_key"][:10]))
        if not lines:
            messagebox.showinfo(
                "检查重复",
                f"没有发现问题：{len(entries)} 条记录里没有重复上传的车牌，也没有车型冲突。",
                parent=self,
            )
            return
        if quality["duplicates"] and messagebox.askyesno(
            "检查重复",
            "\n\n".join(lines) + f"\n\n要现在把重复的合并成一条吗？（共 {len(entries)} 条记录）",
            parent=self,
        ):
            merged, removed = district_bus_merge_duplicates(entries)
            self._save_entries(merged, f"已合并重复车牌：去掉 {removed} 条")
            return
        messagebox.showwarning("检查重复", "\n\n".join(lines), parent=self)

    # -- 车牌记录的增改删 ------------------------------------------------ #
    def _save_entries(self, entries: list[DistrictBusEntry], message: str) -> None:
        companies, _old, ranges, empty = self.app.district_bus
        try:
            save_district_bus(application_dir(), companies, ranges, empty, entries)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.reload_district_bus()
        self.refresh()
        self.app.status_var.set(message)

    def add_entry(self) -> None:
        window = DistrictBusEntryWindow(self, cars=district_bus_entry_cars(self.app.district_bus[1]))
        self.wait_window(window)
        if window.result is None:
            return
        entries = list(self.app.district_bus[1])
        updated, what, message = upsert_district_bus_entry(entries, window.result)
        if what == "冲突":
            messagebox.showwarning("这个车牌记过别的车型", message + "\n\n没有改动原来的记录。要改请用「修改选中」。", parent=self)
            return
        self._save_entries(updated, message)

    def _selected_entry(self) -> DistrictBusEntry | None:
        selection = self.entry_table.tree.selection()
        if not selection:
            return None
        key = str(self.entry_table.tree.item(selection[0], "values")[0])
        return next((item for item in self.app.district_bus[1] if item.label == key), None)

    def edit_entry(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            messagebox.showinfo("修改", "请先在表里选中一条记录。", parent=self)
            return
        window = DistrictBusEntryWindow(self, entry=entry, cars=district_bus_entry_cars(self.app.district_bus[1]))
        self.wait_window(window)
        if window.result is None:
            return
        entries = remove_district_bus_entry(list(self.app.district_bus[1]), entry.key)
        entries.append(window.result)
        self._save_entries(entries, f"{window.result.label} 已更新")

    def delete_entry(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            messagebox.showinfo("删除", "请先在表里选中一条记录。", parent=self)
            return
        if not messagebox.askyesno("删除记录", f"删掉 {entry.label}（{entry.model}）这条记录？", parent=self):
            return
        self._save_entries(remove_district_bus_entry(list(self.app.district_bus[1]), entry.key), f"已删除 {entry.label}")

    def paste_entries(self) -> None:
        """批量粘贴：每行「车牌 车型 [线路]」（空格、逗号、制表符都行）。"""
        window = tk.Toplevel(self)
        window.title("批量粘贴车牌记录")
        window.transient(self)
        window.configure(background=self.theme.background)
        pad = self.theme.px
        ttk.Label(
            window,
            text="每行一条：车牌 车型 [线路]　　例：\n京B12345 XMQ6900 昌1路\n京G54321 · 专2路",
            style="Muted.TLabel",
            justify="left",
        ).pack(anchor="w", padx=pad(14), pady=(pad(14), pad(6)))
        box = tk.Text(window, height=10, width=52)
        box.pack(padx=pad(14), fill="both", expand=True)
        box.focus_set()

        def apply() -> None:
            added = 0
            conflicts: list[str] = []
            entries = list(self.app.district_bus[1])
            for line in box.get("1.0", "end").splitlines():
                parts = [part for part in re.split(r"[\s,，、\t]+", line.strip()) if part]
                if len(parts) < 2:
                    continue
                plate, model = parts[0], parts[1]
                line_name = parts[2] if len(parts) > 2 else ""
                plate = normalize_plate(plate)
                if not looks_like_plate(plate):
                    continue
                entries, what, message = upsert_district_bus_entry(
                    entries, DistrictBusEntry(company="", plate=plate, model=model, line=line_name)
                )
                if what == "冲突":
                    conflicts.append(message.split("，")[0])
                elif what in ("新增", "更新"):
                    added += 1
            window.destroy()
            if not added and not conflicts:
                messagebox.showinfo("批量粘贴", "没有认出可用的行（每行至少要有车牌和车型）。", parent=self)
                return
            self._save_entries(entries, f"批量粘贴：记下/更新 {added} 条" + (f"，{len(conflicts)} 条与已有记录冲突未改" if conflicts else ""))
            if conflicts:
                messagebox.showwarning("有冲突没改", "这些车牌已经记过别的车型，需你自己判断：\n" + "\n".join(conflicts[:8]), parent=self)

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=pad(14), pady=pad(12))
        ttk.Button(buttons, text="记下来", style="Primary.TButton", command=apply).pack(side="right")
        ttk.Button(buttons, text="取消", style="Secondary.TButton", command=window.destroy).pack(side="right", padx=pad(8))
        window.bind("<Escape>", lambda _event: window.destroy())

    def _build_toolbar(self, bar: ttk.Frame) -> None:
        pad = self.theme.px
        self.model_var = tk.StringVar(value="全部车型")
        self.model_combo = ttk.Combobox(bar, textvariable=self.model_var, state="readonly", width=22)
        self.model_combo.pack(side="left")
        self.model_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh())
        self.count_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.count_var, style="CardMuted.TLabel").pack(side="left", padx=pad(10))
        ttk.Button(bar, text="复制车型", style="Ghost.TButton",
                   command=lambda: self._copy_column(self.table, "model", "车型")).pack(side="right")

    # -- 查询 ------------------------------------------------------------- #
    def query(self, number: str | None = None) -> None:
        if isinstance(number, str) and number:
            self.number_var.set(number)
        companies, entries, ranges, empty = self.app.district_bus
        text = self.number_var.get().strip()
        if not companies:
            self.answer_var.set("还没有区属公交数据。")
            self.detail_var.set(f"把数据填进 {DISTRICT_BUS_FILENAME} 之后，这里就能按自编或车牌查车。")
            return
        if not text:
            self.answer_var.set("请输入自编或车牌。")
            self.detail_var.set("")
            return
        # 自编和车牌都走这一个入口（像车牌就当车牌查）
        result = district_bus_lookup_any(text, companies, ranges, empty, entries)
        when = " · ".join(
            part for part in (
                f"{result.year} 年" if result.year else "",
                f"第 {result.batch} 批" if result.batch else "",
                f"当年第 {result.serial} 辆" if result.serial else "",
            ) if part
        )
        if result.status == "已收录":
            self.answer_var.set(f"{result.query}　{result.model or '—'}　（{result.status}）")
        else:
            self.answer_var.set(f"{result.query}　{result.status}")
        parts = [
            part for part in (
                result.company,
                when,
                f"线路 {result.line}" if result.line else "",
                f"号段 {result.span}" if result.span else "",
                result.note,
            ) if part
        ]
        self.detail_var.set("　｜　".join(parts))
        # 表格里定位这一段，方便看前后号段（车牌查询没有号段）
        if result.span:
            for iid in self.table.tree.get_children():
                if self.table.tree.item(iid, "values")[1] == result.span:
                    self.table.tree.selection_set(iid)
                    self.table.tree.see(iid)
                    break

    def toggle_explain(self) -> None:
        """展开 / 收起「各区县现状 + 其他区怎么加」那段说明（平时不占地方）。"""
        if self.explain_shown:
            self.reserved_label.pack_forget()
            self.tools_button.configure(text="数据说明 ▾")
        else:
            self.reserved_label.pack(fill="x", pady=(0, self.theme.px(8)), after=self.detail_label)
            self.tools_button.configure(text="收起说明 ▴")
        self.explain_shown = not self.explain_shown

    def reload(self) -> None:
        self.app.reload_district_bus()
        self.refresh()
        self.query()
        self.app.status_var.set(f"已重新读取 {DISTRICT_BUS_FILENAME}")

    def refresh(self) -> None:
        companies, _entries, ranges, empty = self.app.district_bus
        models = ["全部车型"] + sorted({item.model for item in ranges if item.model})
        chosen = self.model_var.get() if hasattr(self, "model_var") else "全部车型"
        self.model_combo.configure(values=models)
        if chosen not in models:
            chosen = "全部车型"
            self.model_var.set(chosen)

        shown = [item for item in ranges if chosen in ("全部车型", item.model)]
        rows = []
        for item in shown:
            described = next((c.describe(f"{item.start:08d}") for c in companies if c.matches_number(f"{item.start:08d}")), {})
            when = "/".join(part for part in (
                f"{described.get('year', '')} 年" if described.get("year") else "",
                f"第 {described.get('batch', '')} 批" if described.get("batch") else "",
            ) if part)
            rows.append((
                item.company,
                item.label,
                item.real_count,
                when,
                item.model,
                f"尾号 {'/'.join(item.skip_tail)} 未编" if item.skip_tail else "",
            ))
        self.table.set_rows(rows, keep_selection=False)
        total = sum(item.count for item in shown)
        real = sum(item.real_count for item in shown)
        skipped = total - real
        tail_note = f"（含 {skipped} 个整批未编的编号）" if skipped else ""
        empty_note = f" · 逐条空号 {len(empty)} 个" if empty else ""
        self.count_var.set(f"{len(rows)} 段 · {real} 辆{tail_note}{empty_note}")
        all_skipped = sum(item.skipped_count for item in ranges)
        self.summary_var.set(
            f"{len(companies)} 家公司 · {len(ranges)} 段 · 覆盖 {sum(item.count for item in ranges)} 个编号"
            f" · 实车 {sum(item.real_count for item in ranges)} 辆"
            + (f"（{all_skipped} 个编号整批未编）" if all_skipped else "")
            + (f" · 空号 {len(empty)} 个" if empty else "")
        )
        # ② 车牌记录表 + 汇总（车型 → 车牌数）
        plate_rows = [
            (item.label, item.model, item.line, item.company, item.date, item.note)
            for item in sorted(_entries, key=lambda row: (row.date, row.label))
        ]
        self.entry_table.set_rows(plate_rows, keep_selection=False)
        self.entry_table.empty.update(
            "还没有车牌记录：点「＋ 记一辆」把坐过的车牌和车型记下来（也可以批量粘贴）" if not plate_rows else None,
            has_rows=bool(plate_rows),
        )
        summary = district_bus_entry_summary(_entries)
        conflicts = district_bus_entry_conflicts(_entries)
        if summary:
            top = "、".join(f"{model} {count} 辆" for model, count in summary[:5])
            self.plate_count_var.set(f"{len(_entries)} 条记录 · {len(summary)} 种车型：{top}")
        else:
            self.plate_count_var.set("还没有车牌记录")
        if conflicts:
            self.count_var.set(self.count_var.get() + f"　｜　⚠ {len(conflicts)} 个车牌记了不同车型（看②）")
        # 各区县现状 + 其他区（只有车牌）的预留说明
        districts = district_bus_districts(companies, ranges, _entries)
        covered = "；".join(f"{name}{'（' + state + '）' if state else ''}" for name, state, _count in districts if _count)
        wait = [name for name, _state, count in districts if not count]
        pieces: list[str] = []
        if covered:
            pieces.append(f"已有数据：{covered}")
        if wait:
            pieces.append(f"还空着：{'、'.join(wait)}")
        if any(company.all_plate_prefixes for company in companies) or not any(item.plate for item in _entries):
            pieces.append(
                "其他区目前基本只有车牌号、没有自编：在 companies 里加一条（number_kind 填「车牌」、"
                "plate_prefixes 填如 [\"京B\"]），再把车记进 entries 的 plate 字段，这个查询框直接输车牌就能查"
            )
        self.reserved_var.set("　｜　".join(pieces))

    def create_template(self) -> None:
        path = district_bus_path(application_dir())
        if path.exists():
            messagebox.showinfo("区属公交数据", f"{DISTRICT_BUS_FILENAME} 已经存在了，没有覆盖它。", parent=self)
            self.refresh()
            return
        try:
            save_district_bus_template(application_dir())
        except ProfileError as error:
            messagebox.showerror("建立失败", str(error), parent=self)
            return
        self.app.reload_district_bus()
        self.refresh()
        self.app.status_var.set(f"已建立 {DISTRICT_BUS_FILENAME}（里面有一家示例公司，按实际情况改）")

    def open_folder(self) -> None:
        try:
            os.startfile(application_dir())  # noqa: S606 - 打开资源管理器
        except OSError as error:
            messagebox.showerror("打开失败", str(error), parent=self)


class MyRecordsWindow(WindowMixin, tk.Toplevel):
    """我的记录：读一次外面的运转记录，把三件事一起给你。

    以前分成「数据对账」（只读看一眼）和「线路管理」（手工/导入）两个窗口，
    读的是同一份文件——看完差异没有下一步动作，所以显得鸡肋。现在合并：
    ① 线路数据（可一键导入，笔误与外地记录也全都留着）
    ② 对账差异（写法不一致等，可导出）
    ③ 号段表缺的编号（直接去批量补录）
    """

    CATEGORIES = (
        ("matched", "编号+型号一致", ("车型", "条数")),
        ("mismatch", "型号不一致（按值得查的程度排）", ("差异性质", "号段表 → 记录里", "条数", "例子（线路 自编号）")),
        ("number_gap", "号段表缺的编号", ("自编号", "记录里的车型", "条数")),
        ("model_gap", "没填编号的型号", ("车型", "条数")),
        ("metro_missing", "地铁数据里没有的型号", ("车型", "条数")),
        ("out_of_town", "外地记录（已排除）", ("车型", "条数", "线路")),
    )

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "我的记录", 1000, 660)
        pad = self.theme.px
        self.path: Path | None = None
        self.report: dict[str, object] | None = None
        self.links: list[LineLink] = []

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)

        self.dialog_header(
            frame,
            "我的记录（导入线路 / 对账）",
            "选一个外面的运转记录（例如 D:\\Bus Document\\乘车运转记录表.xlsx）：程序读出每一张年度表，"
            "一次给你三样——能导入的线路数据、跟号段表的对账差异、以及号段表里缺的编号。"
            "只读你的文件，不会改它。",
            wrap=880,
        )

        picker = ttk.Frame(frame)
        picker.pack(fill="x", pady=(pad(10), pad(6)))
        ttk.Button(picker, text="选择运转记录…", style="Primary.TButton", command=self.choose_file).pack(side="left")
        ttk.Button(picker, text="重新读取", style="Secondary.TButton", command=self.run).pack(side="left", padx=pad(8))
        self.file_var = tk.StringVar(value="还没有选择文件")
        ttk.Label(picker, textvariable=self.file_var, style="Muted.TLabel", wraplength=pad(600)).pack(side="left", padx=pad(10))

        self.summary_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel", wraplength=pad(900), justify="left").pack(
            fill="x", pady=(0, pad(8))
        )

        actions = ttk.Frame(frame)
        actions.pack(side="bottom", fill="x", pady=(pad(10), 0))
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")

        self.notebook = ttk.Notebook(frame)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.add(self._build_lines_tab(), text="① 线路数据")
        self.notebook.add(self._build_reconcile_tab(), text="② 对账差异")
        self.notebook.add(self._build_missing_tab(), text="③ 号段表缺的编号")

    # -- 三个页签 --------------------------------------------------------- #
    def _build_lines_tab(self) -> ttk.Frame:
        pad = self.theme.px
        tab = ttk.Frame(self.notebook, padding=(pad(10), pad(10)))
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        self.line_table = DataTable(
            tab,
            self.theme,
            (
                TableColumn("line", "线路", 150, "w"),
                TableColumn("model", "车型", 190, "w"),
                TableColumn("numbers", "自编号", 70, "e", numeric=True),
                TableColumn("span", "记录时间", 140, "w"),
                TableColumn("note", "备注", 260, "w"),
            ),
            empty_text="先选一个运转记录文件",
        )
        self.line_table.card.grid(row=0, column=0, sticky="nsew")
        self.line_table.tree.configure(height=8)  # 压矮一点：三个页签 + 说明也要放得下
        bar = ttk.Frame(tab)
        bar.grid(row=1, column=0, sticky="ew", pady=(pad(8), 0))
        ttk.Button(bar, text="导入 / 更新到线路数据", style="Primary.TButton", command=self.import_lines).pack(side="left")
        ttk.Button(bar, text="手工编辑线路", style="Secondary.TButton", command=lambda: LineManagerWindow(self.app)).pack(
            side="left", padx=pad(8)
        )
        ttk.Button(bar, text="⇩ 导出 CSV", style="Ghost.TButton",
                   command=lambda: self.line_table.export(self, "我的线路记录.csv")).pack(side="left")
        # 检查类、删除类各收一个菜单：底下一排 8 个按钮谁也分不清主次
        compact_menu(
            bar,
            "检查 ▾",
            [
                ("哪些网上查的该刷新了", self.check_online_stale),
                ("用线路目录校验线路号", self.validate_line_numbers),
            ],
        ).pack(side="left", padx=pad(8))
        compact_menu(
            bar,
            "删除 ▾",
            [
                ("✕ 删除导入的（保留手工加的）", self.delete_imported),
                ("✕ 删除网上查的（保留自己记的）", self.delete_online),
                ("✕ 清空全部线路记录", self.delete_all_lines),
            ],
        ).pack(side="right")
        ttk.Label(bar, text="笔误和外地记录也会导入并标注在备注里（以后可用来查外地号/本地区属公交）",
                  style="CardMuted.TLabel").pack(side="left", padx=pad(10))
        return tab

    def _build_reconcile_tab(self) -> ttk.Frame:
        pad = self.theme.px
        tab = ttk.Frame(self.notebook, padding=(pad(10), pad(10)))
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        self.category_titles = {title: key for key, title, _columns in self.CATEGORIES}
        self.category_var = tk.StringVar(value=self.CATEGORIES[0][1])

        def build_toolbar(bar: ttk.Frame) -> None:
            self.category_combo = ttk.Combobox(
                bar,
                textvariable=self.category_var,
                values=[title for _key, title, _columns in self.CATEGORIES],
                state="readonly",
                width=26,
            )
            self.category_combo.pack(side="left")
            self.category_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())
            self.count_var = tk.StringVar(value="")
            ttk.Label(bar, textvariable=self.count_var, style="CardMuted.TLabel").pack(side="left", padx=pad(10))

        self.table = DataTable(
            tab,
            self.theme,
            self._columns_for(self.CATEGORIES[0][0]),
            toolbar=build_toolbar,
            sort_on_load=False,  # 保持报表自己的顺序（型号不一致那一类按"值得查"排好了）
            empty_text="先选一个运转记录文件，再点「重新读取」",
        )
        self.table.card.grid(row=0, column=0, sticky="nsew")
        self.table.tree.configure(height=8)  # 与另外两个页签一样矮，免得它把整个窗口撑高
        self.tree = self.table.tree
        self.empty_state = self.table.empty
        ttk.Button(tab, text="⇩ 导出这一类的 CSV", style="Secondary.TButton", command=self.export_rows).grid(
            row=1, column=0, sticky="w", pady=(pad(8), 0)
        )
        return tab

    def _build_missing_tab(self) -> ttk.Frame:
        pad = self.theme.px
        tab = ttk.Frame(self.notebook, padding=(pad(10), pad(10)))
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        self.missing_table = DataTable(
            tab,
            self.theme,
            (
                TableColumn("number", "自编号", 110, "w"),
                TableColumn("model", "记录里的车型", 200, "w"),
                TableColumn("count", "条数", 70, "e", numeric=True),
            ),
            empty_text="没有号段表里查不到的编号",
        )
        self.missing_table.card.grid(row=0, column=0, sticky="nsew")
        self.missing_table.tree.configure(height=8)
        bar = ttk.Frame(tab)
        bar.grid(row=1, column=0, sticky="ew", pady=(pad(8), 0))
        ttk.Button(bar, text="去批量补录这些编号", style="Primary.TButton", command=self.open_bulk_range).pack(side="left")
        ttk.Button(bar, text="管理号段", style="Secondary.TButton", command=lambda: RangeManagerWindow(self.app)).pack(
            side="left", padx=pad(8)
        )
        ttk.Label(bar, text="这些编号在你的记录里出现过，但号段表里查不到——补上之后查车就能出车型",
                  style="CardMuted.TLabel").pack(side="left", padx=pad(10))
        return tab

    @staticmethod
    def _columns_for(key: str) -> tuple[TableColumn, ...]:
        titles = next((columns for item_key, _title, columns in MyRecordsWindow.CATEGORIES if item_key == key), ("对象",))
        widths = (360, 110, 300) if len(titles) >= 3 else (360, 120)
        return tuple(
            TableColumn(f"c{index}", title, widths[min(index, len(widths) - 1)], "e" if title in ("条数", "编号数") else "w", numeric=title == "条数")
            for index, title in enumerate(titles)
        )

    # -- 读取与对账 -------------------------------------------------------- #
    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="选择运转记录（Excel）",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.path = Path(path)
            self.file_var.set(self.path.name)
            self.run()

    def run(self) -> None:
        if self.path is None or not self.path.is_file():
            messagebox.showinfo("我的记录", "请先选择一个运转记录文件（.xlsx）。", parent=self)
            return
        try:
            records = load_ride_records(self.path)
        except DataError as error:
            messagebox.showerror("读不了", str(error), parent=self)
            return
        if not records:
            messagebox.showinfo("我的记录", "这个文件里没有找到「线路 + 车型」这样的表头，可能不是运转记录表。", parent=self)
            return
        known = set(model_names(self.app.ranges))
        self.report = reconcile_rides(records, known, {item.model for item in self.app.metro_ranges}, self.app.ranges)
        # 线路数据：**全都留着**（笔误与外地记录也留着，以后可以拿来做外地号/本地区属公交查询）
        self.links = build_links_from_records(records, source="import")
        totals = self.report["totals"]
        outside = len([link for link in self.links if link.model and link.model not in known])
        self.summary_var.set(
            f"共 {totals['all']} 条记录 → 线路记录 {len(self.links)} 条"
            f"（{len(all_line_names(self.links))} 条线路 · {len({link.model for link in self.links if link.model})} 种车型，"
            f"其中 {outside} 条的车型不在号段表里，会标注在备注里）　｜　"
            f"对账：一致 {sum(int(count) for _name, count in self.report['matched'])} 条"
            f" · 型号不一致 {sum(int(row[2]) for row in self.report['mismatch'])} 条"
            f" · 号段表缺 {self.report['number_gap_count']} 个编号"
            f" · 没填编号的型号 {len(self.report['model_gap'])} 种"
        )
        self.refresh_lines()
        self.refresh_table()
        self.refresh_missing()

    def refresh_lines(self) -> None:
        known = set(model_names(self.app.ranges))
        rows = []
        for link in self.links:
            note = "" if (link.model in known or not link.model) else "车型不在号段表里（可能是外地/区属公交/笔误）"
            rows.append((link.line, link.model or "—", len(link.numbers), line_display(link), note))
        self.line_table.set_rows(rows, keep_selection=False)
        self.line_table.empty.update("先选一个运转记录文件" if not rows else None, has_rows=bool(rows))

    def refresh_missing(self) -> None:
        rows = [tuple(row) for row in (self.report or {}).get("number_gap", [])]  # type: ignore[union-attr]
        self.missing_table.set_rows(rows, keep_selection=False)
        self.missing_table.empty.update(
            "号段表里没有查不到的编号（对账结果里这一项是空的）" if not rows else None, has_rows=bool(rows)
        )

    def import_lines(self) -> None:
        """把整理好的线路记录并进 线路.json（保留全部，包括笔误与外地）。"""
        if not self.links:
            messagebox.showinfo("我的记录", "先选一个运转记录文件并读取。", parent=self)
            return
        plan = merge_lines(self.app.lines, self.links)
        if not plan.changed:
            messagebox.showinfo("我的记录", f"线路数据已经是最新的（{len(self.app.lines)} 条），没有需要更新的。", parent=self)
            return
        if not messagebox.askyesno(
            "导入线路数据",
            f"将把 {len(self.links)} 条线路记录并进现有 {len(self.app.lines)} 条：\n"
            f"新增 {len(plan.added)} · 合并更新 {len(plan.updated)}\n\n"
            "同一「线路 + 车型」会合并（时间取并集、自编号取并集）；笔误与外地记录也一起留着。\n"
            "写入前会自动备份 线路.json。确定继续吗？",
            parent=self,
        ):
            return
        merged = apply_line_plan(self.app.lines, plan)
        try:
            save_lines(application_dir(), merged)
        except ProfileError as error:
            messagebox.showerror("导入失败", str(error), parent=self)
            return
        self.app.lines = merged
        self.app.notify_data_changed({"lines"})
        self.app.status_var.set(f"已导入线路记录：现在共 {len(merged)} 条（查车时会显示「我记过」）")
        messagebox.showinfo(
            "导入完成",
            f"线路数据现在共 {len(merged)} 条。\n查任一自编号时，结果里会多一行「我记过：…」。",
            parent=self,
        )

    def validate_line_numbers(self) -> None:
        """用本地线路目录校验你线路数据里的线路号（打错号 / 服务端没有的）。"""
        index = self._local_catalog()
        if not index:
            messagebox.showinfo(
                "校验线路号",
                "还没有本地线路目录。先打开「线路资料（联网）」，点「更多 ▾」→「更新线路目录」"
                "（需要联网），之后校验线路号就不用联网了。",
                parent=self,
            )
            return
        result = validate_line_numbers(self.app.lines, index)
        unknown = result["unknown"]
        if not unknown:
            messagebox.showinfo("校验线路号", f"{len(self.app.lines)} 条线路记录里的线路号，本地目录都能对上。", parent=self)
            return
        lines = []
        for name, model, count in unknown[:15]:
            guesses = result["suggest"].get(name) or []
            lines.append(
                f"　{name}　{model or '（车型未知）'}　{count} 辆"
                + (f"　你是不是想记：{'、'.join(guesses)}" if guesses else "")
            )
        messagebox.showinfo(
            "这些线路号本地目录里没有",
            f"共 {len(unknown)} 个线路号在第三方数据里找不到：\n\n"
            + "\n".join(lines)
            + (f"\n　…还有 {len(unknown) - 15} 个" if len(unknown) > 15 else "")
            + "\n\n可能的原因：① 号记错了（看后面的「你是不是想记」）② 这条线不在第三方数据里"
            "（例如你记的外地线路、已改号的线路）。\n"
            "要改：打开「线路管理」直接编辑；这只是提示，不会动你的数据。",
            parent=self,
        )

    def _local_catalog(self) -> dict:
        """直接从 线路目录.json 读一份（不依赖别的窗口）。"""
        payload = load_route_list(application_dir())
        return route_list_numbers(payload.get("routes") or {})

    def check_online_stale(self) -> None:
        """列出"网上查的、并且已经很久没更新"的线路记录。

        为什么要这个：服务端的数据一直在变（车调走了、线路改了），
        网上查来的记录放着不管就会慢慢变成错的；这里把超过 30 天的挑出来，
        让用户知道该重新查一遍（或者干脆删掉）。
        """
        stale = online_stale(self.app.lines, days=30)
        if not stale:
            messagebox.showinfo(
                "网上查的记录", "没有需要刷新的：网上查来的记录都在 30 天内更新过（或者你还没导入过联网数据）。", parent=self
            )
            return
        lines = "\n".join(
            f"　{link.line}　{link.model or '（车型未知）'}　{line_display(link)}　共 {len(link.numbers)} 辆"
            for link in stale[:12]
        )
        messagebox.showinfo(
            "网上查的记录该刷新了",
            f"有 {len(stale)} 条网上查来的记录超过 30 天没更新：\n\n{lines}"
            f"{chr(10) + '　…' if len(stale) > 12 else ''}\n\n"
            "想刷新：打开「线路资料（联网）」查这条线，再点「写入线路数据」（会合并，不会重复）。\n"
            "不想留：点右边「✕ 删除网上查的」一键清掉（你自己记的不会动）。",
            parent=self,
        )

    def delete_imported(self) -> None:
        """一键删掉「从外部表格导进来」的线路记录（手工加的不动）。"""
        self._delete_by_source("import", "导入来的")

    def delete_online(self) -> None:
        """一键删掉「网上查的」线路记录（手工记的、表格导入的都不动）。"""
        self._delete_by_source("online", "网上查的")

    def _delete_by_source(self, source: str, label: str) -> None:
        targets = [link for link in self.app.lines if link.source == source]
        if not targets:
            messagebox.showinfo("我的记录", f"线路数据里没有{label}记录（其它来源不受影响）。", parent=self)
            return
        if not messagebox.askyesno(
            f"删除{label}线路记录",
            f"将删除 {len(targets)} 条{label}线路记录，其它来源的会保留。\n\n"
            f"删完剩 {len(self.app.lines) - len(targets)} 条。写入前自动备份 线路.json。\n\n确定继续吗？",
            parent=self,
        ):
            return
        self._write_lines(delete_by_source(self.app.lines, source), f"已删除{label}线路记录")

    def delete_all_lines(self) -> None:
        """一键清空全部线路数据。"""
        if not self.app.lines:
            messagebox.showinfo("我的记录", "线路数据本来就是空的。", parent=self)
            return
        if not messagebox.askyesno(
            "清空全部线路记录",
            f"将清空全部 {len(self.app.lines)} 条线路记录（手工加的也会删）。\n\n"
            "写入前自动备份 线路.json，删完还能用「备份与恢复」找回来。\n\n确定继续吗？",
            parent=self,
        ):
            return
        self._write_lines([], "已清空线路记录")

    def _write_lines(self, links: list[LineLink], message: str) -> None:
        try:
            save_lines(application_dir(), links)
        except ProfileError as error:
            messagebox.showerror("删除失败", str(error), parent=self)
            return
        self.app.lines = list(links)
        self.app.notify_data_changed({"lines"})
        self.app.status_var.set(f"{message}：现在共 {len(links)} 条")

    def open_bulk_range(self) -> None:
        """把「号段表缺的编号」直接带去批量补录。"""
        rows = (self.report or {}).get("number_gap", [])  # type: ignore[union-attr]
        numbers = [int(row[0]) for row in rows if str(row[0]).isdigit()]
        if not numbers:
            messagebox.showinfo("我的记录", "没有号段表里查不到的编号。", parent=self)
            return
        counts: Counter = Counter(str(row[1]) for row in rows)
        default_model = counts.most_common(1)[0][0] if counts else ""
        BulkRangeWindow(self, group_missing_numbers(numbers), default_model)

    def refresh_table(self) -> None:
        key = self.category_titles.get(self.category_var.get(), self.CATEGORIES[0][0])
        self.table.set_columns(self._columns_for(key))
        rows: list[tuple] = []
        if self.report is not None:
            if key in ("matched", "model_gap", "metro_missing"):
                rows = [tuple(row) for row in self.report[key]]  # type: ignore[index]
            elif key == "mismatch":
                rows = [tuple(row) for row in self.report["mismatch"]]  # type: ignore[index]
            elif key == "number_gap":
                rows = [(row[0], row[1], row[2]) for row in self.report["number_gap"]]  # type: ignore[index]
            elif key == "out_of_town":
                rows = [tuple(row) for row in self.report["out_of_town"]]  # type: ignore[index]
        self.table.set_rows(rows, keep_selection=False)
        self.count_var.set(f"{len(rows)} 行")
        messages = {
            "matched": "这些编号在号段表里查到同样的车型，说明两边数据是一致的",
            "mismatch": "同一个自编号，号段表和记录里写的型号不一样。判据是「同一台车两个说法」这个硬事实，"
                        "不比名字像不像——同一款车换批次（-1 / -5）或换能源（BTB / BTD）名字本来就像。"
                        "按值得查的程度从上往下排：完全不搭的最该看，纯写法差异（多了厂内代号、括号）可以不管。",
            "number_gap": "这些自编号在号段表里查不到，是真正缺的号段；记录里的车型可以用来补录",
            "model_gap": "这些记录没填自编号，所以只能用型号比对；型号也不在号段表里",
            "metro_missing": "记录里有、但地铁数据里没有的型号",
            "out_of_town": "线路名看着是外地的，已从公交对账里排除",
        }
        self.empty_state.update(messages.get(key, "") if not rows else None, has_rows=bool(rows))

    def export_rows(self) -> None:
        if self.report is None:
            messagebox.showinfo("导出", "请先对账。", parent=self)
            return
        if not self.table.rows():
            messagebox.showinfo("导出", "当前这一类没有可导出的内容。", parent=self)
            return
        path = self.table.export(self, "数据对账.csv")
        if path:
            self.app.status_var.set(f"已导出对账结果到 {path}")


# --------------------------------------------------------------------------- #
# 拍车日志
# --------------------------------------------------------------------------- #


class SpotLogWindow(WindowMixin, tk.Toplevel):
    """拍车记录：看、筛、导出。"""

    COLUMNS = (
        ("time", "时间", 155, "center"),
        ("number", "自编号", 80, "center"),
        ("model", "车型", 210, "w"),
        ("note", "备注", 200, "w"),
        ("photo", "照片", 60, "center"),
    )

    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "拍车日志", 860, 620)
        pad = self.theme.px
        self.entries: list[SpotEntry] = []
        self.shown: list[SpotEntry] = []

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text="拍车日志", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.count_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.count_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")

        filter_row = ttk.Frame(frame)
        filter_row.grid(row=1, column=0, sticky="ew", pady=(pad(8), pad(10)))
        filter_row.columnconfigure(1, weight=1)
        ttk.Label(filter_row, text="搜索").grid(row=0, column=0, padx=(0, pad(6)))
        self.filter_var = tk.StringVar()
        search = ttk.Entry(filter_row, textvariable=self.filter_var, style="Search.TEntry")
        search.grid(row=0, column=1, sticky="ew")
        attach_placeholder(search, "自编号 / 车型 / 备注", self.theme)
        ttk.Label(filter_row, text="车型").grid(row=0, column=2, padx=(pad(10), pad(6)))
        self.model_var = tk.StringVar(value=ANY_TAG)
        self.model_combo = ttk.Combobox(filter_row, textvariable=self.model_var, values=[ANY_TAG], state="readonly", width=18)
        self.model_combo.grid(row=0, column=3)
        ttk.Label(filter_row, text="日期").grid(row=0, column=4, padx=(pad(10), pad(6)))
        self.day_var = tk.StringVar()
        day = ttk.Entry(filter_row, textvariable=self.day_var, width=12, justify="center")
        day.grid(row=0, column=5)
        attach_placeholder(day, "2026-09-30", self.theme)
        ttk.Button(filter_row, text="⟳ 重置", style="Ghost.TButton", command=self.reset_filters).grid(row=0, column=6, padx=(pad(8), 0))
        self._debounced_refresh = debounce(self, "spot-search", 150, self.refresh)
        for variable in (self.filter_var, self.model_var, self.day_var):
            variable.trace_add("write", self._debounced_refresh)

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor)
                for identifier, title, width, anchor in self.COLUMNS
            ),
            selectmode="extended",
            empty_text="还没有拍车记录\n查询到车型后点主窗口的「记一笔」",
            on_open=lambda: self.edit_selected(),
            on_delete=self.delete_selected,
            on_search=lambda: (self.filter_var.set(""), search.focus_set()),
        )
        self.table.card.grid(row=2, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.empty_state = self.table.empty
        self.tree.bind("<Return>", self.edit_selected)

        self.summary_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel", wraplength=pad(780)).grid(row=3, column=0, sticky="w", pady=(pad(8), 0))

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(actions, text="编辑选中", style="Primary.TButton", command=self.edit_selected).pack(side="left")
        ttk.Button(actions, text="查看车型详情", style="Secondary.TButton", command=self.open_selected).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="看照片", style="Secondary.TButton", command=self.open_photo).pack(side="left")
        ttk.Button(actions, text="✕ 删除选中", style="Danger.TButton", command=self.delete_selected).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="⇩ 导出 CSV", style="Secondary.TButton", command=self.export_rows).pack(side="right")
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))

        self.reload()

    def reload(self) -> None:
        try:
            self.entries = load_spots(application_dir())
        except ProfileError as error:
            messagebox.showerror("读取失败", str(error), parent=self)
            self.entries = []
        self.model_combo.configure(values=[ANY_TAG] + spot_models(self.entries))
        if self.model_var.get() not in self.model_combo.cget("values"):
            self.model_var.set(ANY_TAG)
        self.refresh()

    def reset_filters(self) -> None:
        self.filter_var.set("")
        self.model_var.set(ANY_TAG)
        self.day_var.set("")

    def refresh(self, *_args: object) -> None:
        self.shown = filter_spots(
            self.entries,
            self.filter_var.get(),
            None if self.model_var.get() == ANY_TAG else self.model_var.get(),
            self.day_var.get().strip(),
        )
        rows = [
            (entry.time, entry.number, entry.model, entry.note, "有" if entry.photo else "—")
            for entry in self.shown
        ]
        # iid 是 self.shown 里的原始下标：编辑/删除/看照片都靠它取记录，排序后也不能变
        self.table.set_rows(rows, keep_selection=True, iid=lambda position, row: str(position))
        total = len(self.entries)
        self.count_var.set(f"共 {total} 条记录" if len(self.shown) == total else f"匹配 {len(self.shown)} / {total} 条")
        summary = spot_summary(self.entries)
        top = summary.get("top") or []
        extra = "　拍到最多：" + "、".join(f"{name} ×{count}" for name, count in list(top)[:3]) if top else ""
        self.summary_var.set(spot_summary_text(summary) + extra)
        self.empty_state.update(
            "没有匹配的记录\n试试清空搜索或把筛选改回「全部」" if not self.shown else None,
            has_rows=bool(self.shown),
        )

    def open_selected(self, _event: tk.Event | None = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        entry = self.shown[int(selection[0])]
        if entry.model in set(model_names(self.app.ranges)):
            ModelDetailWindow(self, entry.model)

    def edit_selected(self, _event: tk.Event | None = None) -> None:
        """双击一条 → 编辑这条记录（备注、照片）。"""
        selection = self.tree.selection()
        if not selection:
            return
        entry = self.shown[int(selection[0])]
        try:
            index = self.entries.index(entry)
        except ValueError:
            return
        SpotDialog(
            self,
            index=index,
            number=entry.number,
            model=entry.model,
            note=entry.note,
            photo=entry.photo,
            cover=entry.image,
            time_text=entry.time,
        )

    def open_photo(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("看照片", "请先选中一条记录。", parent=self)
            return
        entry = self.shown[int(selection[0])]
        if not entry.photo:
            messagebox.showinfo("看照片", "这条记录还没有照片。", parent=self)
            return
        path = stored_photo_path(application_dir(), entry.photo)
        if not path.is_file():
            messagebox.showwarning("看照片", f"照片文件找不到了：{entry.photo}", parent=self)
            return
        try:
            os.startfile(path)  # noqa: S606 - 交给系统看图程序
        except OSError as error:
            messagebox.showerror("打不开", str(error), parent=self)

    def delete_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("删除记录", "请先选中要删除的记录。", parent=self)
            return
        targets = {self.shown[int(item)] for item in selection}
        if not messagebox.askyesno("删除记录", f"确定删除选中的 {len(targets)} 条记录吗？", parent=self):
            return
        remaining = [entry for entry in self.entries if entry not in targets]
        try:
            save_spots(application_dir(), remaining)
        except ProfileError as error:
            messagebox.showerror("删除失败", str(error), parent=self)
            return
        self.reload()

    def export_rows(self) -> None:
        if not tree_rows(self.tree):
            messagebox.showinfo("导出", "当前没有可导出的记录。", parent=self)
            return
        path = export_tree_rows(self, self.tree, "拍车日志.csv")
        if path:
            self.app.status_var.set(f"已导出 {len(self.shown)} 条拍车记录到 {path}")


# --------------------------------------------------------------------------- #
# Tag rules manager
# --------------------------------------------------------------------------- #


class TagConfirmWindow(WindowMixin, tk.Toplevel):
    """燃料待确认：把「有依据的建议」列出来，勾选后一次性写入（写前自动备份）。

    优先用查证过的公开来源（带链接），其次用「你已经标过的同后缀车型」归纳出的规律；
    两者都没有的，按后缀分组让你一次给一组定一个值——不猜。
    """

    COLUMNS = (
        ("pick", "选", 34, "center"),
        ("model", "车型", 180, "w"),
        ("source", "依据", 116, "w"),
        ("value", "写入值", 90, "center"),
        ("evidence", "依据详情", 520, "w"),
    )

    def __init__(self, app: BusLookupApp, field: str = "fuel") -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.field = field
        self.overrides = app.tag_overrides
        self.setup_dialog(app, "燃料待确认", 1040, 620)
        pad = self.theme.px

        targets = [name for name in model_names(app.ranges) if self._current(name) == UNKNOWN_TAG]
        suggested, unsure = suggest_tags_for_models(self.overrides, field, targets)
        self.suggestions = {item.model: item for item in suggested}
        self.unsure = unsure
        self.checked: set[str] = {item.model for item in applicable_suggestions(suggested)}
        self.manual: dict[str, str] = {}  # 按后缀分组时人工指定的值
        self.rows: list[str] = sorted(targets, key=lambda name: name.casefold())
        self.groups = group_by_suffix(self.rows, None, field)

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        self.dialog_header(
            frame,
            "燃料待确认",
            "有依据的已经替你勾好（依据栏会写明是「查证来源」还是「从你的标注学到」）。"
            "没有依据的在下面按后缀分组，一次给一组定一个值。写入前会自动备份 车型标签.json，"
            "写错了可以用数据体检里的「撤销上次标签修改」回退。",
            wrap=920,
        )

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(pad(10), pad(6)))
        ttk.Label(toolbar, text="按后缀分组").pack(side="left", padx=(0, pad(6)))
        self.group_var = tk.StringVar(value=self.groups[0][0] if self.groups else "")
        self.group_combo = ttk.Combobox(
            toolbar,
            textvariable=self.group_var,
            values=[f"{suffix}（{len(models)} 台）" for suffix, models in self.groups],
            state="readonly",
            width=18,
        )
        self.group_combo.pack(side="left")
        ttk.Label(toolbar, text="设成").pack(side="left", padx=(pad(8), pad(6)))
        self.value_var = tk.StringVar()
        self.value_combo = ttk.Combobox(toolbar, textvariable=self.value_var, values=self._value_options(), width=12)
        self.value_combo.pack(side="left")
        ttk.Button(toolbar, text="套用到这组", style="Secondary.TButton", command=self.apply_group).pack(side="left", padx=pad(8))
        ttk.Button(toolbar, text="全选有依据的", style="Ghost.TButton", command=self.check_all_suggested).pack(side="left")
        ttk.Button(toolbar, text="全不选", style="Ghost.TButton", command=self.check_none).pack(side="left", padx=pad(6))

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor)
                for identifier, title, width, anchor in self.COLUMNS
            ),
            empty_text="没有需要确认的车型",
        )
        self.table.card.pack(fill="both", expand=True)
        self.tree = self.table.tree
        self.tree.bind("<Button-1>", self._toggle_from_click)

        self.summary_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel", wraplength=pad(940)).pack(
            fill="x", pady=(pad(8), 0)
        )
        self.detail_var = tk.StringVar(value="选中一行看完整依据（查证来源会带链接）")
        ttk.Label(frame, textvariable=self.detail_var, style="Muted.TLabel", wraplength=pad(940), justify="left").pack(
            fill="x", pady=(pad(4), 0)
        )
        self.tree.bind("<<TreeviewSelect>>", self._show_detail)

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(pad(8), 0))
        ttk.Button(actions, text="应用勾选项（自动备份）", style="Primary.TButton", command=self.apply_checked).pack(side="left")
        ttk.Button(actions, text="⇩ 导出待确认表 CSV", style="Secondary.TButton", command=self.export_rows).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")

        self.refresh()

    # -- 数据 ------------------------------------------------------------- #
    def _current(self, model: str) -> str:
        return model_tags(model, self.overrides).fuel

    def _value_options(self) -> list[str]:
        values = {tag.fuel for tag in (model_tags(name, self.overrides) for name in model_names(self.app.ranges))}
        values.discard(UNKNOWN_TAG)
        return sorted(values)

    def _effective(self, model: str) -> tuple[str, str, str]:
        """返回 (依据, 要写入的值, 详情)。"""
        if model in self.manual:
            return "你按后缀指定", self.manual[model], f"后缀分组里你选了「{self.manual[model]}」"
        suggestion = self.suggestions.get(model)
        if suggestion is not None:
            return suggestion.source, suggestion.value, suggestion.evidence()
        return "需人工判断", "", "公开来源查不到、你的标注里也没有同后缀的例子"

    def refresh(self) -> None:
        rows = []
        for model in self.rows:
            source, value, evidence = self._effective(model)
            pick = "☑" if model in self.checked else "☐"
            rows.append((pick, model, source, value or "—", self._short_evidence(model, source, evidence)))
        self.table.set_rows(rows, keep_selection=False, iid=lambda position, row: row[1])
        ready = sum(1 for model in self.rows if model in self.checked and self._effective(model)[1])
        self.summary_var.set(
            f"共 {len(self.rows)} 个未标注车型 · 有依据 {len(self.suggestions)} 个 · "
            f"需人工 {len(self.unsure)} 个 · 当前勾选可写入 {ready} 个"
        )

    def _short_evidence(self, model: str, source: str, evidence: str) -> str:
        """表格里只放一句话；完整依据（含链接）放下面的详情行，免得把表格撑到屏幕外。"""
        suggestion = self.suggestions.get(model)
        if suggestion is not None and suggestion.fact is not None:
            return f"{suggestion.fact.source}：{suggestion.fact.detail[:34]}…"
        if suggestion is not None and suggestion.hint is not None:
            return f"同后缀 {suggestion.hint.suffix} 你标过 {suggestion.hint.support} 台都是「{suggestion.hint.value}」"
        if model in self.manual:
            return f"你按后缀分组指定为「{self.manual[model]}」"
        return "没有依据，可从同后缀或同厂商的车型推断"

    # -- 交互 ------------------------------------------------------------- #
    def _show_detail(self, _event: tk.Event | None = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        model = selection[0]
        source, value, evidence = self._effective(model)
        target = f"→ 「{value}」" if value else "→（还没有值）"
        self.detail_var.set(f"{model} {target}　依据：{source}｜{evidence}")

    def _toggle_from_click(self, event: tk.Event) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":  # 只认「选」这一列
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        if row_id in self.checked:
            self.checked.discard(row_id)
        else:
            self.checked.add(row_id)
        self.refresh()

    def check_all_suggested(self) -> None:
        self.checked = {model for model, (source, value, _e) in ((m, self._effective(m)) for m in self.rows) if value}
        self.refresh()

    def check_none(self) -> None:
        self.checked.clear()
        self.refresh()

    def apply_group(self) -> None:
        value = self.value_var.get().strip()
        if not value:
            messagebox.showinfo("燃料待确认", "请先在右边的输入框里填一个燃料值（也可以从下拉里选）。", parent=self)
            return
        label = self.group_var.get().split("（", 1)[0]
        targets = next((models for suffix, models in self.groups if suffix == label), [])
        if not targets:
            return
        for model in targets:
            self.manual[model] = value
            self.checked.add(model)
        self.refresh()
        self.app.status_var.set(f"已把后缀 {label} 的 {len(targets)} 个车型标成「{value}」，确认无误后点「应用勾选项」写入")

    def apply_checked(self) -> None:
        targets = {model: self._effective(model)[1] for model in self.rows if model in self.checked}
        targets = {model: value for model, value in targets.items() if value}
        if not targets:
            messagebox.showinfo("燃料待确认", "还没有勾选任何可写入的车型。", parent=self)
            return
        if not messagebox.askyesno(
            "燃料待确认",
            f"将给 {len(targets)} 个车型写入燃料标签：\n" + "、".join(f"{m}→{v}" for m, v in list(targets.items())[:6])
            + ("…" if len(targets) > 6 else "")
            + "\n\n写入前会自动备份 车型标签.json。确定继续吗？",
            parent=self,
        ):
            return
        for model, value in targets.items():
            apply_model_tag_values(self.overrides, model, {self.field: value})
        try:
            save_tag_overrides(application_dir(), self.overrides)
        except ProfileError as error:
            messagebox.showerror("写入失败", str(error), parent=self)
            return
        self.app.reload_data()
        self.app.notify_data_changed({"tags"})
        self.app.status_var.set(f"已写入 {len(targets)} 个燃料标签（可撤销）")
        self.destroy()

    def export_rows(self) -> None:
        if not self.table.rows():
            messagebox.showinfo("导出", "没有可导出的内容。", parent=self)
            return
        path = self.table.export(self, "燃料待确认.csv")
        if path:
            self.app.status_var.set(f"已导出待确认表到 {path}")


class TagRulesWindow(WindowMixin, tk.Toplevel):
    """标签规则：按顺序匹配型号名，命中就套用对应标签（不用一个个改）。"""

    COLUMNS = (
        ("pattern", "匹配（子串或正则，不区分大小写）", 240, "w"),
        ("manufacturer", "厂商", 110, "w"),
        ("fuel", "燃料", 90, "center"),
        ("length", "车长", 90, "center"),
        ("hits", "命中", 60, "center"),
    )

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)
        self.theme = self.app.theme
        self.setup_dialog(parent, "标签规则", 800, 560)
        pad = self.theme.px
        self.rules: list[dict] = [dict(rule) for rule in self.app.tag_overrides.get("rules", []) if isinstance(rule, dict)]
        self.count_var = tk.StringVar()

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text="标签规则", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(frame, textvariable=self.count_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")
        ttk.Label(
            frame,
            text="从上往下匹配，某一条以最先定义的为准；规则优先于前缀映射，单车覆盖仍然最优先。改完立即生效。",
            style="Muted.TLabel",
            wraplength=pad(740),
        ).grid(row=1, column=0, sticky="w", pady=(pad(4), pad(10)))

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor, numeric=identifier == "hits")
                for identifier, title, width, anchor in self.COLUMNS
            ),
            empty_text="还没有标签规则\n点「新增规则」可以给一批车型统一设标签",
            on_delete=self.delete_rule,
        )
        self.table.card.grid(row=2, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.tree.bind("<Return>", lambda _event: self.edit_rule())

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(actions, text="新增规则", style="Primary.TButton", command=self.add_rule).pack(side="left")
        ttk.Button(actions, text="修改选中", style="Secondary.TButton", command=self.edit_rule).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="删除选中", style="Danger.TButton", command=self.delete_rule).pack(side="left")
        ttk.Button(actions, text="上移", style="Secondary.TButton", command=lambda: self.move(-1)).pack(side="left", padx=(pad(8), 0))
        ttk.Button(actions, text="下移", style="Secondary.TButton", command=lambda: self.move(1)).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")

        self.refresh()

    # -- state ----------------------------------------------------------- #
    def hit_count(self, rule: dict) -> int:
        return len(rule_hits(str(rule.get("pattern", "")), model_names(self.app.ranges)))

    def refresh(self, select: int | None = None) -> None:
        rows = [
            (
                rule.get("pattern", ""),
                rule.get("manufacturer", ""),
                rule.get("fuel", ""),
                rule.get("length", ""),
                self.hit_count(rule),
            )
            for rule in self.rules
        ]
        # iid 是 self.rules 里的下标：上移/下移/修改/删除都靠它，排序后也必须指对
        self.table.set_rows(rows, keep_selection=False, iid=lambda position, row: str(position))
        matched = sum(1 for rule in self.rules if self.hit_count(rule))
        self.count_var.set(f"共 {len(self.rules)} 条规则 · {matched} 条有命中")
        if select is not None and 0 <= select < len(self.rules):
            self.tree.selection_set(str(select))
            self.tree.focus(str(select))
        elif self.rules:
            self.tree.selection_set("0")

    def persist(self) -> bool:
        overrides = clean_tag_overrides(self.app.tag_overrides if isinstance(self.app.tag_overrides, dict) else {})
        if self.rules:
            overrides["rules"] = [dict(rule) for rule in self.rules]
        else:
            overrides.pop("rules", None)
        cleaned = clean_tag_overrides(overrides)
        try:
            save_tag_overrides(application_dir(), cleaned)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return False
        self.app.set_tag_overrides(cleaned)
        self.rules = [dict(rule) for rule in cleaned.get("rules", [])]
        self.refresh()
        return True

    def selected_index(self) -> int | None:
        selection = self.tree.selection()
        return int(selection[0]) if selection else None

    # -- actions --------------------------------------------------------- #
    def add_rule(self) -> None:
        TagRuleEditorWindow(self, None)

    def edit_rule(self) -> None:
        index = self.selected_index()
        if index is not None:
            TagRuleEditorWindow(self, index)

    def delete_rule(self) -> None:
        index = self.selected_index()
        if index is None:
            messagebox.showinfo("删除规则", "请先选中要删除的规则。", parent=self)
            return
        rule = self.rules[index]
        if not messagebox.askyesno("删除规则", f"确定删除规则「{rule.get('pattern', '')}」吗？", parent=self):
            return
        self.rules.pop(index)
        self.persist()

    def move(self, offset: int) -> None:
        index = self.selected_index()
        if index is None:
            return
        target = index + offset
        if not 0 <= target < len(self.rules):
            return
        self.rules[index], self.rules[target] = self.rules[target], self.rules[index]
        if self.persist():
            self.refresh(select=target)


class TagRuleEditorWindow(WindowMixin, tk.Toplevel):
    """新增 / 修改一条标签规则，输入时实时预览会命中哪些车型。"""

    FIELDS = (("manufacturer", "厂商"), ("fuel", "燃料"), ("length", "车长"), ("body", "车体形式"))

    def __init__(self, manager: TagRulesWindow, index: int | None) -> None:
        super().__init__(manager)
        self.manager = manager
        self.app = manager.app
        self.theme = self.app.theme
        self.index = index
        self.models = model_names(self.app.ranges)
        rule = manager.rules[index] if index is not None else {}
        self.setup_dialog(manager, "修改规则" if index is not None else "新增规则", 560, 440, resizable=False)
        pad = self.theme.px

        suggestions = tag_suggestions(self.app.ranges, self.app.tag_overrides)
        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(14), pady=pad(14))
        body = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        ttk.Label(body, text="匹配条件", style="CardMuted.TLabel").grid(row=0, column=0, sticky="w", pady=pad(5))
        self.pattern_var = tk.StringVar(value=str(rule.get("pattern", "")))
        pattern_entry = ttk.Entry(body, textvariable=self.pattern_var)
        pattern_entry.grid(row=0, column=1, sticky="ew", pady=pad(5), padx=(pad(10), 0))
        pattern_entry.bind("<Return>", lambda _event: self.save())
        self.pattern_var.trace_add("write", self.update_preview)

        self.vars: dict[str, tk.StringVar] = {}
        for row, (field, label) in enumerate(self.FIELDS, start=1):
            ttk.Label(body, text=label, style="CardMuted.TLabel").grid(row=row, column=0, sticky="w", pady=pad(5))
            variable = tk.StringVar(value=str(rule.get(field, "")))
            self.vars[field] = variable
            combo = ttk.Combobox(body, textvariable=variable, values=suggestions[field], state="normal")
            combo.grid(row=row, column=1, sticky="ew", pady=pad(5), padx=(pad(10), 0))
            combo.bind("<Return>", lambda _event: self.save())

        self.preview_var = tk.StringVar()
        next_row = 1 + len(self.FIELDS)
        ttk.Label(body, textvariable=self.preview_var, style="CardMuted.TLabel", wraplength=pad(460), justify="left").grid(
            row=next_row, column=0, columnspan=2, sticky="w", pady=(pad(10), 0)
        )
        ttk.Label(
            body,
            text="只填要改的项；例：匹配 C7BTD、燃料=天然气。",
            style="CardMuted.TLabel",
        ).grid(row=next_row + 1, column=0, columnspan=2, sticky="w", pady=(pad(8), 0))

        actions = ttk.Frame(body, style="Card.TFrame")
        actions.grid(row=next_row + 2, column=0, columnspan=2, sticky="ew", pady=(pad(14), 0))
        ttk.Button(actions, text="保存", style="Primary.TButton", command=self.save).pack(side="right")
        ttk.Button(actions, text="取消", style="Secondary.TButton", command=self.close_dialog).pack(side="right", padx=pad(8))

        self.update_preview()
        pattern_entry.focus_set()

    def update_preview(self, *_args: object) -> None:
        pattern = self.pattern_var.get().strip()
        if not pattern:
            self.preview_var.set("请输入匹配条件：可以是型号里的一段（如 C7BTD），也可以是正则。")
            return
        if compiled_pattern(pattern) is None:
            self.preview_var.set("⚠ 正则写法有误，这条规则会被忽略。")
            return
        hits = rule_hits(pattern, self.models)
        if not hits:
            self.preview_var.set("没有车型命中这条规则。")
            return
        sample = "、".join(hits[:5]) + ("…" if len(hits) > 5 else "")
        self.preview_var.set(f"命中 {len(hits)} 个车型：{sample}")

    def save(self) -> None:
        pattern = self.pattern_var.get().strip()
        values = {field: self.vars[field].get().strip() for field, _label in self.FIELDS}
        if not pattern:
            messagebox.showerror("无法保存", "请填写匹配条件。", parent=self)
            return
        if compiled_pattern(pattern) is None:
            messagebox.showerror("无法保存", "正则写法有误，请检查匹配条件。", parent=self)
            return
        if not any(values.values()):
            messagebox.showerror("无法保存", "至少要填一个要套用的标签。", parent=self)
            return
        rule = {"pattern": pattern}
        rule.update({field: value for field, value in values.items() if value})
        if self.index is None:
            self.manager.rules.append(rule)
            position = len(self.manager.rules) - 1
        else:
            self.manager.rules[self.index] = rule
            position = self.index
        if not self.manager.persist():
            return
        self.manager.refresh(select=position)
        self.app.status_var.set(f"已保存标签规则：{pattern}")
        self.destroy()


# --------------------------------------------------------------------------- #
# Metro batch lookup
# --------------------------------------------------------------------------- #


class MetroBatchWindow(WindowMixin, tk.Toplevel):
    """批量查询地铁列车编号。"""

    def __init__(self, parent: "MetroWindow") -> None:
        super().__init__(parent)
        self.app = parent.app
        self.theme = self.app.theme
        self.setup_dialog(parent, "地铁批量查询", 660, 560)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(12)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(4, weight=1)

        ttk.Label(frame, text="每行一个列车编号，也可以用空格、逗号分隔；重复的会自动去重。", style="Muted.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, pad(6))
        )
        self.number_text = self.theme.text_box(frame, 6)
        self.number_text.grid(row=1, column=0, sticky="ew")
        self.number_text.bind("<Control-Return>", lambda _event: (self.run_batch(), "break")[1])

        controls = ttk.Frame(frame)
        controls.grid(row=2, column=0, sticky="ew", pady=(pad(8), pad(10)))
        ttk.Button(controls, text="查询", style="Primary.TButton", command=self.run_batch).pack(side="left")
        ttk.Button(controls, text="清空", style="Secondary.TButton", command=self.clear).pack(side="left", padx=pad(8))
        ttk.Button(controls, text="⇩ 导出 CSV", style="Secondary.TButton", command=self.export_rows).pack(side="right")
        ttk.Button(controls, text="复制选中行", style="Secondary.TButton", command=self.copy_rows).pack(side="right", padx=pad(8))

        self.table = DataTable(
            frame,
            self.theme,
            (
                TableColumn("number", "列车编号", 120, "center"),
                TableColumn("line", "线路", 220, "w"),
                TableColumn("model", "车型", 200, "w"),
            ),
            empty_text="等待查询",
        )
        self.table.card.grid(row=4, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.tree.tag_configure("missing", foreground=self.theme.accent)

        self.summary_var = tk.StringVar(value="等待查询")
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel").grid(row=5, column=0, sticky="w", pady=(pad(6), 0))
        self.number_text.focus_set()

    def clear(self) -> None:
        self.number_text.delete("1.0", "end")
        self.table.set_rows([], keep_selection=False)
        self.summary_var.set("等待查询")
        self.number_text.focus_set()

    def run_batch(self) -> None:
        numbers, ignored = parse_metro_number_list(self.number_text.get("1.0", "end-1c"))
        rows: list[tuple[str, ...]] = []
        if not numbers and not ignored:
            self.table.set_rows([], keep_selection=False)
            self.summary_var.set("请先粘贴或输入列车编号。")
            return
        hits = 0
        for number in numbers:
            matches = find_metro_models(number, self.app.metro_ranges)
            if not matches:
                rows.append((number, "未收录", "—"))
                continue
            hits += 1
            for offset, item in enumerate(matches):
                rows.append((number if offset == 0 else "", item.line, item.model))
        self.table.set_rows(rows, keep_selection=False)
        # 「未收录」整行标红（斑马纹之外再加一个前景色 tag）
        for item_id, row in zip(self.tree.get_children(), rows):
            if row[1] == "未收录":
                self.tree.item(item_id, tags=("missing",))
        parts = [f"共 {len(numbers)} 条编号", f"命中 {hits}", f"未收录 {len(numbers) - hits}"]
        if ignored:
            parts.append(f"忽略 {len(ignored)} 条无效内容（{'、'.join(ignored[:5])}{'…' if len(ignored) > 5 else ''}）")
        self.summary_var.set(" · ".join(parts))

    def copy_rows(self) -> None:
        count = copy_tree_rows(self, self.tree)
        self.summary_var.set(f"已复制 {count} 行到剪贴板。" if count else "没有可复制的结果，请先查询。")

    def export_rows(self) -> None:
        if not tree_rows(self.tree):
            self.summary_var.set("没有可导出的结果，请先查询。")
            return
        path = export_tree_rows(self, self.tree, "地铁批量查询结果.csv")
        if path:
            self.summary_var.set(f"已导出 {len(tree_rows(self.tree))} 行到 {path}")


# --------------------------------------------------------------------------- #
# 增量导入 Excel
# --------------------------------------------------------------------------- #


class MergeWindow(WindowMixin, tk.Toplevel):
    """把外部号段表与当前数据合并：先看差异，再决定策略。"""

    COLUMNS = (
        ("kind", "类型", 80, "center"),
        ("start", "起始", 100, "center"),
        ("end", "结束", 100, "center"),
        ("model", "车型", 230, "w"),
        ("before", "原来是", 230, "w"),
    )

    def __init__(self, parent: tk.Misc, incoming: list[BusRange], source: str = "") -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)  # type: ignore[attr-defined]
        self.theme = self.app.theme
        self.setup_dialog(parent, "增量导入号段", 860, 620)
        pad = self.theme.px
        self.incoming = incoming
        self.plan = merge_ranges(self.app.ranges, incoming)
        self.strategy_var = tk.StringVar(value="incoming")

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        ttk.Label(frame, text="增量导入号段", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.summary_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")
        ttk.Label(
            frame,
            text=f"来源：{source or '所选文件'}；当前 {len(self.app.ranges)} 个号段，导入表 {len(incoming)} 个。",
            style="Muted.TLabel",
            wraplength=pad(780),
        ).grid(row=1, column=0, sticky="w", pady=(pad(4), pad(8)))

        strategy_row = ttk.Frame(frame)
        strategy_row.grid(row=2, column=0, sticky="ew", pady=(0, pad(8)))
        ttk.Label(strategy_row, text="合并方式").pack(side="left", padx=(0, pad(8)))
        ttk.Radiobutton(strategy_row, text="以导入表为准（含删除）", value="incoming", variable=self.strategy_var, command=self.refresh).pack(
            side="left"
        )
        ttk.Radiobutton(strategy_row, text="只增不改（保留现有数据）", value="additive", variable=self.strategy_var, command=self.refresh).pack(
            side="left", padx=pad(10)
        )

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor, numeric=identifier in {"start", "end"})
                for identifier, title, width, anchor in self.COLUMNS
            ),
            empty_text="没有差异",
        )
        self.table.card.grid(row=3, column=0, sticky="nsew")
        self.tree = self.table.tree
        # 差异类型用前景色区分：新增 / 修改 / 冲突 / 删除
        for tag, colour in (
            ("added", self.theme.success),
            ("updated", self.theme.primary),
            ("removed", self.theme.accent),
            ("conflict", self.theme.danger),
        ):
            self.tree.tag_configure(tag, foreground=colour)

        self.result_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.result_var, style="Muted.TLabel", wraplength=pad(780)).grid(row=4, column=0, sticky="w", pady=(pad(8), 0))
        actions = ttk.Frame(frame)
        actions.grid(row=5, column=0, sticky="ew", pady=(pad(8), 0))
        ttk.Button(actions, text="应用并保存", style="Primary.TButton", command=self.apply).pack(side="left")
        ttk.Button(actions, text="取消", style="Secondary.TButton", command=self.close_dialog).pack(side="right")

        self.refresh()

    def refresh(self, *_args: object) -> None:
        plan = self.plan
        rows: list[tuple[str, str, int, int, str, str]] = []
        rows += [("added", "新增", item.start, item.end, item.model, "") for item in plan.added]
        rows += [("updated", "修改", item.start, item.end, item.model, old) for item, old in plan.updated]
        rows += [("conflict", "冲突", item.start, item.end, item.model, f"与 {old} 重叠") for item, old in plan.conflict]
        rows += [("removed", "删除", item.start, item.end, item.model, "导入表里没有") for item in plan.removed]
        self.table.set_rows([row[1:] for row in rows], keep_selection=False)
        # 前景色标签要按行补上（DataTable 只负责斑马纹）
        for item_id, row in zip(self.tree.get_children(), rows):
            self.tree.item(item_id, tags=(row[0],))
        self.summary_var.set(plan.summary)
        kept = apply_merge(self.app.ranges, plan, self.strategy_var.get())
        self.result_var.set(
            f"按当前方式保存后：{len(kept)} 个号段"
            + ("（只增不改不会删除或覆盖任何现有号段）" if self.strategy_var.get() == "additive" else "（导入表里没有的现有号段会被删除）")
        )

    def apply(self) -> None:
        strategy = self.strategy_var.get()
        merged = apply_merge(self.app.ranges, self.plan, strategy)
        if not merged:
            messagebox.showinfo("增量导入", "合并结果是空表，已取消保存。", parent=self)
            return
        if not messagebox.askyesno(
            "增量导入",
            f"将把号段表写成 {len(merged)} 行（当前 {len(self.app.ranges)} 行）。\n保存前会自动备份原文件。\n\n确定继续吗？",
            parent=self,
        ):
            return
        try:
            save_ranges(application_dir() / DATA_FILENAME, merged)
        except RangeError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.update_ranges(merged)
        self.app.status_var.set(f"增量导入完成：{self.plan.summary} → 共 {len(merged)} 个号段")
        self.destroy()


# --------------------------------------------------------------------------- #
# 半自动建档
# --------------------------------------------------------------------------- #


class DraftProfilesWindow(WindowMixin, tk.Toplevel):
    """给没有资料的车型按模板起草简介与参数。"""

    COLUMNS = (
        ("model", "车型", 190, "w"),
        ("manufacturer", "厂商", 110, "w"),
        ("fuel", "燃料", 90, "center"),
        ("length", "车长", 80, "center"),
        ("preview", "生成的简介", 380, "w"),
    )

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)  # type: ignore[attr-defined]
        self.theme = self.app.theme
        self.setup_dialog(parent, "半自动建档", 980, 640)
        pad = self.theme.px
        images = existing_image_names(application_dir())
        self.targets = empty_profile_models(self.app.profiles, self.app.ranges, images)
        self.drafts: dict[str, dict[str, str]] = {}

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        ttk.Label(frame, text="半自动建档", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.summary_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")

        template_row = ttk.Frame(frame)
        template_row.grid(row=1, column=0, sticky="ew", pady=(pad(6), pad(4)))
        template_row.columnconfigure(1, weight=1)
        ttk.Label(template_row, text="简介模板").grid(row=0, column=0, padx=(0, pad(8)))
        self.template_var = tk.StringVar(value=DRAFT_TEMPLATE)
        ttk.Entry(template_row, textvariable=self.template_var, style="Search.TEntry").grid(row=0, column=1, sticky="ew")
        ttk.Button(template_row, text="重新生成", style="Secondary.TButton", command=self.regenerate).grid(row=0, column=2, padx=(pad(8), 0))
        ttk.Label(
            frame,
            text="可用占位符：{厂商} {燃料} {车长} {车型} {号段} {编号数} {号段数}；只会填空资料的车型，已有内容不会被覆盖。",
            style="Muted.TLabel",
            wraplength=pad(880),
        ).grid(row=2, column=0, sticky="w", pady=(0, pad(8)))

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor)
                for identifier, title, width, anchor in self.COLUMNS
            ),
            selectmode="extended",
            empty_text="没有需要建档的车型",
        )
        self.table.card.grid(row=3, column=0, sticky="nsew")
        self.tree = self.table.tree

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(actions, text="全选", style="Secondary.TButton", command=self.select_all).pack(side="left")
        ttk.Button(actions, text="生成选中的资料", style="Primary.TButton", command=self.apply_selected).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")
        ttk.Label(actions, text="Ctrl / Shift 可多选", style="Muted.TLabel").pack(side="left", padx=pad(10))

        self.regenerate()

    def regenerate(self) -> None:
        template = self.template_var.get().strip() or DRAFT_TEMPLATE
        self.drafts = {}
        for model in self.targets:
            tags = model_tags(model, self.app.tag_overrides)
            self.drafts[model] = draft_profile(model, tags, self.app.ranges, template)
        self.refresh()

    def refresh(self) -> None:
        rows = []
        for model in self.targets:
            tags = model_tags(model, self.app.tag_overrides)
            draft = self.drafts.get(model, {})
            rows.append((model, tags.manufacturer, tags.fuel, tags.length, draft.get("introduction", "")[:60]))
        # iid 用车型代号：生成资料时按它定位
        self.table.set_rows(rows, keep_selection=False, iid=lambda position, row: row[0])
        self.summary_var.set(f"共 {len(self.targets)} 个车型还没有资料")

    def select_all(self) -> None:
        self.tree.selection_set(self.tree.get_children())

    def apply_selected(self) -> None:
        models = [str(item) for item in self.tree.selection()]
        if not models:
            messagebox.showinfo("半自动建档", "请先选中要生成资料的车型。", parent=self)
            return
        created = 0
        for model in models:
            profile = dict(self.app.profiles.get(model, {}))
            draft = self.drafts.get(model, {})
            if not str(profile.get("introduction", "")).strip() and draft:
                profile["introduction"] = draft["introduction"]
            if not str(profile.get("specifications", "")).strip() and draft:
                profile["specifications"] = draft["specifications"]
            images = profile_images(profile)
            set_profile_images(profile, images)
            self.app.profiles[model] = profile
            created += 1
        try:
            save_profiles(application_dir(), self.app.profiles)
        except ProfileError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.notify_data_changed({"profiles"})
        self.app.status_var.set(f"已为 {created} 个车型起草资料（可再逐条修改）")
        self.targets = [model for model in self.targets if model not in set(models)]
        self.regenerate()


# --------------------------------------------------------------------------- #
# 统计面板
# --------------------------------------------------------------------------- #


class StatsWindow(WindowMixin, tk.Toplevel):
    """车型 / 号段 / 标签的分布统计。"""

    RANK_COLUMNS = (
        ("model", "车型", 200, "w"),
        ("manufacturer", "厂商", 120, "w"),
        ("ranges", "号段数", 80, "e"),
        ("numbers", "编号数", 90, "e"),
    )

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)  # type: ignore[attr-defined]
        self.theme = self.app.theme
        self.setup_dialog(parent, "数据统计", 1030, 660)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text="数据统计", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.summary_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")

        charts = ttk.Frame(frame)
        charts.grid(row=1, column=0, sticky="ew", pady=(pad(8), pad(10)))
        self.charts: list[tk.Canvas] = []
        for index, title in enumerate(("厂商 Top 8", "燃料", "车长", "车体形式")):
            charts.columnconfigure(index, weight=1, uniform="chart")  # 四张图等宽，不会被挤出窗口
            holder = self.theme.card(charts)
            holder.grid(row=0, column=index, sticky="nsew", padx=(0, pad(8)) if index < 3 else 0)
            ttk.Label(holder, text=title, style="Section.TLabel").pack(anchor="w", padx=pad(10), pady=(pad(8), 0))
            canvas = tk.Canvas(holder, height=pad(150), background=self.theme.surface, highlightthickness=0, bd=0)
            canvas.pack(fill="both", expand=True, padx=1, pady=(0, 1))
            self.charts.append(canvas)

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor, numeric=identifier in {"ranges", "numbers"})
                for identifier, title, width, anchor in self.RANK_COLUMNS
            ),
            sort_key=self._rank_sort_key,
            empty_text="还没有可以统计的数据",
        )
        self.table.card.grid(row=2, column=0, sticky="nsew")
        self.tree = self.table.tree

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(actions, text="⟳ 刷新", style="Secondary.TButton", command=self.refresh).pack(side="left")
        ttk.Button(actions, text="⇩ 导出 CSV", style="Secondary.TButton", command=self.export_rows).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")
        ttk.Label(actions, text="按覆盖的自编号数量排序，只列出前 40 个车型", style="Muted.TLabel").pack(side="left", padx=pad(10))

        self.refresh()

    def refresh(self) -> None:
        theme = self.theme
        counts = Counter(item.model for item in self.app.ranges)
        covered = Counter()
        for item in self.app.ranges:
            covered[item.model] += item.end - item.start + 1
        tags = {name: model_tags(name, self.app.tag_overrides) for name in model_names(self.app.ranges)}
        total_numbers = sum(covered.values())

        ranked = sorted(counts, key=lambda name: (-covered[name], name.casefold()))[:40]
        self.table.set_rows(
            [(model, tags[model].manufacturer, counts[model], covered[model]) for model in ranked],
            keep_selection=False,
            iid=lambda position, row: row[0],
        )
        self.summary_var.set(f"{len(counts)} 种车型 · {len(self.app.ranges)} 个号段 · 覆盖 {total_numbers} 个自编号")

        for canvas, (field, limit) in zip(
            self.charts,
            (("manufacturer", 8), ("fuel", None), ("length", None), ("body", None)),
        ):
            counter = Counter(getattr(tag, field) for tag in tags.values())
            self._draw_bars(canvas, counter.most_common(limit) if limit else sorted(counter.items(), key=lambda item: -item[1]))

    def _rank_sort_key(self, column: str, row: tuple[str, ...]) -> object:
        """统计表的排序键：号段数/编号数按数值，其余按中文比较。"""
        index = self.tree["columns"].index(column)
        if column in {"ranges", "numbers"}:
            return (0, int(row[index]))
        return (0, row[index].casefold())

    def _draw_bars(self, canvas: tk.Canvas, rows: list[tuple[str, int]]) -> None:
        canvas.delete("all")
        pad = self.theme.px
        if not rows:
            canvas.create_text(pad(10), pad(10), anchor="nw", text="没有数据", fill=self.theme.muted, font=self.theme.font(9))
            return
        canvas.update_idletasks()
        width = max(canvas.winfo_width(), pad(220))
        label_width = pad(78)
        bar_width = max(width - label_width - pad(40), pad(60))
        row_height = pad(15)
        top = pad(6)
        highest = max(count for _name, count in rows) or 1
        for index, (name, count) in enumerate(rows):
            y = top + index * row_height
            canvas.create_text(pad(6), y + pad(5), anchor="w", text=name[:9], fill=self.theme.text, font=self.theme.font(9))
            length = int(bar_width * count / highest)
            canvas.create_rectangle(label_width, y, label_width + max(length, 1), y + pad(10), fill=self.theme.primary, outline="")
            canvas.create_text(label_width + max(length, 1) + pad(4), y + pad(5), anchor="w", text=str(count), fill=self.theme.muted, font=self.theme.font(8))

    def export_rows(self) -> None:
        if not tree_rows(self.tree):
            messagebox.showinfo("导出", "当前没有可导出的数据。", parent=self)
            return
        path = export_tree_rows(self, self.tree, "车型统计.csv")
        if path:
            self.app.status_var.set(f"已导出统计到 {path}")


# --------------------------------------------------------------------------- #
# Bus range manager
# --------------------------------------------------------------------------- #


class RangeManagerWindow(WindowMixin, tk.Toplevel):
    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "号段管理", 720, 560)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        ttk.Label(frame, text="号段管理", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.count_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.count_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")
        ttk.Label(frame, text=f"保存会重写 {DATA_FILENAME}（原文件自动备份为 .bak），第 4 列之后的额外列会尽量保留。新增或修改时，任何与现有号段重叠的范围都无法保存。", style="Muted.TLabel", wraplength=pad(620), justify="left").grid(
            row=1, column=0, sticky="w", pady=(pad(4), pad(8))
        )

        filter_row = ttk.Frame(frame)
        filter_row.grid(row=2, column=0, sticky="ew", pady=(0, pad(10)))
        ttk.Label(filter_row, text="搜索").pack(side="left", padx=(0, pad(6)))
        self.filter_var = tk.StringVar()
        search = ttk.Entry(filter_row, textvariable=self.filter_var, style="Search.TEntry")
        search.pack(side="left", fill="x", expand=True)
        attach_placeholder(search, "按号段或车型筛选", self.theme)
        ttk.Label(filter_row, text="车型").pack(side="left", padx=(pad(12), pad(6)))
        self.model_var = tk.StringVar(value=ANY_TAG)
        self.model_combo = ttk.Combobox(
            filter_row,
            textvariable=self.model_var,
            values=[ANY_TAG] + model_names(self.app.ranges),
            state="readonly",
            width=22,
        )
        self.model_combo.pack(side="left")
        ttk.Button(filter_row, text="⟳ 重置", style="Ghost.TButton", command=self.reset_filters).pack(side="left", padx=(pad(8), 0))
        self._debounced_refresh = debounce(self, "range-search", 150, self.refresh)
        self.filter_var.trace_add("write", self._debounced_refresh)
        self.model_var.trace_add("write", self._debounced_refresh)

        self.table = DataTable(
            frame,
            self.theme,
            (
                TableColumn("start", "号段起始", 120, "e", numeric=True),
                TableColumn("end", "号段结束", 120, "e", numeric=True),
                TableColumn("model", "车型", 380, "w"),
            ),
            empty_text="没有号段",
            on_open=self.edit_selected,
            on_delete=self.delete_selected,
            on_search=lambda: (self.filter_var.set(""), search.focus_set()),
        )
        self.table.card.grid(row=3, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.empty_state = self.table.empty

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(actions, text="新增号段", style="Primary.TButton", command=self.add_range).pack(side="left")
        ttk.Button(actions, text="修改选中", style="Secondary.TButton", command=self.edit_selected).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="删除选中", style="Danger.TButton", command=self.delete_selected).pack(side="left")
        ttk.Label(actions, text="双击一行可直接修改", style="Muted.TLabel").pack(side="right")

        self.refresh()

    def reset_filters(self) -> None:
        self.filter_var.set("")
        self.model_var.set(ANY_TAG)

    def on_data_changed(self, kinds: set[str]) -> None:
        if kinds & {"all", "ranges"}:
            self.model_combo.configure(values=[ANY_TAG] + model_names(self.app.ranges))
            if self.model_var.get() not in self.model_combo.cget("values"):
                self.model_var.set(ANY_TAG)
            self.refresh()

    def visible_indices(self) -> list[int]:
        query = self.filter_var.get().strip().casefold()
        model = self.model_var.get()
        result: list[int] = []
        for index, item in enumerate(self.app.ranges):
            if model != ANY_TAG and item.model != model:
                continue
            if query and query not in f"{item.start} {item.end} {item.model}".casefold():
                continue
            result.append(index)
        return result

    def refresh(self, *_args: object) -> None:
        indices = self.visible_indices()
        rows = [
            (self.app.ranges[index].start, self.app.ranges[index].end, self.app.ranges[index].model)
            for index in indices
        ]
        # iid 仍用「在 self.app.ranges 里的下标」：编辑/删除都靠它定位。
        # 表头排序会打乱显示顺序，所以按行内容反查下标（号段不重叠，起点唯一）。
        owner = {str(self.app.ranges[index].start): str(index) for index in indices}
        self.table.set_rows(rows, keep_selection=True, iid=lambda position, row: owner[row[0]])
        total = len(self.app.ranges)
        self.count_var.set(f"共 {total} 个号段" if len(indices) == total else f"匹配 {len(indices)} / {total} 个号段")
        self.empty_state.update(
            "没有匹配的号段\n试试换个搜索词或把车型筛选改回「全部」" if not indices else None,
            has_rows=bool(indices),
        )

    def add_range(self) -> None:
        RangeEditorWindow(self, None)

    def edit_selected(self) -> None:
        selected = self.tree.selection()
        if selected:
            RangeEditorWindow(self, int(selected[0]))

    def delete_selected(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("删除号段", "请先选中要删除的号段。", parent=self)
            return
        index = int(selected[0])
        item = self.app.ranges[index]
        if not messagebox.askyesno("删除号段", f"确定删除 {item.start}-{item.end}（{item.model}）吗？", parent=self):
            return
        ranges = self.app.ranges.copy()
        ranges.pop(index)
        try:
            save_ranges(application_dir() / DATA_FILENAME, ranges)
        except RangeError as error:
            messagebox.showerror("删除失败", str(error), parent=self)
            return
        self.app.update_ranges(ranges)
        self.refresh()


class RangeEditorWindow(WindowMixin, tk.Toplevel):
    def __init__(self, parent: tk.Misc, index: int | None, initial: tuple[int, int, str] | None = None) -> None:
        """parent 需要提供 app / theme（号段管理、数据体检、快查工具都满足）。"""
        super().__init__(parent)
        self.manager = parent
        self.app: BusLookupApp = getattr(parent, "app", parent)  # type: ignore[attr-defined]
        self.theme = self.app.theme
        self.index = index
        self.setup_dialog(parent, "新增号段" if index is None else "修改号段", 430, 320, resizable=False)
        pad = self.theme.px

        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(16), pady=pad(16))
        header = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14), pad(16), 0))
        header.pack(fill="x")
        self.dialog_header(
            header,
            "修改号段" if index is not None else "新增号段",
            "号段与车型一一对应，保存时会写回 Excel（原文件自动备份）。",
            inside_card=True,
        )
        frame = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        self.start_var = tk.StringVar()
        self.end_var = tk.StringVar()
        self.model_var = tk.StringVar()
        labels = (("号段起始", self.start_var), ("号段结束", self.end_var), ("对应车型", self.model_var))
        for row, (label, variable) in enumerate(labels):
            ttk.Label(frame, text=label, style="CardMuted.TLabel").grid(row=row, column=0, sticky="w", pady=pad(5))
            entry = ttk.Entry(frame, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", pady=pad(5), padx=(pad(10), 0))
            entry.bind("<Return>", lambda _event: self.save())
        if index is not None:
            item = self.app.ranges[index]
            self.start_var.set(str(item.start))
            self.end_var.set(str(item.end))
            self.model_var.set(item.model)
        elif initial is not None:
            self.start_var.set(str(initial[0]))
            self.end_var.set(str(initial[1]))
            self.model_var.set(initial[2] if len(initial) > 2 else "")
            ttk.Label(
                frame,
                text="这是「未收录」区间，填上车型即可补录；保存后会写入 Excel 并自动备份。",
                style="CardMuted.TLabel",
                wraplength=pad(300),
                justify="left",
            ).grid(row=len(labels), column=0, columnspan=2, sticky="w", pady=(pad(4), 0))

        actions = ttk.Frame(frame, style="Card.TFrame")
        actions.grid(row=len(labels) + 1, column=0, columnspan=2, pady=(pad(14), 0), sticky="e")
        ttk.Button(actions, text="取消", style="Secondary.TButton", command=self.destroy).pack(side="left", padx=(0, pad(8)))
        ttk.Button(actions, text="保存", style="Primary.TButton", command=self.save).pack(side="left")

    def save(self) -> None:
        try:
            item = validate_range(self.start_var.get(), self.end_var.get(), self.model_var.get(), self.app.ranges, self.index)
        except RangeError as error:
            messagebox.showerror("无法保存", str(error), parent=self)
            return
        ranges = self.app.ranges.copy()
        if self.index is None:
            ranges.append(item)
        else:
            ranges[self.index] = item
        try:
            save_ranges(application_dir() / DATA_FILENAME, ranges)
        except RangeError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.update_ranges(ranges)
        refresher = getattr(self.manager, "refresh", None)
        if callable(refresher):
            refresher()
        self.destroy()


class BulkRangeWindow(WindowMixin, tk.Toplevel):
    """把一批未收录的自编号补录成号段。"""

    COLUMNS = (
        ("start", "号段起始", 110, "center"),
        ("end", "号段结束", 110, "center"),
        ("count", "编号数", 80, "center"),
        ("model", "车型", 240, "w"),
    )

    def __init__(self, parent: tk.Misc, intervals: list[tuple[int, int]], default_model: str = "") -> None:
        super().__init__(parent)
        self.app: BusLookupApp = getattr(parent, "app", parent)
        self.theme = self.app.theme
        self.setup_dialog(parent, "批量补录号段", 720, 560)
        pad = self.theme.px
        self.rows: list[tuple[int, int, str]] = [(start, end, default_model) for start, end in intervals]

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        ttk.Label(frame, text="批量补录号段", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            frame,
            text="这些自编号当前没有收录。填好车型后一次写入 Excel（保存前自动备份）；与现有号段重叠或车型为空的行会被拦下。",
            style="Muted.TLabel",
            wraplength=pad(660),
        ).grid(row=1, column=0, sticky="w", pady=(pad(4), pad(8)))

        model_row = ttk.Frame(frame)
        model_row.grid(row=2, column=0, sticky="ew", pady=(0, pad(8)))
        ttk.Label(model_row, text="车型").pack(side="left", padx=(0, pad(6)))
        self.model_var = tk.StringVar(value=default_model)
        combo = ttk.Combobox(model_row, textvariable=self.model_var, values=model_names(self.app.ranges), width=30)
        combo.pack(side="left")
        ttk.Button(model_row, text="应用到全部", style="Secondary.TButton", command=lambda: self.apply_model(all_rows=True)).pack(
            side="left", padx=pad(8)
        )
        ttk.Button(model_row, text="应用到选中", style="Secondary.TButton", command=self.apply_model).pack(side="left")

        self.table = DataTable(
            frame,
            self.theme,
            tuple(
                TableColumn(identifier, title, width, anchor, numeric=identifier in {"count"})
                for identifier, title, width, anchor in self.COLUMNS
            ),
            selectmode="extended",
            empty_text="没有待补的号段",
        )
        self.table.card.grid(row=3, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.tree.tag_configure("empty", foreground=self.theme.accent)

        self.summary_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.summary_var, style="Muted.TLabel").grid(row=4, column=0, sticky="w", pady=(pad(8), 0))
        actions = ttk.Frame(frame)
        actions.grid(row=5, column=0, sticky="ew", pady=(pad(8), 0))
        ttk.Button(actions, text="保存全部", style="Primary.TButton", command=self.save_all).pack(side="left")
        ttk.Button(actions, text="关闭", style="Secondary.TButton", command=self.close_dialog).pack(side="right")

        self.refresh()

    def refresh(self, *_args: object) -> None:
        filled = sum(1 for _start, _end, model in self.rows if model)
        rows = [
            (start, end, end - start + 1, model or "（待填）")
            for start, end, model in self.rows
        ]
        # iid 是 self.rows 的下标：批量套用车型与保存都靠它，排序后也不能变
        self.table.set_rows(rows, keep_selection=True, iid=lambda position, row: str(position))
        self.summary_var.set(f"共 {len(self.rows)} 个待补区间 · 已填车型 {filled} · 合计 {sum(end - start + 1 for start, end, _ in self.rows)} 个自编号")

    def apply_model(self, all_rows: bool = False) -> None:
        model = self.model_var.get().strip()
        if not model:
            messagebox.showinfo("批量补录", "请先填写车型。", parent=self)
            return
        if all_rows:
            targets = range(len(self.rows))
        else:
            selection = self.tree.selection()
            if not selection:
                messagebox.showinfo("批量补录", "请先选中要套用的行（Ctrl / Shift 可多选）。", parent=self)
                return
            targets = [int(item) for item in selection]
        for index in targets:
            start, end, _old = self.rows[index]
            self.rows[index] = (start, end, model)
        self.refresh()

    def build_ranges(self) -> tuple[list[BusRange], list[str]]:
        """算出要新增的号段；返回 (号段, 错误说明)。"""
        errors: list[str] = []
        pending: list[BusRange] = []
        for start, end, model in self.rows:
            if not model:
                errors.append(f"{start}-{end}：没有填车型")
                continue
            try:
                pending.append(validate_range(str(start), str(end), model, self.app.ranges + pending))
            except RangeError as error:
                errors.append(f"{start}-{end}：{error}")
        return pending, errors

    def save_all(self) -> None:
        pending, errors = self.build_ranges()
        if errors:
            messagebox.showerror("无法保存", "请先处理这些问题：\n" + "\n".join(errors[:8]) + ("\n…" if len(errors) > 8 else ""), parent=self)
            return
        if not pending:
            messagebox.showinfo("批量补录", "没有需要补录的区间。", parent=self)
            return
        if not messagebox.askyesno("批量补录", f"确定新增 {len(pending)} 个号段吗？\n（保存前会自动备份原表）", parent=self):
            return
        ranges = self.app.ranges + pending
        try:
            save_ranges(application_dir() / DATA_FILENAME, ranges)
        except RangeError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.update_ranges(ranges)
        self.app.status_var.set(f"已补录 {len(pending)} 个号段")
        messagebox.showinfo("完成", f"已写入 {len(pending)} 个号段，共 {sum(item.end - item.start + 1 for item in pending)} 个自编号。", parent=self)
        self.destroy()


# --------------------------------------------------------------------------- #
# Metro windows
# --------------------------------------------------------------------------- #


class MetroWindow(WindowMixin, tk.Toplevel):
    def __init__(self, app: BusLookupApp) -> None:
        super().__init__(app)
        self.app = app
        self.theme = app.theme
        self.setup_dialog(app, "北京地铁编号查询", 540, 430)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text="北京地铁编号查询", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )

        query_card = self.theme.card(frame)
        query_card.grid(row=1, column=0, sticky="ew", pady=(pad(12), pad(12)))
        query_inner = ttk.Frame(query_card, style="Card.TFrame", padding=(pad(14), pad(12)))
        query_inner.pack(fill="both", expand=True)
        query_inner.columnconfigure(1, weight=1)
        ttk.Label(query_inner, text="列车编号", style="CardMuted.TLabel").grid(row=0, column=0, sticky="w", padx=(0, pad(10)))
        self.number_var = tk.StringVar()
        self.entry = ttk.Entry(query_inner, textvariable=self.number_var, style="Query.TEntry")
        self.entry.grid(row=0, column=1, sticky="ew")
        self.entry.bind("<Return>", lambda _event: self.query())
        ttk.Button(query_inner, text="查询", style="Primary.TButton", command=self.query).grid(row=0, column=2, padx=(pad(10), 0))
        self.result_var = tk.StringVar(value="请输入完整列车编号，例如 01001、T401 或 YZ001。")
        ttk.Label(query_inner, textvariable=self.result_var, style="CardValue.TLabel", foreground=self.theme.primary, justify="left", wraplength=pad(430)).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(pad(12), 0)
        )
        self.copy_button = ttk.Button(query_inner, text="复制选中行", style="Secondary.TButton", command=self.copy_result)
        self.copy_button.grid(row=2, column=0, columnspan=3, sticky="w", pady=(pad(10), 0))
        self.copy_button.state(["disabled"])

        manage_card = self.theme.card(frame)
        manage_card.grid(row=2, column=0, sticky="nsew")
        manage_inner = ttk.Frame(manage_card, style="Card.TFrame", padding=(pad(14), pad(12)))
        manage_inner.pack(fill="both", expand=True)
        ttk.Label(manage_inner, text="数据管理", style="Section.TLabel").pack(anchor="w", pady=(0, pad(8)))
        ttk.Button(manage_inner, text="地铁线路简介", style="Secondary.TButton", command=lambda: MetroLineCatalogWindow(self)).pack(
            fill="x", pady=(0, pad(8))
        )
        ttk.Button(manage_inner, text="批量查询列车编号", style="Secondary.TButton", command=lambda: MetroBatchWindow(self)).pack(
            fill="x", pady=(0, pad(8))
        )
        ttk.Button(manage_inner, text="管理地铁号段与车型", style="Secondary.TButton", command=lambda: MetroRangeManagerWindow(self)).pack(fill="x")
        ttk.Separator(manage_inner).pack(fill="x", pady=pad(10))
        ttk.Button(manage_inner, text="把当前数据写回 txt 种子", style="Secondary.TButton", command=self.export_source).pack(fill="x", pady=(0, pad(8)))
        ttk.Button(manage_inner, text="重新从 txt 导入（覆盖当前数据）", style="Danger.TButton", command=self.import_source).pack(fill="x")
        ttk.Label(manage_inner, text=f"数据保存在 {METRO_DATA_FILENAME}", style="CardMuted.TLabel").pack(anchor="w", pady=(pad(8), 0))

        self.entry.focus_set()

    def export_source(self) -> None:
        try:
            path = export_metro_source(application_dir(), self.app.metro_ranges)
        except MetroError as error:
            messagebox.showerror("写入失败", str(error), parent=self)
            return
        self.app.status_var.set(f"已把当前地铁数据写回 {path.name}")

    def import_source(self) -> None:
        if not messagebox.askyesno(
            "重新导入",
            f"将用 {METRO_SOURCE_FILENAME} 的内容覆盖当前地铁数据（原数据会先备份为 .bak），确定继续吗？",
            parent=self,
        ):
            return
        try:
            ranges = import_metro_source(application_dir())
        except MetroError as error:
            messagebox.showerror("导入失败", str(error), parent=self)
            return
        self.app.metro_ranges = ranges
        messagebox.showinfo("导入完成", f"已导入 {len(ranges)} 条地铁号段。", parent=self)

    def query(self) -> None:
        try:
            matches = find_metro_models(self.number_var.get(), self.app.metro_ranges)
        except MetroError:
            self.result_var.set("请输入由字母和数字组成的完整列车编号。")
            self.copy_button.state(["disabled"])
            return
        if not matches:
            self.result_var.set("未收录 / 暂无对应车型，请核对编号。")
            self.copy_button.state(["disabled"])
            return
        self.result_var.set("；".join(f"{item.line}：{item.model}" for item in matches))
        self.copy_button.state(["!disabled"])

    def copy_result(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.result_var.get())
        self.update()


class MetroLineCatalogWindow(WindowMixin, tk.Toplevel):
    def __init__(self, parent: MetroWindow) -> None:
        super().__init__(parent)
        self.app = parent.app
        self.theme = self.app.theme
        self.setup_dialog(parent, "地铁线路简介", 480, 540)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        ttk.Label(frame, text="地铁线路简介", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w", pady=(0, pad(10))
        )

        self.table = DataTable(
            frame,
            self.theme,
            (
                TableColumn("line", "线路", 200, "w"),
                TableColumn("ranges", "号段", 70, "e", numeric=True),
                TableColumn("models", "车型", 70, "e", numeric=True),
            ),
            empty_text="还没有地铁数据",
            on_open=self.open_selected,
        )
        self.table.card.grid(row=1, column=0, sticky="nsew")
        self.tree = self.table.tree
        self.tree.bind("<Return>", self.open_selected)

        self.table.set_rows(
            [
                (line, len(metro_ranges_for_line(line, self.app.metro_ranges)),
                 len({item.model for item in metro_ranges_for_line(line, self.app.metro_ranges)}))
                for line in metro_lines(self.app.metro_ranges)
            ],
            keep_selection=False,
            iid=lambda position, row: row[0],  # iid 就是线路名，双击靠它打开详情
        )

        footer = ttk.Frame(frame)
        footer.grid(row=2, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Label(footer, text="双击一行查看该线路的车型与号段", style="Muted.TLabel").pack(side="left")
        ttk.Button(footer, text="查看线路信息", style="Primary.TButton", command=self.open_selected).pack(side="right")

    def open_selected(self, _event: tk.Event | None = None) -> None:
        selection = self.tree.selection()
        if selection:
            MetroLineDetailWindow(self, selection[0])


class MetroLineDetailWindow(WindowMixin, tk.Toplevel):
    def __init__(self, parent: MetroLineCatalogWindow, line: str) -> None:
        super().__init__(parent)
        self.app = parent.app
        self.theme = self.app.theme
        self.setup_dialog(parent, f"线路简介 - {line}", 580, 480)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text=line, style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="车型与已知号段", style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(pad(4), pad(8)))

        self.table = DataTable(
            frame,
            self.theme,
            (
                TableColumn("start", "号段起始", 130, "e"),
                TableColumn("end", "号段结束", 130, "e"),
                TableColumn("model", "车型", 240, "w"),
            ),
            empty_text="这条线路还没有号段数据",
        )
        self.table.card.grid(row=2, column=0, sticky="nsew")
        tree = self.table.tree
        self.tree = tree
        self.table.set_rows(
            [(item.start, item.end, item.model) for item in metro_ranges_for_line(line, parent.app.metro_ranges)],
            keep_selection=False,
        )

        ttk.Button(frame, text="关闭", style="Secondary.TButton", command=self.close_dialog).grid(row=3, column=0, sticky="e", pady=(pad(10), 0))


class MetroRangeManagerWindow(WindowMixin, tk.Toplevel):
    def __init__(self, parent: MetroWindow) -> None:
        super().__init__(parent)
        self.app = parent.app
        self.theme = self.app.theme
        self.setup_dialog(parent, "管理地铁号段与车型", 800, 560)
        pad = self.theme.px

        frame = ttk.Frame(self, padding=(pad(16), pad(14), pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text="地铁号段与车型管理", style="HeaderTitle.TLabel", background=self.theme.background, foreground=self.theme.text).grid(
            row=0, column=0, sticky="w"
        )
        self.count_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.count_var, style="Muted.TLabel").grid(row=0, column=0, sticky="e")
        ttk.Label(frame, text=f"同一线路、相同字母前缀的号段不可重叠；保存会重写 {METRO_DATA_FILENAME}（原文件自动备份为 .bak）。", style="Muted.TLabel", wraplength=pad(740)).grid(
            row=1, column=0, sticky="w", pady=(pad(4), pad(10))
        )

        self.table = DataTable(
            frame,
            self.theme,
            (
                TableColumn("line", "线路", 200, "w"),
                TableColumn("start", "号段起始", 130, "e"),
                TableColumn("end", "号段结束", 130, "e"),
                TableColumn("model", "车型", 240, "w"),
            ),
            empty_text="还没有地铁号段",
            on_open=self.edit,
        )
        self.table.card.grid(row=2, column=0, sticky="nsew")
        self.tree = self.table.tree

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, sticky="ew", pady=(pad(10), 0))
        ttk.Button(actions, text="新增", style="Primary.TButton", command=lambda: MetroRangeEditorWindow(self, None)).pack(side="left")
        ttk.Button(actions, text="修改选中", style="Secondary.TButton", command=self.edit).pack(side="left", padx=pad(8))
        ttk.Button(actions, text="删除选中", style="Danger.TButton", command=self.delete).pack(side="left")
        ttk.Label(actions, text="双击一行可直接修改", style="Muted.TLabel").pack(side="right")

        self.refresh()

    def refresh(self) -> None:
        rows = [(item.line, item.start, item.end, item.model) for item in self.app.metro_ranges]
        # iid 用原始下标（编辑/删除靠它定位）；排序后按行内容反查，避免改错行
        owner = {f"{item.line}|{item.start}|{item.end}|{item.model}": str(index) for index, item in enumerate(self.app.metro_ranges)}
        self.table.set_rows(rows, keep_selection=True, iid=lambda position, row: owner["|".join(row)])
        lines = len({item.line for item in self.app.metro_ranges})
        self.count_var.set(f"共 {len(self.app.metro_ranges)} 条号段 · {lines} 条线路")

    def edit(self) -> None:
        selected = self.tree.selection()
        if selected:
            MetroRangeEditorWindow(self, int(selected[0]))

    def delete(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("删除地铁号段", "请先选中要删除的号段。", parent=self)
            return
        index = int(selected[0])
        item = self.app.metro_ranges[index]
        if not messagebox.askyesno(
            "删除地铁号段", f"确定删除 {item.line} 的 {item.start}-{item.end}（{item.model}）吗？", parent=self
        ):
            return
        ranges = self.app.metro_ranges.copy()
        ranges.pop(index)
        try:
            save_metro_ranges(application_dir(), ranges)
        except MetroError as error:
            messagebox.showerror("删除失败", str(error), parent=self)
            return
        self.app.update_metro_ranges(ranges)
        self.refresh()


class MetroRangeEditorWindow(WindowMixin, tk.Toplevel):
    def __init__(self, manager: MetroRangeManagerWindow, index: int | None) -> None:
        super().__init__(manager)
        self.manager = manager
        self.app = manager.app
        self.theme = manager.theme
        self.index = index
        self.setup_dialog(manager, "新增地铁号段" if index is None else "修改地铁号段", 460, 300, resizable=False)
        pad = self.theme.px

        card = self.theme.card(self)
        card.pack(fill="both", expand=True, padx=pad(16), pady=pad(16))
        frame = ttk.Frame(card, style="Card.TFrame", padding=(pad(16), pad(14)))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        self.variables = [tk.StringVar() for _ in range(4)]
        for row, label in enumerate(("线路", "号段起始", "号段结束", "车型")):
            ttk.Label(frame, text=label, style="CardMuted.TLabel").grid(row=row, column=0, sticky="w", pady=pad(5))
            entry = ttk.Entry(frame, textvariable=self.variables[row])
            entry.grid(row=row, column=1, sticky="ew", pady=pad(5), padx=(pad(10), 0))
            entry.bind("<Return>", lambda _event: self.save())
        if index is not None:
            item = self.app.metro_ranges[index]
            for variable, value in zip(self.variables, (item.line, item.start, item.end, item.model)):
                variable.set(value)

        actions = ttk.Frame(frame, style="Card.TFrame")
        actions.grid(row=4, column=0, columnspan=2, pady=(pad(14), 0), sticky="e")
        ttk.Button(actions, text="取消", style="Secondary.TButton", command=self.destroy).pack(side="left", padx=(0, pad(8)))
        ttk.Button(actions, text="保存", style="Primary.TButton", command=self.save).pack(side="left")

    def save(self) -> None:
        try:
            item = validate_metro_range(*(variable.get() for variable in self.variables), self.app.metro_ranges, self.index)
        except MetroError as error:
            messagebox.showerror("无法保存", str(error), parent=self)
            return
        ranges = self.app.metro_ranges.copy()
        if self.index is None:
            ranges.append(item)
        else:
            ranges[self.index] = item
        try:
            save_metro_ranges(application_dir(), ranges)
        except MetroError as error:
            messagebox.showerror("保存失败", str(error), parent=self)
            return
        self.app.update_metro_ranges(ranges)
        self.manager.refresh()
        self.destroy()


USAGE = f"""用法：
  {APP_TITLE}.exe                     打开图形界面
  {APP_TITLE}.exe 20000               直接查一个自编号（结果打印到控制台）
  {APP_TITLE}.exe --export-html 文件.html   导出单文件网页版
  {APP_TITLE}.exe --net-check         自检联网（这个程序能不能连上第三方服务）
  {APP_TITLE}.exe --help              显示这段说明

双击运行时没有控制台窗口，结果会改成弹窗显示并复制到剪贴板。
"""


def attach_console() -> bool:
    """打包成 --windowed 后没有控制台；这里附着到父进程的控制台。"""
    if sys.platform != "win32":
        return False
    try:
        return bool(ctypes.windll.kernel32.AttachConsole(-1))
    except Exception:
        return False


def console_write(text: str) -> bool:
    """把文本送到调用者的终端。

    三种情况：源码运行时 sys.stdout 直接可用；打包成 --windowed 后 sys.stdout 是空的，
    就附着到父进程的控制台并写 CONOUT$（先把输出页设成 UTF-8，写完恢复）。
    """
    if sys.stdout is not None:
        try:
            print(text)
            return True
        except (OSError, ValueError):
            pass
    if sys.platform != "win32":
        return False
    try:
        kernel32 = ctypes.windll.kernel32
        previous = kernel32.GetConsoleOutputCP()
    except Exception:
        return False
    try:
        kernel32.SetConsoleOutputCP(65001)
        with open("CONOUT$", "w", encoding="utf-8", errors="replace") as console:
            console.write(text + "\n")
            console.flush()
        return True
    except (OSError, ValueError):
        return False
    finally:
        try:
            kernel32.SetConsoleOutputCP(previous)
        except Exception:
            pass


def cli_report(title: str, lines: list[str]) -> None:
    """命令行模式的输出：优先写终端；拿不到终端就**落盘**，绝不弹窗。

    为什么不再弹窗：打包成 --windowed 后没有控制台，一旦回退到 messagebox 就会**永久阻塞**
    （实测 `exe --self-test > out.txt` 卡死 5 分钟、只能强杀）。命令行就该安静退出，
    所以改成写一个文件，路径写在 README 里。
    """
    text = "\n".join(lines)
    if console_write(text):
        return
    # 打包后 sys.stdout 是空的：先附着父进程控制台再试一次
    if attach_console() and console_write(text):
        return
    safe_title = "".join(char for char in title if char.isalnum() or char in " _-") or "output"
    target = application_dir() / f"命令行输出-{safe_title}.txt"
    try:
        target.write_text(text + "\n", encoding="utf-8")
    except OSError:
        pass


def cli_lookup(number_text: str) -> tuple[int, list[str]]:
    """查一个自编号，返回 (退出码, 输出行)。"""
    base = application_dir()
    try:
        ranges = load_ranges(base / DATA_FILENAME)
    except DataError as error:
        return 2, [f"读不到 {DATA_FILENAME}：{error}"]
    number = validate_number(number_text)
    if number is None:
        return 2, [f"「{number_text}」不是五位自编号。"]
    overrides = load_tag_overrides(base)
    model = find_model(number, ranges)
    if model is None:
        return 1, [f"{number:05d} 未收录"]
    item = find_range(number, ranges)
    tags = model_tags(model, overrides)
    lines = [f"{number:05d} → {model}", f"厂商：{tags.manufacturer}　燃料：{tags.fuel}　车长：{tags.length}"]
    if item is not None:
        lines.append(f"所属号段：{item.start}-{item.end}")
    return 0, lines


def cli_export_html(target: str) -> tuple[int, list[str]]:
    base = application_dir()
    try:
        ranges = load_ranges(base / DATA_FILENAME)
        profiles = load_profiles(base)
        metro_ranges = load_metro_ranges(base)
    except (DataError, ProfileError, MetroError) as error:
        return 2, [f"读取数据失败：{error}"]
    overrides = load_tag_overrides(base)
    try:
        written, size = export_web_page(Path(target), ranges, overrides, profiles, metro_ranges, base)
    except DataError as error:
        return 2, [str(error)]
    return 0, [f"已导出：{written}", f"大小：{size / 1024:.0f} KB", "双击即可打开，也可以发到手机上离线查看。"]


def _selfcheck_tag_item(base: Path):
    """给自检用的标签覆盖率检查（标签推导在界面层，所以放这儿注入）。"""
    try:
        ranges = load_ranges(base / DATA_FILENAME)
    except Exception as error:  # noqa: BLE001
        return CheckItem("tags", "车型标签", False, f"读不到号段表：{error}", "先修号段表")
    names = model_names(ranges)
    if not names:
        return CheckItem("tags", "车型标签", False, "没有车型可标", "先修号段表")
    overrides = load_tag_overrides(base / TAG_FILENAME)
    unknown = 0
    for name in names:
        tag = model_tags(name, overrides)
        if any(value == UNKNOWN_TAG for value in (tag.manufacturer, tag.fuel, tag.length, tag.body)):
            unknown += 1
    covered = len(names) - unknown
    ratio = covered / len(names) if names else 0
    if ratio < 0.9:
        return CheckItem(
            "tags", "车型标签", False,
            f"{len(names)} 种车型里 {unknown} 种还有未确认项（覆盖率 {ratio:.0%}）",
            "在「数据体检」里逐项确认，或用「标签规则」一次标一批",
            blocking=False,  # 只是提示：没标完不影响查车型
        )
    return CheckItem("tags", "车型标签", True, f"{covered}/{len(names)} 种齐全（{ratio:.0%}）")


def cli_self_test(as_json: bool = False, deep: bool = False) -> int:
    """命令行自检：不开窗口，直接报告数据/磁盘/版本（`--deep` 再测联网）。"""
    import bpt_selfcheck as check

    base = application_dir()
    items = check.run_self_checks(
        base, load_settings(base), deep=deep, tag_checker=_selfcheck_tag_item
    )
    if as_json:
        for line in check.as_json(items).splitlines():
            console_write(line)
    else:
        cli_report("自检", check.format_report(items))
    return 0 if not check.has_blocking_failure(items) else 1


def cli_net_check() -> int:
    """自检联网：这个 EXE 自己能不能连上第三方服务（用来区分"服务端问题"和"本机拦了"）。

    为什么要单独做这个：联网失败时界面只说"连不上"，用户没法判断是服务端挂了、
    还是**这个程序被防火墙/安全软件拦了**（python 探针通、EXE 被拒就是后者的典型症状）。
    """
    import socket
    import time
    import urllib.request

    lines: list[str] = []
    settings = online_settings()
    lines.append(f"服务地址：{settings.base_url}")
    lines.append(f"联网开关：{'开' if settings.enabled else '关'}"
                 f"　超时 {settings.timeout_sec}s　连接超时 {settings.connect_timeout_sec}s")
    if offline_forced():
        lines.append("程序被设为离线（联网功能整体关闭）")
        cli_report("联网自检", lines)
        return 1
    if not settings.usable:
        lines.append("设置里把联网关了（联网功能整体关闭）")
        cli_report("联网自检", lines)
        return 1

    host = settings.base_url.split("//", 1)[-1].split("/", 1)[0]
    lines.append(f"\n① 域名解析 {host}")
    try:
        address = socket.gethostbyname(host)
        lines.append(f"   解析到 {address}")
    except OSError as error:
        lines.append(f"   失败：{type(error).__name__}: {error}（DNS 的问题）")
        cli_report("联网自检", lines)
        return 1

    lines.append("\n② 裸 TCP 连接 443（不走代理，和程序里的探通一样）")
    try:
        mark = time.perf_counter()
        with socket.create_connection((host, 443), timeout=settings.connect_timeout_sec):
            lines.append(f"   通了 {time.perf_counter() - mark:.2f}s")
    except OSError as error:
        lines.append(f"   失败：{type(error).__name__}: {error}")
        lines.append("   → 连 TCP 都不通：本机网络 / 防火墙 / VPN 拦了这个程序")
        cli_report("联网自检", lines)
        return 1

    lines.append("\n③ HTTPS 请求（走 urllib，会读系统代理设置）")
    proxies = urllib.request.getproxies()
    lines.append(f"   系统代理：{proxies or '（没有）'}")
    try:
        request = urllib.request.Request(
            f"{settings.base_url}/sse",
            headers={"Accept": "text/event-stream", "User-Agent": USER_AGENT},
        )
        mark = time.perf_counter()
        with urllib.request.urlopen(request, timeout=settings.connect_timeout_sec) as response:
            lines.append(f"   成功 HTTP {response.status}（{time.perf_counter() - mark:.2f}s）")
    except Exception as error:  # noqa: BLE001 - 自检要把原因原样告诉用户
        lines.append(f"   失败：{type(error).__name__}: {error}")
        lines.append(f"   → {error_text(error)}")
        cli_report("联网自检", lines)
        return 1

    lines.append("\n④ MCP 握手 + 查一次线路号（真用一遍程序里的客户端）")
    client = MCPClient(settings=settings)
    try:
        mark = time.perf_counter()
        client.connect()
        lines.append(f"   握手成功 {time.perf_counter() - mark:.1f}s")
        mark = time.perf_counter()
        routes = client.query_route_ids("886")
        lines.append(f"   查 886：{time.perf_counter() - mark:.1f}s → "
                     + "；".join(f"{first} → {last}" for _rid, first, last in routes))
        lines.append("\n结论：联网功能正常。")
        code = 0
    except OnlineError as error:
        lines.append(f"   失败：{type(error).__name__}: {error}")
        lines.append(f"   → {error_text(error)}")
        code = 1
    finally:
        client.close()
    cli_report("联网自检", lines)
    return code


def main(argv: list[str] | None = None) -> int:
    """图形界面 + 命令行两种入口。"""
    args = list(sys.argv[1:] if argv is None else argv)
    if args:
        if args[0] in ("-h", "--help", "/?"):
            cli_report("用法", USAGE.strip().splitlines())
            return 0
        if args[0] == "--export-html":
            target = args[1] if len(args) > 1 else f"{APP_TITLE}.html"
            code, lines = cli_export_html(target)
            cli_report("导出网页版", lines)
            return code
        if args[0] == "--net-check":
            return cli_net_check()
        if args[0] in ("--self-test", "--selftest"):
            # 默认不联网（快、不会卡在慢网络上）；要测联网加 --deep，或用 --net-check
            return cli_self_test(as_json="--json" in args, deep="--deep" in args)
        code, lines = cli_lookup(args[0])
        cli_report(f"{APP_TITLE} - {args[0]}", lines)
        return code
    BusLookupApp().mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
