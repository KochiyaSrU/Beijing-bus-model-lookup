"""数据与文件层：常量、异常、备份、日志、xlsx 读写、号段与编号的解析和查找。

这里不依赖 tkinter，也不认识界面；主文件 BPT_BusDigitalNumber.py 会把本模块的名字
全部再导出一次，所以老的 `from BPT_BusDigitalNumber import ...` 仍然可用。
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import traceback
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

DATA_FILENAME = "BPT_BusDigitalNumber.xlsx"


LINES_FILENAME = "线路.json"
ONLINE_CACHE_FILENAME = "联网缓存.json"
LINE_INFO_FILENAME = "线路资料.json"
ROUTE_LIST_FILENAME = "线路目录.json"
METRO_SOURCE_FILENAME = "北京地铁车型.txt"


METRO_DATA_FILENAME = "北京地铁数据.json"


PROFILE_FILENAME = "车型资料.json"


TAG_FILENAME = "车型标签.json"


ICON_FILENAME = "北京公交车型查询器.ico"


IMAGE_DIRECTORY_NAME = "车型图片"


BACKUP_DIRECTORY_NAME = "备份"


BACKUP_KEEP = 10


ERROR_LOG_FILENAME = "错误日志.txt"


MAX_ERROR_LOG_BYTES = 256 * 1024


SETTINGS_FILENAME = "界面设置.json"


HISTORY_FILENAME = "历史记录.json"


SPOT_LOG_FILENAME = "拍车日志.json"

PHOTO_DIRECTORY_NAME = "拍车照片"


HISTORY_LIMIT = 20


CLIPBOARD_INTERVAL_MS = 400


CREATE_NO_WINDOW = 0x08000000


IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"}


MAX_IMAGE_PIXELS = (1600, 1200)


JPEG_QUALITY = 82


EXPECTED_HEADERS = ("号段起始", "号段结束", "车型")


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


OFFICE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


APP_TITLE = "北京公交车型查询器"


class DataError(Exception):
    """Raised when the external workbook cannot provide a safe mapping."""


class ProfileError(Exception):
    """Raised when local model profiles or their images cannot be managed safely."""


class RangeError(Exception):
    """Raised when a bus-number range cannot be safely changed."""


class MetroError(Exception):
    """Raised when metro mapping data is invalid or cannot be saved."""


@dataclass(frozen=True)
class BusRange:
    start: int
    end: int
    model: str


@dataclass(frozen=True)
class MetroRange:
    line: str
    start: str
    end: str
    model: str


def application_dir() -> Path:
    """Return the folder containing the executable, or this script in development."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(name: str) -> Path:
    """Locate a bundled resource, whether running from source or a frozen EXE."""
    bundled = getattr(sys, "_MEIPASS", None)
    if bundled:
        candidate = Path(bundled) / name
        if candidate.is_file():
            return candidate
    return application_dir() / name


def backup_directory(base_directory: Path) -> Path:
    """所有自动备份都放在这个子目录里（不存在时按需创建）。"""
    return base_directory / BACKUP_DIRECTORY_NAME


def backup_file(path: Path, keep: int = BACKUP_KEEP) -> Path | None:
    """把现有文件复制成 备份/ 下带时间戳的副本，只保留最近 keep 份。

    备份失败不阻塞保存本身，返回备份路径（没备份成返回 None）。
    """
    if not path.is_file():
        return None
    directory = backup_directory(path.parent)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    target = directory / f"{path.stem}.{stamp}{path.suffix}"
    try:
        directory.mkdir(exist_ok=True)
        shutil.copy2(path, target)
    except OSError:
        return None
    prune_backups(directory, path.stem, keep)
    return target


def prune_backups(directory: Path, stem: str, keep: int = BACKUP_KEEP) -> None:
    """同一份数据的备份只保留最近 keep 份。"""
    if keep <= 0:
        return
    try:
        candidates = sorted(
            (item for item in directory.iterdir() if item.is_file() and item.name.split(".")[0] == stem),
            key=lambda item: item.name,
            reverse=True,
        )
    except OSError:
        return
    for stale in candidates[keep:]:
        try:
            stale.unlink()
        except OSError:
            pass


BACKUP_SOURCES: tuple[tuple[str, str], ...] = (
    (DATA_FILENAME, "公交号段表"),
    (PROFILE_FILENAME, "车型资料"),
    (TAG_FILENAME, "车型标签"),
    (METRO_DATA_FILENAME, "地铁数据"),
)


def list_backups(base_directory: Path) -> list[tuple[str, Path, str, float, int]]:
    """列出可恢复的备份：(数据名, 备份文件, 原文件名, 修改时间, 大小)，新的在前。

    既包括 备份/ 目录里的，也包括早期版本留在根目录的 *.bak。
    """
    entries: list[tuple[str, Path, str, float, int]] = []
    by_stem = {Path(name).stem: (label, name) for name, label in BACKUP_SOURCES}
    directory = backup_directory(base_directory)
    if directory.is_dir():
        for path in directory.iterdir():
            if not path.is_file():
                continue
            found = by_stem.get(path.name.split(".")[0])
            if not found:
                continue
            label, original = found
            stat = path.stat()
            entries.append((label, path, original, stat.st_mtime, stat.st_size))
    for name, label in BACKUP_SOURCES:
        legacy = base_directory / (name + ".bak")
        if legacy.is_file():
            stat = legacy.stat()
            entries.append((f"{label}（旧版 .bak）", legacy, name, stat.st_mtime, stat.st_size))
    entries.sort(key=lambda item: item[3], reverse=True)
    return entries


def restore_backup(base_directory: Path, backup: Path, original_name: str) -> Path:
    """用备份覆盖当前数据；覆盖前会先把当前状态也备份一份。"""
    target = base_directory / original_name
    backup_file(target)
    try:
        shutil.copy2(backup, target)
    except OSError as error:
        raise DataError(f"无法恢复：{error}") from error
    return target


def log_exception(exc_type: type, exc_value: BaseException, exc_tb: object, context: str = "") -> Path | None:
    """把异常堆栈追加写入 错误日志.txt（过大时轮转一次）。"""
    path = application_dir() / ERROR_LOG_FILENAME
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))  # type: ignore[arg-type]
    entry = f"[{stamp}] {context or '未捕获异常'}\n{text}{'-' * 60}\n"
    try:
        if path.is_file() and path.stat().st_size > MAX_ERROR_LOG_BYTES:
            path.replace(path.with_name(path.stem + ".old" + path.suffix))
        with path.open("a", encoding="utf-8") as handle:
            handle.write(entry)
    except OSError:
        return None
    return path


def history_path(base_directory: Path) -> Path:
    return base_directory / HISTORY_FILENAME


def load_history(base_directory: Path) -> list[str]:
    """最近查询过的自编号（五位数字，新的在前）。"""
    path = history_path(base_directory)
    if not path.is_file():
        return []
    try:
        contents = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    raw = contents.get("numbers") if isinstance(contents, dict) else contents
    if not isinstance(raw, list):
        return []
    numbers: list[str] = []
    for item in raw:
        number = validate_number(str(item))
        if number is not None:
            text = f"{number:05d}"
            if text not in numbers:
                numbers.append(text)
    return numbers[:HISTORY_LIMIT]


def remember_number(base_directory: Path, number: str, limit: int = HISTORY_LIMIT) -> list[str]:
    """把一次查询记进历史（去重、最新在前、超长截断），返回新的历史列表。"""
    value = validate_number(number)
    if value is None:
        return load_history(base_directory)
    text = f"{value:05d}"
    numbers = [item for item in load_history(base_directory) if item != text]
    numbers.insert(0, text)
    numbers = numbers[: max(1, limit)]
    _write_json(history_path(base_directory), {"numbers": numbers})
    return numbers


@dataclass(frozen=True)
class SpotEntry:
    """一条拍车记录。

    image 是「车型主图」的文件名（只作参考），photo 是用户自己拍的照片文件名。
    """

    time: str
    number: str
    model: str
    note: str = ""
    image: str = ""
    photo: str = ""


def spot_log_path(base_directory: Path) -> Path:
    return base_directory / SPOT_LOG_FILENAME


def load_spots(base_directory: Path) -> list[SpotEntry]:
    path = spot_log_path(base_directory)
    if not path.is_file():
        return []
    try:
        contents = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    raw = contents.get("entries") if isinstance(contents, dict) else contents
    if not isinstance(raw, list):
        return []
    entries: list[SpotEntry] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        number = str(item.get("number", "")).strip()
        model = str(item.get("model", "")).strip()
        if not number and not model:
            continue
        entries.append(
            SpotEntry(
                time=str(item.get("time", "")).strip(),
                number=number,
                model=model,
                note=str(item.get("note", "")).strip(),
                image=str(item.get("image", "")).strip(),
                photo=str(item.get("photo", "")).strip(),
            )
        )
    return entries


def save_spots(base_directory: Path, entries: list[SpotEntry]) -> None:
    _write_json(spot_log_path(base_directory), {"entries": [entry.__dict__ for entry in entries]}, backup=True)


def append_spot(base_directory: Path, entry: SpotEntry) -> list[SpotEntry]:
    """追加一条拍车记录并落盘（保存前自动备份）。"""
    entries = load_spots(base_directory)
    entries.insert(0, entry)
    save_spots(base_directory, entries)
    return entries


def make_spot(number: str, model: str, note: str = "", image: str = "", photo: str = "") -> SpotEntry:
    value = validate_number(number)
    return SpotEntry(
        time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        number=f"{value:05d}" if value is not None else str(number).strip(),
        model=model.strip(),
        note=note.strip(),
        image=image.strip(),
        photo=photo.strip(),
    )


def replace_spot(base_directory: Path, index: int, entry: SpotEntry) -> list[SpotEntry]:
    """替换第 index 条记录（编辑已有拍车记录用），返回新的列表。"""
    entries = load_spots(base_directory)
    if not 0 <= index < len(entries):
        raise IndexError(index)
    entries[index] = entry
    save_spots(base_directory, entries)
    return entries


def spot_summary(entries: list[SpotEntry]) -> dict[str, object]:
    """拍车日志的简单统计：条数、车型种类、Top 车型、日期范围、有照片的条数。"""
    models = Counter(entry.model for entry in entries if entry.model)
    days = sorted({entry.time[:10] for entry in entries if entry.time[:10]})
    return {
        "total": len(entries),
        "models": len(models),
        "numbers": len({entry.number for entry in entries if entry.number}),
        "with_photo": sum(1 for entry in entries if entry.photo),
        "with_note": sum(1 for entry in entries if entry.note),
        "top": models.most_common(5),
        "first_day": days[0] if days else "",
        "last_day": days[-1] if days else "",
        "days": len(days),
    }


def spot_summary_text(summary: dict[str, object]) -> str:
    """把统计结果写成一行中文。"""
    total = int(summary.get("total", 0))
    if not total:
        return "还没有拍车记录"
    parts = [f"共 {total} 条", f"{summary.get('models', 0)} 种车型", f"{summary.get('numbers', 0)} 个自编号"]
    if summary.get("days"):
        parts.append(f"{summary.get('days')} 天（{summary.get('first_day')} ~ {summary.get('last_day')}）")
    if summary.get("with_photo"):
        parts.append(f"{summary['with_photo']} 条带照片")
    return " · ".join(str(part) for part in parts)


def filter_spots(entries: list[SpotEntry], query: str = "", model: str | None = None, day: str = "") -> list[SpotEntry]:
    """按关键字、车型、日期前缀筛拍车记录（model 为空表示不限车型）。"""
    text = query.strip().casefold()
    result: list[SpotEntry] = []
    for entry in entries:
        if model and entry.model != model:
            continue
        if day and not entry.time.startswith(day):
            continue
        if text and text not in f"{entry.number} {entry.model} {entry.note} {entry.time}".casefold():
            continue
        result.append(entry)
    return result


def spot_models(entries: list[SpotEntry]) -> list[str]:
    return sorted({entry.model for entry in entries if entry.model}, key=str.casefold)


@dataclass(frozen=True)
class RideRecord:
    """一条乘车运转记录（来自用户自己的记录表）。"""

    year: str
    date: str
    line: str
    number: str
    model: str
    board: str = ""
    alight: str = ""
    note: str = ""


RIDE_HEADERS = ("线路", "车型")

# 地铁 / 铁路车型的写法，用来把「乘车记录」里的非公交部分分出来
METRO_MODEL_PATTERN = re.compile(r"^(DKZ|SFM|ZBM|BDK|CCD|QKZ|XJZ|CRH|CR400|25[A-Z]|\d{2}A\d{2})", re.IGNORECASE)
OUT_OF_TOWN_KEYWORDS = (
    "洛阳", "沈阳", "上海", "太原", "锦州", "承德", "天津", "保定", "涿州", "廊坊", "唐山",
    "石家庄", "张家口", "邯郸", "秦皇岛", "济南", "青岛", "郑州", "南京", "杭州", "武汉",
    "西安", "成都", "重庆", "广州", "深圳", "长沙", "合肥", "福州", "厦门", "昆明", "贵阳",
    "哈尔滨", "长春", "大连", "呼和浩特", "包头", "银川", "兰州", "西宁", "乌鲁木齐", "拉萨",
    "浦东", "浦西", "武清", "三河", "燕郊", "香河", "大厂", "怀来", "兴隆", "滦平", "丰宁",
    "芦台", "汉沽", "迁安", "遵化", "玉田", "蓟州", "宝坻", "静海", "沧州", "衡水", "邢台",
)


def looks_like_metro_model(model: str) -> bool:
    """DKZ / SFM / 02A05 / 25G 这类是地铁或铁路车型，不属于公交号段表。"""
    return bool(METRO_MODEL_PATTERN.match(model.strip()))


def looks_like_out_of_town(line: str) -> bool:
    """线路名里出现外地城市/地区名，基本可以判定不是北京公交。

    铁路车次（G/D/K/T/Z + 数字）也算在外，它们不属于公交线网。
    """
    text = line.strip()
    if not text:
        return False
    if re.fullmatch(r"[GDKTZ]\d{1,4}", text.upper()):
        return True
    if re.fullmatch(r"(T|S|K|Z|G|D)\s?\d+", text.upper()):
        return True
    return any(keyword in text for keyword in OUT_OF_TOWN_KEYWORDS)


def parse_ride_sheet(rows: list[list[str]]) -> list[RideRecord]:
    """把一张运转记录工作表解析成记录列表（自动找表头行）。"""
    if not rows:
        return []
    header_index = None
    for index, row in enumerate(rows[:20]):
        if all(any(cell.strip() == name for cell in row) for name in RIDE_HEADERS):
            header_index = index
            break
    if header_index is None:
        return []
    head = rows[header_index]
    columns = {name: head.index(name) for name in head if name}
    year = ""
    for row in rows[:header_index]:
        digits = re.sub(r"\D", "", row[0]) if row else ""
        if len(digits) >= 4:
            year = digits[:4]
            break
    records: list[RideRecord] = []
    for row in rows[header_index + 1:]:
        def cell(name: str) -> str:
            index = columns.get(name)
            return row[index].strip() if index is not None and len(row) > index else ""

        line, model = cell("线路"), cell("车型")
        if not line or not model:
            continue
        records.append(
            RideRecord(
                year=year,
                date=cell("日期"),
                line=line,
                number=cell("车辆编号"),
                model=model,
                board=cell("上车站"),
                alight=cell("下车站"),
                note=cell("备注"),
            )
        )
    return records


def load_ride_records(path: Path) -> list[RideRecord]:
    """读一个运转记录工作簿里的所有年度工作表。"""
    records: list[RideRecord] = []
    for _name, rows in read_sheets(path):
        records.extend(parse_ride_sheet(rows))
    return records


# 型号名"看起来像"不等于笔误：同一款车换批次会有 -1 / -5，换能源会变成 BTD / BTB，
# 有些还带厂内代号（(HC9)）。所以这里**只描述两个写法的差异性质**，判据始终是
# "同一个自编号、两个来源说法不同"这个硬事实，不做相似度猜笔误。
MISMATCH_ORDER = (
    "完全不搭（可能记录写错/号段表是别的车）",
    "批次或能源后缀不同",
    "少/多一个字符",
    "写法差异（多了厂内代号/括号）",
)


def _model_skeleton(text: str) -> str:
    """去掉括号内容、空格、连字符、下划线再转大写（判断是不是"同一个型号的两种写法"）。"""
    value = re.sub(r"[（(][^)）]*[)）]", "", str(text).upper())
    return re.sub(r"[\s\-_·/]", "", value)


def _edit_distance(left: str, right: str) -> int:
    if left == right:
        return 0
    if not left or not right:
        return len(left) or len(right)
    previous = list(range(len(right) + 1))
    for i, lchar in enumerate(left, start=1):
        current = [i]
        for j, rchar in enumerate(right, start=1):
            current.append(min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (lchar != rchar)))
        previous = current
    return previous[-1]


def classify_model_mismatch(ours: str, theirs: str) -> str:
    """同一个自编号在号段表和记录里写法不同时，差异属于哪一类（只描述差异，不判对错）。"""
    ours, theirs = str(ours).strip(), str(theirs).strip()
    flat_ours, flat_theirs = ours.replace(" ", "").upper(), theirs.replace(" ", "").upper()
    if flat_ours == flat_theirs:
        return MISMATCH_ORDER[3]
    # 一个把另一个整个包住、并且多出来的部分**带分隔符或括号**（例如 "(HC9)"）→ 多半是厂内代号/说明。
    # 注意：只是尾部多了个字母（C7BHD vs C7BH）不算，那可能是别的批次或真的少写了一位。
    for longer, shorter in ((ours, theirs), (theirs, ours)):
        if longer.upper().startswith(shorter.upper()) and len(longer) > len(shorter):
            rest = longer[len(shorter):]
            if rest[:1] in (" ", "(", "（", "-", "_", "·", "/") or rest[:1].isdigit():
                return MISMATCH_ORDER[3]
    # 去掉括号与连字符后完全相同 → 纯写法差异
    if _model_skeleton(ours) == _model_skeleton(theirs):
        return MISMATCH_ORDER[3]
    # 前半段完全一样、只有尾部（批次 / 能源那一位）不同
    if "-" in ours and "-" in theirs:
        head_ours, _sep, tail_ours = ours.rpartition("-")
        head_theirs, _sep2, tail_theirs = theirs.rpartition("-")
        if head_ours.upper() == head_theirs.upper() and tail_ours.upper() != tail_theirs.upper():
            return MISMATCH_ORDER[1]
    # 长度接近、只差一两个字符（BHD/BH、BTB/BTD 这种）——值得瞄一眼，但未必是错
    skeleton_ours, skeleton_theirs = _model_skeleton(ours), _model_skeleton(theirs)
    distance = _edit_distance(skeleton_ours, skeleton_theirs)
    if distance <= 2 and abs(len(skeleton_ours) - len(skeleton_theirs)) <= 2:
        return MISMATCH_ORDER[2]
    return MISMATCH_ORDER[0]


def reconcile_rides(
    records: list[RideRecord],
    our_models: set[str],
    metro_models: set[str] | None = None,
    ranges: list[BusRange] | None = None,
) -> dict[str, object]:
    """把运转记录和我们的号段表对账。

    关键是**用记录里的自编号去号段表里查**（自编号才是硬证据），然后分四种情况：
      matched      —— 编号查得到，且型号与记录一致
      mismatch     —— 编号查得到，但型号与记录不同（要么号段表旧了，要么记录里写错了）
      number_gap   —— 编号在号段表里查不到（我们缺这段号段；记录里的型号正好可以用来补录）
      model_gap    —— 编号查不到，且型号也不在我们的车型表里
    地铁 / 外地记录单独分出来。
    """
    metro_known = set(metro_models or set())
    totals = {"all": len(records), "metro": 0, "out_of_town": 0, "no_number": 0}

    matched: Counter = Counter()
    mismatch: Counter = Counter()  # (我们的型号, 记录里的型号)
    mismatch_examples: dict[tuple[str, str], set[str]] = {}
    number_gap: Counter = Counter()  # (编号, 记录里的型号) —— 我们缺的号段
    model_gap: Counter = Counter()
    metro_missing: Counter = Counter()
    out_of_town: Counter = Counter()
    out_of_town_lines: dict[str, set[str]] = {}
    line_models: dict[str, Counter] = {}
    seen_models: set[str] = set()
    in_table_models: set[str] = set()

    for record in records:
        if looks_like_metro_model(record.model):
            totals["metro"] += 1
            if record.model not in metro_known:
                metro_missing[record.model] += 1
            continue
        if looks_like_out_of_town(record.line):
            totals["out_of_town"] += 1
            out_of_town[record.model] += 1
            out_of_town_lines.setdefault(record.model, set()).add(record.line)
            continue
        seen_models.add(record.model)
        number = validate_number(record.number)
        if number is None:
            totals["no_number"] = int(totals["no_number"]) + 1
            recent = find_range(int(record.number), ranges) if (ranges and record.number.isdigit()) else None
            model = recent.model if recent else ""
            if model:
                in_table_models.add(model)
                pair = (model, record.model)
                mismatch[pair] += 1
                mismatch_examples.setdefault(pair, set()).add(f"{record.line} {record.number}")
            else:
                model_gap[record.model] += 1
            continue
        item = find_range(number, ranges) if ranges else None
        if item is None:
            number_gap[(f"{number:05d}", record.model)] += 1
            continue
        in_table_models.add(item.model)
        line_models.setdefault(record.line, Counter())[item.model] += 1
        if item.model == record.model:
            matched[item.model] += 1
        else:
            pair = (item.model, record.model)
            mismatch[pair] += 1
            mismatch_examples.setdefault(pair, set()).add(f"{record.line} {record.number}")

    mismatch_rows = [
        (
            classify_model_mismatch(ours_model, their_model),
            f"{ours_model} → {their_model}",
            str(count),
            "、".join(sorted(mismatch_examples.get((ours_model, their_model), set()))[:5]),
        )
        for (ours_model, their_model), count in mismatch.most_common()
    ]
    # 按"值得查"的程度排：完全不搭的最前，纯写法的最后（不猜相似度，证据是同一个自编号）
    mismatch_rows.sort(key=lambda row: (MISMATCH_ORDER.index(row[0]) if row[0] in MISMATCH_ORDER else 9, -int(row[2])))
    number_rows = [(number, model, str(count)) for (number, model), count in number_gap.most_common()]
    return {
        "totals": totals,
        "matched": [(model, str(count)) for model, count in matched.most_common()],
        "mismatch": mismatch_rows,
        "mismatch_pairs": len(mismatch),
        "number_gap": number_rows,
        "number_gap_count": len(number_gap),
        "model_gap": [(model, str(count)) for model, count in model_gap.most_common()],
        "metro_missing": [(model, str(count)) for model, count in metro_missing.most_common()],
        "out_of_town": [
            (model, str(count), "、".join(sorted(out_of_town_lines.get(model, set()))[:4]))
            for model, count in out_of_town.most_common()
        ],
        "unridden": sorted(our_models - in_table_models, key=str.casefold),
        "line_models": {line: dict(counter) for line, counter in line_models.items()},
        "line_count": len(line_models),
        "pair_count": sum(len(counter) for counter in line_models.values()),
        "model_count": len(seen_models),
    }


@dataclass(frozen=True)
class TagFact:
    """查证过的标签事实：型号 → 值，附原始描述与来源链接。

    value 为空串表示"查到了但没法直接对上本程序的标签值"（例如只查到"混合动力"，
    而本程序的燃料取值里没有这一项），这时交给用户自己选，不硬塞。
    """

    model: str
    field: str
    value: str
    detail: str
    source: str
    url: str = ""


# 车迷/公告类公开来源里能直接读到的型号事实（只放**查证过**的，一条一来源）。
# 目前只用于燃料：这些型号原来的写法看不出来，靠体型/后缀猜会猜错（GS 事件）。
TAG_FACTS: tuple[TagFact, ...] = (
    TagFact("BJ6127C8BTB", "fuel", "天然气", "福田欧辉 BJ6127C8BTB 客车（天然气国五 24-63 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/10651.html"),
    TagFact("BJ6127C8MTB", "fuel", "天然气", "福田欧辉 BJ6127C8MTB 客车（天然气国五 24-63 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/10159.html"),
    TagFact("BJ6127C8BHB", "fuel", "柴油", "福田欧辉 BJ6127C8BHB 公交车（柴油国五 24-55 座）",
            "chinabuses", "https://m.chinabuses.com/product/buses/10685.html"),
    TagFact("XMQ6120C2", "fuel", "柴油", "金龙 XMQ6120C2 客车（柴油国四 24-67 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/10075.html"),
    TagFact("KLQ6129TAE41", "fuel", "柴油", "金龙 KLQ6129TAE41 客车（柴油国四 24-65 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/11634.html"),
    TagFact("ZK6125HGQAA", "fuel", "柴油", "宇通 ZK6125HGQAA 公交车（柴油国四 10-49 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/11229.html"),
    TagFact("ZK6127HE", "fuel", "柴油", "宇通 ZK6127HE 客车公告：燃料种类 柴油（国四，轴数 2）",
            "17vin 公告", "https://www.17vin.com/gonggao/399271/ZK6127HE.html"),
    TagFact("HFF6121G15C", "fuel", "天然气", "安凯 HFF6121G15C 公交车（天然气国五 24-46 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/15314.html"),
    TagFact("HFF6105G39C", "fuel", "天然气", "安凯 HFF6105G39C 公交车（天然气国五 25-40 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/12939.html"),
    TagFact("JNP6130GSC", "fuel", "天然气", "青年 JNP6130GSC 双层公交车（液化天然气 LNG 国五 30-82 座）",
            "chinabuses", "https://www.chinabuses.com/product/buses/10655.html"),
    TagFact("JNP6140GVC", "fuel", "天然气", "青年 JNP6140GVC 公交车（天然气国五 30-49 座；30-49 座说明是单层，"
            "不是双层——本表按 14 米级判三轴与此一致）", "chinabuses",
            "https://www.chinabuses.com/product/buses/11901.html"),
    TagFact("ZK6902HGA", "fuel", "柴油", "宇通 ZK6902HGA 城市客车公告：燃料种类 柴油（国Ⅳ，轴数 2）",
            "17vin 公告", "https://www.17vin.com/gonggao/400626/ZK6902HGA.html"),
    TagFact("KLQ6128GE4", "fuel", "柴油", "金龙 KLQ6128GE4 城市客车公告：燃料种类 柴油（国Ⅳ，轴数 2）",
            "17vin 公告", "https://www.17vin.com/gonggao/216246/KLQ6128GE4.html"),
    TagFact("HFF6123KZ-2", "fuel", "柴油", "安凯 HFF6123KZ-2 客车公告：燃料种类 柴油（国Ⅳ，轴数 2）",
            "auto-che 公告", "https://cn.auto-che.com/v/hff/hff6123kz-2-264-ankai.html"),
    TagFact("KLQ6129GE5", "fuel", "柴油", "海格 KLQ6129GE5 城市客车公告：燃料种类 柴油（轴数 2）",
            "17vin 公告", "https://www.17vin.com/gonggao/216292/KLQ6129GE5.html"),
    TagFact("ZK6129HGA9", "fuel", "柴油", "宇通 ZK6129HGA9 城市客车公告：燃料种类 柴油（国Ⅳ，轴数 2）",
            "17vin 公告", "https://www.17vin.com/gonggao/399368/ZK6129HGA9.html"),
    TagFact("BJ6123C7C4D", "fuel", "", "福田牌 BJ6123C7C4D 混合动力城市客车（公告只写「混合动力」，"
            "本程序燃料取值里没有这一项，请你选一个）", "17vin 公告",
            "https://www.17vin.com/gonggao/30560/BJ6123C7C4D.html"),
)
_TAG_FACT_INDEX = {(fact.model, fact.field): fact for fact in TAG_FACTS}


def fact_for(model: str, field: str) -> TagFact | None:
    return _TAG_FACT_INDEX.get((model.strip(), field))


@dataclass(frozen=True)
class TagRuleHint:
    """从用户已有标注里学到的「型号后缀 → 标签值」规律。"""

    field: str
    suffix: str
    value: str
    support: int
    examples: tuple[str, ...] = ()


@dataclass(frozen=True)
class TagSuggestion:
    """给某个车型的建议标签值，附上它是怎么推出来的。"""

    model: str
    field: str
    value: str
    hint: TagRuleHint | None = None
    fact: TagFact | None = None

    @property
    def source(self) -> str:
        if self.fact is not None:
            return "已核实" if self.fact.value else "查到待你选"
        return "从你的标注学到"

    def evidence(self) -> str:
        if self.fact is not None:
            text = f"{self.fact.source}：{self.fact.detail}"
            return f"{text}（{self.fact.url}）" if self.fact.url else text
        if self.hint is None:
            return ""
        names = "、".join(self.hint.examples[:3])
        more = f" 等 {len(self.hint.examples)} 台" if len(self.hint.examples) > 3 else ""
        return f"你标过的 {names}{more} 都是「{self.hint.value}」（后缀 {self.hint.suffix}）"


TAG_SUFFIX_LENGTHS = (3, 2)
MIN_TAG_RULE_SUPPORT = 2
BRAND_PREFIX_PATTERN = re.compile(r"^[A-Za-z]+")  # 厂商代号前缀，如 BJ / HFF / CA


def tag_suffix_candidates(model: str) -> list[str]:
    """车型名的后缀候选（长的优先）。

    先去掉 `-1` 这类变体尾巴，再取末尾 3 个 / 2 个字符；要求以字母结尾且至少
    含两个字母，免得把 `HFF6105G39C` 这种别的厂商命名方式也算进来。
    """
    base = re.sub(r"-\d+$", "", model.strip().upper())
    candidates: list[str] = []
    for length in TAG_SUFFIX_LENGTHS:
        if len(base) < length:
            continue
        tail = base[-length:]
        if tail[-1].isalpha() and sum(character.isalpha() for character in tail) >= 2:
            candidates.append(tail)
    return candidates


def model_tag_override(overrides: dict | None, model: str, field: str) -> str:
    """单车型覆盖里这个字段的值（没有就返回空串）。"""
    values = ((overrides or {}).get("models") or {}).get(model) or {}
    return str(values.get(field, "")).strip()


def learn_tag_rules(overrides: dict | None, field: str = "fuel") -> list[TagRuleHint]:
    """从用户已经手工标注过的车型里，归纳「后缀 → 值」的规律。

    只保留**支持度 ≥ 2 且没有冲突**的后缀（同一个后缀出现过两个不同值就不猜）。
    """
    votes: dict[str, Counter] = {}
    examples: dict[str, dict[str, list[str]]] = {}
    for model, values in (((overrides or {}).get("models")) or {}).items():
        value = str((values or {}).get(field, "")).strip()
        if not value:
            continue
        for suffix in tag_suffix_candidates(model):
            votes.setdefault(suffix, Counter())[value] += 1
            examples.setdefault(suffix, {}).setdefault(value, []).append(model)

    hints: list[TagRuleHint] = []
    for suffix, counter in votes.items():
        if len(counter) != 1:  # 同一个后缀对应过多个值 → 不猜
            continue
        value, support = counter.most_common(1)[0]
        if support < MIN_TAG_RULE_SUPPORT:
            continue
        hints.append(TagRuleHint(field, suffix, value, support, tuple(sorted(examples[suffix][value]))))
    hints.sort(key=lambda item: (-len(item.suffix), -item.support, item.suffix))
    return hints


def suggest_tag(overrides: dict | None, field: str, model: str, hints: list[TagRuleHint]) -> TagSuggestion | None:
    """按查证过的事实、再按学到的规律给一个车型出建议（已手工标过的不再建议）。"""
    if model_tag_override(overrides, model, field):
        return None
    fact = fact_for(model, field)
    if fact is not None:
        return TagSuggestion(model, field, fact.value, None, fact)
    for suffix in tag_suffix_candidates(model):
        hint = next((item for item in hints if item.suffix == suffix), None)
        if hint is not None:
            return TagSuggestion(model, field, hint.value, hint)
    return None


def suggest_tags_for_models(
    overrides: dict | None,
    field: str,
    models: list[str],
) -> tuple[list[TagSuggestion], list[str]]:
    """给一批车型出建议。返回 (有建议的列表, 仍需人工判断的车型列表)。

    有建议的里可能包含 `value == ""` 的：那是"查到了但没法直接对上标签值"，
    界面上要显示证据让人自己选，不能当默认值写进去。
    """
    hints = learn_tag_rules(overrides, field)
    suggested: list[TagSuggestion] = []
    unsure: list[str] = []
    for model in models:
        suggestion = suggest_tag(overrides, field, model, hints)
        if suggestion is None:
            unsure.append(model)
        else:
            suggested.append(suggestion)
    return suggested, unsure


def group_by_suffix(models: list[str], overrides: dict | None = None, field: str = "fuel") -> list[tuple[str, list[str]]]:
    """把车型按「后缀」分组，方便一次决定一组的标签值。

    只统计没有手工标注过的车型（已经定过的不需要再决定）。分组用最长可用后缀，
    这样 `BJ6127C8BTB` 与 `BJ6127C8BTB-1` 会落到同一组；**提不出后缀的按厂商兜底**
    （很多型号结尾是数字，比如 HFF6110GS101D、CA6120URD22，硬切后缀没意义，
    但同厂商同年代的车通常可以一起定值）。数量多的排前面。
    """
    groups: dict[str, list[str]] = {}
    for model in models:
        if model_tag_override(overrides, model, field):
            continue
        candidates = tag_suffix_candidates(model)
        if candidates:
            groups.setdefault(candidates[0], []).append(model)
            continue
        prefix = BRAND_PREFIX_PATTERN.match(model.strip().upper())
        label = f"{prefix.group()} 系（无后缀）" if prefix else "（无法归组）"
        groups.setdefault(label, []).append(model)
    return sorted(groups.items(), key=lambda item: (-len(item[1]), item[0]))


@dataclass(frozen=True)
class LineEvidence:
    """一条"凭什么这么说"的证据：这条线路记录是哪来的、什么时候查的。"""

    source: str = "manual"   # manual / import / online
    bus: str = ""            # 自编号（网上查的、表格导入的会填）
    start: str = ""          # YYYYMMDD
    end: str = ""            # YYYYMMDD
    checked: str = ""        # 查询/导入日期 YYYYMMDD

    @property
    def key(self) -> tuple[str, str, str, str]:
        return (self.source, self.bus, self.start, self.end)

    @property
    def label(self) -> str:
        span = ""
        if self.start and self.end:
            span = self.start if self.start == self.end else f"{self.start}–{self.end}"
        parts = [SOURCE_LABELS.get(self.source, self.source)]
        if self.bus:
            parts.append(self.bus)
        if span:
            parts.append(span)
        if self.checked:
            parts.append(f"{self.checked[:4]}-{self.checked[4:6]}-{self.checked[6:]} 查")
        return "　".join(parts)


@dataclass(frozen=True)
class LineLink:
    """「我记过：某车型跑过某条线路」。

    刻意不用「所属」：记录只能说明**记录过**，所以带时间范围与状态，
    由用户自己确认哪条是"现在还在跑"。
    """

    line: str
    model: str = ""
    numbers: tuple[str, ...] = ()
    first: str = ""  # 首次记录的 年.月，如 2024.9
    last: str = ""  # 末次记录的 年.月
    status: str = "recorded"  # recorded=我记过 / current=确认现在还在跑 / retired=已退出
    source: str = "manual"  # manual=我记过 / import=表格导入 / online=网上查的
    note: str = ""
    evidence: tuple[LineEvidence, ...] = ()


LINE_STATUSES = ("current", "recorded", "retired")
LINE_STATUS_LABELS = {"current": "现在还在跑", "recorded": "我记过", "retired": "已退出"}
LINE_STATUS_ORDER = {"current": 0, "recorded": 1, "retired": 2}

#: 数据来源，以及"谁说话更算数"（数字越小越算数：手工 > 导入 > 网上查的）
SOURCE_LABELS = {"manual": "我记过", "import": "表格导入", "online": "网上查的"}
SOURCE_ORDER = {"manual": 0, "import": 1, "online": 2}


def source_rank(source: str) -> int:
    return SOURCE_ORDER.get(str(source).strip(), 9)


def source_label(source: str) -> str:
    return SOURCE_LABELS.get(str(source).strip(), str(source))

# 同号家族：886路快车 / 886路区间 → 886（只用来提示"要不要合并"，绝不自动合并）
LINE_FAMILY_SUFFIXES = ("快车", "慢车", "区间", "内环", "外环", "支线", "专线", "快", "支")


def lines_path(base_directory: Path) -> Path:
    return base_directory / LINES_FILENAME


def line_key(name: str) -> str:
    """比较用的规范化线路名（不改写存储的原文）。"""
    text = name.strip().replace(" ", "").replace("　", "")
    for source, target in (("（", "("), ("）", ")"), ("-", "-")):
        text = text.replace(source, target)
    return text.upper()


def line_family(name: str) -> str:
    """同号家族的键：`886路区间` → `886路` → `886`。"""
    text = name.strip().replace(" ", "")
    for suffix in LINE_FAMILY_SUFFIXES:
        if text.endswith(suffix) and len(text) > len(suffix):
            text = text[: -len(suffix)]
    return text.rstrip("路") if text.endswith("路") else text


def line_display(link: LineLink, *, with_status: bool = False, with_source: bool = False) -> str:
    """界面上怎么念这条记录，例如「886 路（2024.9–2026.8）」。

    `with_source=True` 时末尾带上来源（「网上查的」这类），**绝不让人误以为
    网上查来的就是自己记的**。
    """
    span = ""
    if link.first and link.last:
        span = link.first if link.first == link.last else f"{link.first}–{link.last}"
    elif link.first or link.last:
        span = link.first or link.last
    text = f"{link.line}（{span}）" if span else link.line
    if with_status and link.status != "recorded":
        text += f" · {LINE_STATUS_LABELS.get(link.status, link.status)}"
    if with_source and source_rank(link.source) > SOURCE_ORDER["manual"]:
        text += f" · {source_label(link.source)}"
    return text


def month_key(text: str) -> tuple[int, int]:
    """把「2025.10」这类年月解析成可比较的 (年, 月)。

    为什么要专门做这个：字符串比较会把 `2025.10` 排在 `2025.8` **前面**（"1" < "8"），
    于是 10/11/12 月的记录就会被当成更早的数据——合并时取 max、排序时都会错。
    """
    parts = [int(piece) for piece in re.findall(r"\d+", str(text) or "")][:2]
    if not parts:
        return (0, 0)
    return (parts[0], parts[1] if len(parts) > 1 else 0)


def _link_from_payload(payload: dict) -> LineLink:
    numbers = payload.get("numbers") or ()
    if isinstance(numbers, str):
        numbers = tuple(part.strip() for part in re.split(r"[、,，\s]+", numbers) if part.strip())
    status = str(payload.get("status", "recorded") or "recorded")
    if status not in LINE_STATUSES:
        status = "recorded"
    evidence: list[LineEvidence] = []
    for item in payload.get("evidence") or []:
        if not isinstance(item, dict):
            continue
        evidence.append(
            LineEvidence(
                source=str(item.get("source", "manual") or "manual").strip(),
                bus=str(item.get("bus", "")).strip(),
                start=str(item.get("start", "")).strip(),
                end=str(item.get("end", "")).strip(),
                checked=str(item.get("checked", "")).strip(),
            )
        )
    return LineLink(
        line=str(payload.get("line", "")).strip(),
        model=str(payload.get("model", "")).strip(),
        numbers=tuple(str(item).strip() for item in numbers if str(item).strip()),
        first=str(payload.get("first", "")).strip(),
        last=str(payload.get("last", "")).strip(),
        status=status,
        source=str(payload.get("source", "manual") or "manual"),
        note=str(payload.get("note", "")).strip(),
        evidence=tuple(evidence),
    )


def load_lines_lenient(base_directory: Path) -> tuple[list[LineLink], list[str]]:
    """读线路数据；坏条目跳过并报告，绝不因为一条坏数据就打不开程序。"""
    path = lines_path(base_directory)
    if not path.is_file():
        return [], []
    try:
        content = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        return [], [f"{LINES_FILENAME} 读不出来（{error}），先当作没有线路数据。"]
    if isinstance(content, dict):
        items = content.get("links") or []
    elif isinstance(content, list):
        items = content
    else:
        return [], [f"{LINES_FILENAME} 的结构不对，应该是 {{'links': [...]}}。"]
    links: list[LineLink] = []
    notes: list[str] = []
    for index, item in enumerate(items, start=1):
        if not isinstance(item, dict):
            notes.append(f"第 {index} 条不是对象，已跳过")
            continue
        link = _link_from_payload(item)
        if not link.line:
            notes.append(f"第 {index} 条没有线路名，已跳过")
            continue
        links.append(link)
    return links, notes


def load_lines(base_directory: Path) -> list[LineLink]:
    links, notes = load_lines_lenient(base_directory)
    if notes:
        raise DataError(notes[0])
    return links


def save_lines(base_directory: Path, links: list[LineLink]) -> None:
    """原子写线路数据（失败抛 ProfileError，由调用方提示）。"""
    ordered = sorted(links, key=lambda item: (line_key(item.line), item.model))
    payload = {
        "links": [
            {
                "line": link.line,
                "model": link.model,
                "numbers": list(link.numbers),
                "first": link.first,
                "last": link.last,
                "status": link.status,
                "source": link.source,
                "note": link.note,
                "evidence": [
                    {"source": item.source, "bus": item.bus, "start": item.start,
                     "end": item.end, "checked": item.checked}
                    for item in link.evidence
                ],
            }
            for link in ordered
        ]
    }
    _write_json(lines_path(base_directory), payload, backup=True)


# --------------------------------------------------------------------------- #
# 联网缓存（服务端每个请求约 6 秒，缓存能把重复查询变成 0 秒）
# --------------------------------------------------------------------------- #

def online_cache_path(base_directory: Path) -> Path:
    return base_directory / ONLINE_CACHE_FILENAME


def load_online_cache(base_directory: Path) -> dict:
    """读联网缓存；没有/坏了都返回空表（缓存绝不能挡住程序）。"""
    path = online_cache_path(base_directory)
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def save_online_cache(base_directory: Path, cache: dict) -> None:
    """写联网缓存（不进备份目录：它随时可以重建，失败也不打扰用户）。"""
    _write_json(online_cache_path(base_directory), cache, backup=False)


def cache_entry(cache: dict, key: str, *, days: int, today: str = "") -> dict | None:
    """取一条缓存；超过 days 天就算过期（days=0 表示不过期）。"""
    item = cache.get(key)
    if not isinstance(item, dict):
        return None
    if days <= 0:
        return item
    checked = str(item.get("checked", ""))
    if not _stamp_within(checked, today or datetime.now().strftime("%Y%m%d"), days):
        return None
    return item


def _stamp_within(stamp: str, today: str, days: int) -> bool:
    try:
        then = datetime.strptime(stamp, "%Y%m%d").date()
        now = datetime.strptime(today, "%Y%m%d").date()
    except ValueError:
        return False
    return (now - then).days <= days


def cache_put(cache: dict, key: str, payload, *, checked: str) -> dict:
    """写一条缓存（返回新表，方便链式使用）。"""
    updated = dict(cache)
    updated[key] = {"checked": checked, "payload": payload}
    return updated


def cache_prune(cache: dict, *, days: int, today: str = "") -> dict:
    """清掉过期条目，免得缓存文件无限长大。"""
    if days <= 0:
        return dict(cache)
    stamp = today or datetime.now().strftime("%Y%m%d")
    return {
        key: value
        for key, value in cache.items()
        if isinstance(value, dict) and _stamp_within(str(value.get("checked", "")), stamp, days)
    }


# --------------------------------------------------------------------------- #
# 区属公交（预留）
#
# 北京区属公交和公交集团不是一套编号：每个区/每家公司自己编，有的区甚至只有车牌号没有自编
# （例如昌平是 59 开头的 8 位自编）。这一块**暂时不开发**——手上还没有数据，
# 所以这里只做三件事：定好文件名与字段、保证没有文件时程序照常跑、提供一个空模板。
# 以后拿到数据，把 区属公交数据.json 填上，界面上的「区属公交查询（预留）」就能直接用。
# --------------------------------------------------------------------------- #

DISTRICT_BUS_FILENAME = "区属公交数据.json"

# 编号形态：自编 / 只有车牌 / 两者都有 / 还没编号
DISTRICT_BUS_NUMBER_KINDS = ("自编", "车牌", "自编+车牌", "无编号")


@dataclass(frozen=True)
class DistrictBusCompany:
    """一家区属公交公司的编号规则（各家不一样，所以规则挂在公司上）。"""

    name: str
    district: str = ""
    number_kind: str = "自编"
    number_rule: str = ""      # 给人看的说明，例如「59 开头的 8 位数字」
    number_prefix: str = ""    # 给程序识别用，可空
    number_length: int = 0     # 位数，0 = 不限
    plate_prefix: str = ""     # 只有车牌的公司，例如「京B」（单个，兼容旧数据）
    plate_prefixes: tuple[str, ...] = ()   # 有的区车牌前缀不止一个；留空则用 plate_prefix
    note: str = ""
    # 编号里各段的字符位置（0 起、左闭右开），用于把自编翻译成「年份 / 批次 / 第几辆」
    year_span: tuple[int, int] = ()
    batch_span: tuple[int, int] = ()
    serial_span: tuple[int, int] = ()
    year_base: int = 2000       # 年份那两位按 20xx 解释

    def matches_number(self, text: str) -> bool:
        """这个编号是不是符合本公司的规则（前缀 + 位数）。"""
        value = str(text).strip().upper()
        if not value:
            return False
        if self.number_prefix and not value.startswith(self.number_prefix.upper()):
            return False
        if self.number_length and len(value) != self.number_length:
            return False
        return bool(self.number_prefix or self.number_length)

    @property
    def all_plate_prefixes(self) -> tuple[str, ...]:
        """这家公司认的车牌前缀（plate_prefixes 优先，兼容单个 plate_prefix）。"""
        if self.plate_prefixes:
            return self.plate_prefixes
        return (self.plate_prefix,) if self.plate_prefix else ()

    def matches_plate(self, text: str) -> bool:
        """这个车牌是不是可能属于这家公司（只看前缀；没有前缀规则就不认）。"""
        value = normalize_plate(text)
        prefixes = self.all_plate_prefixes
        if not value or not prefixes:
            return False
        return any(value.startswith(normalize_plate(item)) for item in prefixes)

    def _slice(self, value: str, span: tuple[int, int]) -> str:
        if not span:
            return ""
        start, end = span
        return value[start:end] if len(value) >= end else ""

    def describe(self, number: str) -> dict[str, str]:
        """把自编翻译成各段（规则没配的段留空）。例：59181007 → 2018 年 / 第 1 批 / 第 7 辆。"""
        value = str(number).strip().upper()
        if not self.matches_number(value):
            return {}
        year = self._slice(value, self.year_span)
        batch = self._slice(value, self.batch_span)
        serial = self._slice(value, self.serial_span)
        result: dict[str, str] = {}
        if year.isdigit():
            result["year"] = str(self.year_base + int(year))
        if batch.isdigit():
            result["batch"] = str(int(batch))
        if serial.isdigit():
            result["serial"] = str(int(serial))
        return result


@dataclass(frozen=True)
class DistrictBusRange:
    """一段区属公交自编（含起止）对应的车型；区间之外或空号另行处理。"""

    company: str
    start: int
    end: int
    model: str = ""
    note: str = ""
    # 这一批没编的尾号（例如某年忌「4」，整批都没有尾号 4 的号）+ 原因说明
    skip_tail: tuple[str, ...] = ()
    skip_reason: str = ""

    def contains(self, number: int) -> bool:
        return self.start <= number <= self.end

    def skips(self, number: int) -> bool:
        """这个编号是不是「这一批本来就没编」的（例如尾号 4）。"""
        if not self.skip_tail:
            return False
        return f"{number:08d}"[-1] in self.skip_tail

    @property
    def count(self) -> int:
        return self.end - self.start + 1

    @property
    def skipped_count(self) -> int:
        if not self.skip_tail:
            return 0
        return sum(1 for value in range(self.start, self.end + 1) if self.skips(value))

    @property
    def real_count(self) -> int:
        return self.count - self.skipped_count

    @property
    def label(self) -> str:
        return f"{self.start:08d}-{self.end:08d}" if self.start >= 10_000_000 else f"{self.start}-{self.end}"


@dataclass(frozen=True)
class DistrictBusResult:
    """一次区属公交查询的结果。"""

    query: str
    company: str = ""
    model: str = ""
    status: str = ""       # 已收录 / 该批未编此号 / 空号（未配车）/ 不在已整理的范围里 / 车牌各种情况
    year: str = ""
    batch: str = ""
    serial: str = ""
    span: str = ""
    note: str = ""
    plate: str = ""
    line: str = ""


@dataclass(frozen=True)
class DistrictBusEntry:
    """一辆区属公交：自编和车牌至少填一个，其余都可以空着。"""

    company: str
    number: str = ""
    plate: str = ""
    line: str = ""
    model: str = ""
    year: str = ""
    note: str = ""
    date: str = ""      # 记录日期（坐过那天）

    @property
    def key(self) -> str:
        """去重用：车牌优先，其次自编。"""
        return normalize_plate(self.plate) or str(self.number).strip().upper()

    @property
    def label(self) -> str:
        return normalize_plate(self.plate) or str(self.number).strip()


def district_bus_path(base_directory: Path) -> Path:
    return base_directory / DISTRICT_BUS_FILENAME


def district_bus_entry_cars(entries: list[DistrictBusEntry]) -> list[str]:
    """已经记过的车型（给「记一笔」的输入框做候选）。"""
    seen: list[str] = []
    for entry in entries:
        value = entry.model.strip()
        if value and value not in seen:
            seen.append(value)
    return seen


def upsert_district_bus_entry(
    entries: list[DistrictBusEntry], entry: DistrictBusEntry
) -> tuple[list[DistrictBusEntry], str, str]:
    """加一条 / 改一条记录。

    返回（新表, 发生了什么, 说明）：新增 / 更新 / 冲突。
    同一个车牌被记成**不同车型**算冲突——不悄悄覆盖，交给用户判断。
    """
    if not entry.key:
        return list(entries), "无效", "车牌和自编至少要填一个"
    for index, existing in enumerate(entries):
        if existing.key != entry.key:
            continue
        if existing.model and entry.model and existing.model != entry.model:
            return list(entries), "冲突", f"{entry.label} 已经记成 {existing.model}，这次填的是 {entry.model}"
        merged = DistrictBusEntry(
            company=entry.company or existing.company,
            number=entry.number or existing.number,
            plate=entry.plate or existing.plate,
            line=entry.line or existing.line,
            model=entry.model or existing.model,
            year=entry.year or existing.year,
            note=entry.note or existing.note,
            date=entry.date or existing.date,
        )
        updated = list(entries)
        updated[index] = merged
        return updated, "更新", f"{entry.label} 已更新"
    return list(entries) + [entry], "新增", f"{entry.label} 已记下"


def remove_district_bus_entry(entries: list[DistrictBusEntry], key: str) -> list[DistrictBusEntry]:
    """按车牌 / 自编删掉一条记录。"""
    wanted = normalize_plate(key) or str(key).strip().upper()
    return [item for item in entries if item.key != wanted]


def district_bus_entry_conflicts(entries: list[DistrictBusEntry]) -> list[tuple[str, tuple[str, ...]]]:
    """同一个车牌被记成多个车型的情况（数据攒多了必然会遇到）。"""
    grouped: dict[str, list[str]] = {}
    for entry in entries:
        if not entry.key or not entry.model:
            continue
        grouped.setdefault(entry.key, [])
        if entry.model not in grouped[entry.key]:
            grouped[entry.key].append(entry.model)
    return [(key, tuple(models)) for key, models in grouped.items() if len(models) > 1]


def district_bus_entry_summary(entries: list[DistrictBusEntry]) -> list[tuple[str, int]]:
    """按车型汇总（车型, 车牌数），多的在前——攒起来之后能整理出的东西。"""
    counted: dict[str, set[str]] = {}
    for entry in entries:
        model = entry.model.strip() or "（车型未填）"
        counted.setdefault(model, set()).add(entry.key or entry.label)
    return sorted(((model, len(keys)) for model, keys in counted.items()), key=lambda row: (-row[1], row[0]))


# 车牌：北京是「京」+ 字母 + 5 位（这里只做统一化，不校验真伪——区属公交的车牌数据以后才有）
PLATE_PATTERN = re.compile(r"^[\u4e00-\u9fa5][A-Z][A-Z0-9]{4,6}$")


def normalize_plate(text: str) -> str:
    """车牌统一化：去掉空格/间隔点/横线，转半角大写。京B·12345 → 京B12345。"""
    value = str(text).strip().upper()
    for junk in (" ", "\u3000", "·", "・", ".", "-", "_"):
        value = value.replace(junk, "")
    return value


def looks_like_plate(text: str) -> bool:
    """像不像车牌（汉字 + 字母 + 数字），用来判断用户输入的是车牌还是自编。"""
    return bool(PLATE_PATTERN.match(normalize_plate(text)))


def load_district_bus(
    base_directory: Path,
) -> tuple[list[DistrictBusCompany], list[DistrictBusEntry], list[DistrictBusRange], tuple[str, ...]]:
    """读区属公交数据；**没有文件就返回空表**（程序照常跑）。

    返回（公司, 单车记录, 号段区间, 空号）。
    """
    path = district_bus_path(base_directory)
    if not path.is_file():
        return [], [], [], ()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return [], [], [], ()
    if not isinstance(payload, dict):
        return [], [], [], ()

    def span(value: object) -> tuple[int, int]:
        if isinstance(value, (list, tuple)) and len(value) == 2:
            try:
                return int(value[0]), int(value[1])
            except (TypeError, ValueError):
                return ()
        return ()

    companies: list[DistrictBusCompany] = []
    for row in payload.get("companies") or []:
        if not isinstance(row, dict) or not str(row.get("name", "")).strip():
            continue
        try:
            length = int(row.get("number_length") or 0)
        except (TypeError, ValueError):
            length = 0
        try:
            year_base = int(row.get("year_base") or 2000)
        except (TypeError, ValueError):
            year_base = 2000
        companies.append(
            DistrictBusCompany(
                name=str(row.get("name", "")).strip(),
                district=str(row.get("district", "")).strip(),
                number_kind=str(row.get("number_kind", "")).strip() or "自编",
                number_rule=str(row.get("number_rule", "")).strip(),
                number_prefix=str(row.get("number_prefix", "")).strip(),
                number_length=length,
                plate_prefix=str(row.get("plate_prefix", "")).strip(),
                plate_prefixes=tuple(
                    str(item).strip() for item in (row.get("plate_prefixes") or []) if str(item).strip()
                ),
                note=str(row.get("note", "")).strip(),
                year_span=span(row.get("year_span")),
                batch_span=span(row.get("batch_span")),
                serial_span=span(row.get("serial_span")),
                year_base=year_base,
            )
        )

    ranges: list[DistrictBusRange] = []
    for row in payload.get("ranges") or []:
        if not isinstance(row, dict):
            continue
        try:
            start = int(str(row.get("start", "")).strip())
            end = int(str(row.get("end", "")).strip())
        except (TypeError, ValueError):
            continue
        if end < start:
            start, end = end, start
        ranges.append(
            DistrictBusRange(
                company=str(row.get("company", "")).strip(),
                start=start,
                end=end,
                model=str(row.get("model", "")).strip(),
                note=str(row.get("note", "")).strip(),
                skip_tail=tuple(
                    str(item).strip() for item in (row.get("skip_tail") or []) if str(item).strip()
                ),
                skip_reason=str(row.get("skip_reason", "")).strip(),
            )
        )

    entries: list[DistrictBusEntry] = []
    for row in payload.get("entries") or []:
        if not isinstance(row, dict):
            continue
        entry = DistrictBusEntry(
            company=str(row.get("company", "")).strip(),
            number=str(row.get("number", "")).strip(),
            plate=str(row.get("plate", "")).strip(),
            line=str(row.get("line", "")).strip(),
            model=str(row.get("model", "")).strip(),
            year=str(row.get("year", "")).strip(),
            note=str(row.get("note", "")).strip(),
            date=str(row.get("date", "")).strip(),
        )
        if entry.number or entry.plate:   # 公司可空：自己坐一次记一条时常常不知道是哪家公司的
            entries.append(entry)

    empty = tuple(str(item).strip() for item in (payload.get("empty_numbers") or []) if str(item).strip())
    return companies, entries, ranges, empty


def district_bus_lookup(
    text: str,
    companies: list[DistrictBusCompany],
    ranges: list[DistrictBusRange],
    empty_numbers: tuple[str, ...] = (),
) -> DistrictBusResult:
    """查一个区属公交自编：是哪家公司、哪年哪批第几辆、什么车型。"""
    value = str(text).strip().upper()
    if not value:
        return DistrictBusResult(query=value, status="请输入自编")
    digits = value.zfill(8) if value.isdigit() and len(value) < 8 else value
    company = next((item for item in companies if item.matches_number(digits)), None)
    described = company.describe(digits) if company else {}
    if digits in {item.upper() for item in empty_numbers}:
        return DistrictBusResult(
            query=value,
            company=company.name if company else "",
            status="空号（未配车）",
            year=described.get("year", ""),
            batch=described.get("batch", ""),
            serial=described.get("serial", ""),
        )
    number = int(digits) if digits.isdigit() else -1
    span = next((item for item in ranges if item.contains(number)), None)
    if span is None:
        status = "不在已整理的范围里" if company else "看不出是哪家（编号规则没对上）"
        return DistrictBusResult(
            query=value,
            company=company.name if company else "",
            status=status,
            year=described.get("year", ""),
            batch=described.get("batch", ""),
            serial=described.get("serial", ""),
        )
    if span.skips(number):
        # 这一批本来就没编这个号（例如某年忌「4」，整批没有尾号 4）——不是空号，是没这个号
        return DistrictBusResult(
            query=value,
            company=span.company or (company.name if company else ""),
            status="该批未编此号",
            year=described.get("year", ""),
            batch=described.get("batch", ""),
            serial=described.get("serial", ""),
            span=span.label,
            note=span.skip_reason,
        )
    return DistrictBusResult(
        query=value,
        company=span.company or (company.name if company else ""),
        model=span.model,
        status="已收录",
        year=described.get("year", ""),
        batch=described.get("batch", ""),
        serial=described.get("serial", ""),
        span=span.label,
        note=span.note or (company.note if company else ""),
    )


def district_bus_lookup_plate(
    text: str,
    companies: list[DistrictBusCompany],
    entries: list[DistrictBusEntry],
) -> DistrictBusResult:
    """按车牌查区属公交（**预留**：其他区基本只有车牌没有自编）。

    思路：车牌记在 entries[].plate 上，公司只提供前缀线索。车牌前缀各区会重（一个「京B」
    下面可能有好几家），所以**不用前缀猜公司**，直接按记录里的车牌查；查不到就照实说。
    """
    value = normalize_plate(text)
    if not value:
        return DistrictBusResult(query=str(text).strip(), status="请输入车牌")
    matched = [item for item in entries if normalize_plate(item.plate) == value]
    if not matched:
        with_plates = [item for item in entries if item.plate]
        if not with_plates:
            # 连一条车牌数据都还没有：说明这块还没开始录
            hinted = [company.name for company in companies if company.matches_plate(value)]
            return DistrictBusResult(
                query=value,
                company="、".join(hinted),
                status="还没有车牌数据",
                note="其他区目前基本只有车牌号、没有自编，车牌数据等以后录进来就能查",
            )
        return DistrictBusResult(query=value, status="这个车牌没记录过")
    entry = matched[0]
    return DistrictBusResult(
        query=value,
        company=entry.company,
        model=entry.model,
        status="已收录",
        line=entry.line,
        plate=entry.plate,
        year=entry.year,
        note=entry.note,
        span="、" .join(sorted({item.plate for item in matched})) if len(matched) > 1 else "",
    )


def district_bus_lookup_any(
    text: str,
    companies: list[DistrictBusCompany],
    ranges: list[DistrictBusRange],
    empty_numbers: tuple[str, ...] = (),
    entries: list[DistrictBusEntry] | None = None,
) -> DistrictBusResult:
    """一个入口同时查自编和车牌：像车牌就走车牌，否则当自编查。"""
    if looks_like_plate(text):
        return district_bus_lookup_plate(text, companies, entries or [])
    return district_bus_lookup(text, companies, ranges, empty_numbers)


def district_bus_districts(
    companies: list[DistrictBusCompany], ranges: list[DistrictBusRange], entries: list[DistrictBusEntry]
) -> list[tuple[str, str, int]]:
    """各区县的数据情况：（区县, 情况说明, 记录条数），用来在界面上显示"哪些区还空着"。"""
    counts: dict[str, int] = {}
    for item in ranges:
        company = next((c for c in companies if c.name == item.company), None)
        district = company.district if company else ""
        counts[district] = counts.get(district, 0) + 1
    for entry in entries:
        company = next((c for c in companies if c.name == entry.company), None)
        district = company.district if company else ""
        counts[district] = counts.get(district, 0) + 1
    rows: list[tuple[str, str, int]] = []
    for company in companies:
        district = company.district or "（未填区县）"
        kinds: list[str] = []
        if any(item.company == company.name for item in ranges):
            kinds.append("有自编号段")
        if company.all_plate_prefixes:
            kinds.append("按车牌（" + "、".join(company.all_plate_prefixes) + "）")
        if any(item.company == company.name and item.plate for item in entries):
            kinds.append("有车牌记录")
        rows.append((district, " · ".join(kinds) or "有公司但还没有数据", counts.get(district, 0)))
    return rows


@dataclass(frozen=True)
class DistrictBusBatchPlan:
    """同一车型批量录车牌的预演结果（先看清楚再写）。"""

    added: tuple[str, ...] = ()            # 会新增的车牌
    refreshed: tuple[str, ...] = ()        # 库里已有、同车型（只补空字段）
    conflicts: tuple[tuple[str, str], ...] = ()   # （车牌, 库里已记的车型）不同车型，不动
    duplicates: tuple[str, ...] = ()       # 这一批里自己重复的
    invalid: tuple[str, ...] = ()          # 认不出是车牌的

    @property
    def total_ok(self) -> int:
        return len(self.added) + len(self.refreshed)


def district_bus_split_plates(text: str) -> list[str]:
    """从一段文本里抠出车牌：换行、空格、逗号、顿号、制表符都能当分隔。"""
    parts = [part for part in re.split(r"[\s,，、;；\t]+", str(text)) if part]
    return [normalize_plate(part) for part in parts if looks_like_plate(part)]


def district_bus_plan_batch(
    entries: list[DistrictBusEntry], plates: list[str], model: str, line: str = "", date: str = ""
) -> DistrictBusBatchPlan:
    """预演「同一车型批量录车牌」：哪些会新增、哪些已有、哪些冲突、哪些重复。"""
    model = model.strip()
    known = {item.key: item.model for item in entries}
    added: list[str] = []
    refreshed: list[str] = []
    conflicts: list[tuple[str, str]] = []
    duplicates: list[str] = []
    invalid: list[str] = []
    seen: set[str] = set()
    for plate in plates:
        value = normalize_plate(plate)
        if not looks_like_plate(value):
            invalid.append(str(plate))
            continue
        if value in seen:
            duplicates.append(value)
            continue
        seen.add(value)
        if value not in known:
            added.append(value)
        elif known[value] and model and known[value] != model:
            conflicts.append((value, known[value]))
        else:
            refreshed.append(value)
    return DistrictBusBatchPlan(
        added=tuple(added),
        refreshed=tuple(refreshed),
        conflicts=tuple(conflicts),
        duplicates=tuple(duplicates),
        invalid=tuple(invalid),
    )


def district_bus_apply_batch(
    entries: list[DistrictBusEntry], plates: list[str], model: str, line: str = "", date: str = ""
) -> tuple[list[DistrictBusEntry], DistrictBusBatchPlan]:
    """按预演结果写入：新增 + 更新（同车型），冲突与重复的不动。"""
    plan = district_bus_plan_batch(entries, plates, model, line, date)
    updated = list(entries)
    for plate in plan.added + plan.refreshed:
        updated, _what, _message = upsert_district_bus_entry(
            updated, DistrictBusEntry(company="", plate=plate, model=model, line=line, date=date)
        )
    return updated, plan


def district_bus_duplicate_keys(entries: list[DistrictBusEntry]) -> list[tuple[str, int]]:
    """同一个车牌/自编出现了多条（手动改 JSON、或从别处导入时会出现）。"""
    counted: dict[str, int] = {}
    for entry in entries:
        if entry.key:
            counted[entry.key] = counted.get(entry.key, 0) + 1
    return sorted(((key, count) for key, count in counted.items() if count > 1), key=lambda row: (-row[1], row[0]))


def district_bus_merge_duplicates(
    entries: list[DistrictBusEntry],
) -> tuple[list[DistrictBusEntry], int]:
    """把同一个车牌的多条合并成一条（字段取并集，车型冲突时保留第一条并记进备注）。

    没有车牌也没有自编的"废记录"原样保留（清理重复不该顺手删数据）。
    """
    merged: dict[str, DistrictBusEntry] = {}
    order: list[str] = []
    keyless: list[DistrictBusEntry] = []
    for entry in entries:
        if not entry.key:
            keyless.append(entry)
            continue
        if entry.key not in merged:
            merged[entry.key] = entry
            order.append(entry.key)
            continue
        first = merged[entry.key]
        note = first.note
        if entry.model and first.model and entry.model != first.model:
            note = (note + f"；另记作 {entry.model}").strip("；")
        merged[entry.key] = DistrictBusEntry(
            company=first.company or entry.company,
            number=first.number or entry.number,
            plate=first.plate or entry.plate,
            line=first.line or entry.line,
            model=first.model or entry.model,
            year=first.year or entry.year,
            note=note,
            date=max(first.date, entry.date),  # 日期取新的（字符串按 YYYY.M.D 可直接比）
        )
    result = [merged[key] for key in order] + keyless
    return result, len(entries) - len(result)


def district_bus_set_fields(
    entries: list[DistrictBusEntry],
    keys: list[str] | tuple[str, ...],
    *,
    company: str | None = None,
    line: str | None = None,
    model: str | None = None,
) -> tuple[list[DistrictBusEntry], int]:
    """批量改选中的几条记录的公司 / 线路 / 车型（只改传进来的项，其它不动）。

    返回（新表, 改了几条）。传 None 表示这一项不动；传空字符串表示清空这一项。
    """
    wanted = {normalize_plate(key) or str(key).strip().upper() for key in keys}
    if not wanted:
        return list(entries), 0
    result: list[DistrictBusEntry] = []
    changed = 0
    for entry in entries:
        if entry.key not in wanted:
            result.append(entry)
            continue
        updated = DistrictBusEntry(
            company=entry.company if company is None else company.strip(),
            number=entry.number,
            plate=entry.plate,
            line=entry.line if line is None else line.strip(),
            model=entry.model if model is None else model.strip(),
            year=entry.year,
            note=entry.note,
            date=entry.date,
        )
        if updated != entry:
            changed += 1
        result.append(updated)
    return result, changed


def district_bus_companies_seen(entries: list[DistrictBusEntry]) -> list[str]:
    """已经填过的公司 / 区县（给批量设置做候选）。"""
    seen: list[str] = []
    for entry in entries:
        value = entry.company.strip()
        if value and value not in seen:
            seen.append(value)
    return sorted(seen)


def district_bus_quality(entries: list[DistrictBusEntry]) -> dict[str, list]:
    """车牌记录的自查结果：重复 / 同车牌不同车型 / 缺车型 / 缺车牌或自编。"""
    return {
        "duplicates": district_bus_duplicate_keys(entries),
        "conflicts": district_bus_entry_conflicts(entries),
        # 缺车型只算"有车牌/自编"的记录；连车牌都没有的已经在 no_key 里报过了
        "no_model": [item.label for item in entries if item.key and not item.model.strip()],
        "no_key": [item.label for item in entries if not item.key],
    }


def save_district_bus(
    base_directory: Path,
    companies: list[DistrictBusCompany],
    ranges: list[DistrictBusRange],
    empty_numbers: tuple[str, ...] = (),
    entries: list[DistrictBusEntry] | None = None,
) -> None:
    """原子写区属公交数据（自动备份）。"""

    def span_text(value: tuple[int, int]) -> list[int]:
        return [value[0], value[1]] if value else []

    payload = {
        "说明": "区属公交数据：companies 存各家公司的编号规则，ranges 存号段与车型，"
                "empty_numbers 存空号，entries 存零散单车记录。",
        "companies": [
            {
                "name": item.name,
                "district": item.district,
                "number_kind": item.number_kind,
                "number_rule": item.number_rule,
                "number_prefix": item.number_prefix,
                "number_length": item.number_length,
                "plate_prefix": item.plate_prefix,
                "plate_prefixes": list(item.plate_prefixes),
                "year_span": span_text(item.year_span),
                "batch_span": span_text(item.batch_span),
                "serial_span": span_text(item.serial_span),
                "year_base": item.year_base,
                "note": item.note,
            }
            for item in companies
        ],
        "ranges": [
            {"company": item.company, "start": f"{item.start:08d}", "end": f"{item.end:08d}",
             "model": item.model, "note": item.note,
             "skip_tail": list(item.skip_tail), "skip_reason": item.skip_reason}
            for item in sorted(ranges, key=lambda row: (row.company, row.start))
        ],
        "empty_numbers": sorted(empty_numbers),
        "entries": [
            {
                "company": item.company,
                "number": item.number,
                "plate": item.plate,
                "line": item.line,
                "model": item.model,
                "year": item.year,
                "note": item.note,
                "date": item.date,
            }
            for item in (entries or [])
        ],
    }
    _write_json(district_bus_path(base_directory), payload, backup=True)


def district_bus_companies_for(text: str, companies: list[DistrictBusCompany]) -> list[DistrictBusCompany]:
    """哪些公司的编号规则能对上这个编号。"""
    return [company for company in companies if company.matches_number(text)]


def district_bus_template() -> dict:
    """空模板：放一家示例公司（昌平 59 开头 8 位的规则），供手工填数据时参照。"""
    return {
        "说明": "区属公交数据。companies 存各家公司的编号规则（编号形态、前缀、位数、"
                "以及年份/批次/序号各占哪几位），ranges 存号段对应的车型，"
                "empty_numbers 存空号，entries 存零散单车记录。字段说明见 使用说明.md。",
        "companies": [
            {
                "name": "某某客运有限公司",
                "district": "昌平",
                "number_kind": "自编",
                "number_rule": "8 位：前两位是公司缩写，第 3-4 位是购车年份，第 5 位是购车批次，后三位是当年购买的第几辆车",
                "number_prefix": "59",
                "number_length": 8,
                "plate_prefix": "",
                "year_span": [2, 4],
                "batch_span": [4, 5],
                "serial_span": [5, 8],
                "year_base": 2000,
                "note": "示例行（有自编的公司），按实际情况改",
            },
            {
                "name": "某某区公交（只有车牌、没有自编）",
                "district": "某某区",
                "number_kind": "车牌",
                "number_rule": "没有自编，只有车牌号",
                "number_prefix": "",
                "number_length": 0,
                "plate_prefix": "",
                "plate_prefixes": ["京B", "京G"],
                "note": "示例行（只有车牌的公司）：车记在 entries 里，填 plate/line/model；"
                        "查询框里直接输车牌就能查",
            },
        ],
        "ranges": [],
        "empty_numbers": [],
        "entries": [],
    }


def save_district_bus_template(base_directory: Path) -> Path:
    """把空模板写到数据目录（已存在就不覆盖）。"""
    path = district_bus_path(base_directory)
    if path.exists():
        return path
    _write_json(path, district_bus_template(), backup=False)
    return path


@dataclass(frozen=True)
class LinePlan:
    """线路数据合并后的差异。"""

    added: list[LineLink]
    updated: list[tuple[LineLink, LineLink]]  # (新, 旧)
    removed: list[LineLink]

    @property
    def changed(self) -> int:
        return len(self.added) + len(self.updated) + len(self.removed)


def _merge_link(old: LineLink, new: LineLink) -> LineLink:
    """同一 (线路, 车型) 的两条记录合并：时间取范围并集，编号取并集。

    **来源按"谁更算数"取**（手工 > 导入 > 网上查的）：网上查到的数据不许把
    「我记过」改写成「网上查的」——那等于把用户自己的记录抹掉了。
    """
    numbers = tuple(sorted(set(old.numbers) | set(new.numbers)))
    firsts = [value for value in (old.first, new.first) if value]
    lasts = [value for value in (old.last, new.last) if value]
    status = new.status
    if old.status == "current" and new.status != "current":
        status = "current"  # 已经确认过"现在还在跑"的，不要被别的来源降级
    evidence: list[LineEvidence] = []
    seen: set[tuple[str, str, str, str]] = set()
    for item in tuple(old.evidence) + tuple(new.evidence):
        if item.key in seen:
            continue
        seen.add(item.key)
        evidence.append(item)
    evidence.sort(key=lambda item: (item.start, item.source, item.bus))
    return LineLink(
        line=new.line or old.line,
        model=new.model or old.model,
        numbers=numbers,
        first=min(firsts, key=month_key) if firsts else "",
        last=max(lasts, key=month_key) if lasts else "",
        status=status,
        source=old.source if source_rank(old.source) <= source_rank(new.source) else new.source,
        note=new.note or old.note,
        evidence=tuple(evidence),
    )


def merge_lines(current: list[LineLink], incoming: list[LineLink]) -> LinePlan:
    """把导入的线路记录并进现有数据（同一 线路+车型 视为同一条）。"""
    index = {(line_key(item.line), item.model): item for item in current}
    added: list[LineLink] = []
    updated: list[tuple[LineLink, LineLink]] = []
    for item in incoming:
        key = (line_key(item.line), item.model)
        old = index.get(key)
        if old is None:
            index[key] = item
            added.append(item)
        else:
            merged = _merge_link(old, item)
            if merged != old:
                index[key] = merged
                updated.append((merged, old))
    removed = [item for item in current if (line_key(item.line), item.model) not in index]
    return LinePlan(added, updated, removed)


def delete_by_source(links: list[LineLink], source: str) -> list[LineLink]:
    """按来源删（例如"把网上查的全删掉"）；手工记录不受影响。"""
    wanted = str(source).strip()
    return [item for item in links if item.source != wanted]


def online_stale(links: list[LineLink], *, days: int = 30, today: str = "") -> list[LineLink]:
    """列出"网上查的、而且已经很久没更新"的记录（提示可以刷新）。"""
    stamp = today or datetime.now().strftime("%Y%m%d")
    stale: list[LineLink] = []
    for link in links:
        if link.source != "online":
            continue
        checked = [item.checked for item in link.evidence if item.checked]
        if not checked:
            stale.append(link)
            continue
        if not _stamp_within(max(checked), stamp, days):
            stale.append(link)
    return stale


def lines_from_history(
    bus_id: str,
    history: dict[str, list[tuple[str, str]]],
    model_of: Callable[[str], str] | None = None,
    *,
    checked: str = "",
    current_within_days: int = 7,
    today: str = "",
) -> list[LineLink]:
    """把一辆车的运营历史变成线路记录（含证据链）——联网写入的唯一入口。

    与 `bpt_online.history_to_links` 的区别：这里**同时写 `evidence`**，
    记录下"哪辆车、哪段日期、哪天查的"，以后能分清哪些是查来的、哪些是自己记的。
    """
    stamp = today or datetime.now().strftime("%Y%m%d")
    number = str(bus_id).strip()
    model = (model_of(number) if model_of else "") or ""
    links: list[LineLink] = []
    for line, ranges in history.items():
        cleaned: list[tuple[str, str]] = []
        for item in ranges:
            if isinstance(item, (tuple, list)) and len(item) == 2:
                cleaned.append((str(item[0]), str(item[1])))
            elif isinstance(item, str) and "-" in item:
                start, _, end = item.partition("-")
                if start.strip().isdigit() and end.strip().isdigit():
                    cleaned.append((start.strip(), end.strip()))
        if not cleaned:
            continue
        starts = [start for start, _end in cleaned]
        ends = [end for _start, end in cleaned]
        latest = max(ends)
        status = "current" if _stamp_within(latest, stamp, current_within_days) else "recorded"
        evidence = tuple(
            LineEvidence(source="online", bus=number, start=start, end=end, checked=checked or stamp)
            for start, end in sorted(cleaned)
        )
        links.append(
            LineLink(
                line=line,
                model=model,
                numbers=(number,) if number else (),
                first=_compact_month(min(starts)),
                last=_compact_month(latest),
                status=status,
                source="online",
                note=f"网上查的（{checked or stamp}）",
                evidence=evidence,
            )
        )
    return sorted(links, key=lambda item: (month_key(item.first), item.line))


def _compact_month(stamp: str) -> str:
    value = str(stamp).strip()
    if len(value) == 8 and value.isdigit():
        return f"{value[:4]}.{int(value[4:6])}"
    return value


def _clean_ranges(ranges) -> list[tuple[str, str]]:
    """把「YYYYMMDD-YYYYMMDD」字符串或 (start, end) 元组统一成有序元组列表。"""
    cleaned: list[tuple[str, str]] = []
    for item in ranges or ():
        if isinstance(item, (tuple, list)) and len(item) == 2:
            pair = (str(item[0]).strip(), str(item[1]).strip())
        elif isinstance(item, str) and "-" in item:
            start, _sep, end = item.partition("-")
            pair = (start.strip(), end.strip())
        else:
            continue
        if len(pair[0]) == 8 and len(pair[1]) == 8 and pair[0].isdigit() and pair[1].isdigit():
            cleaned.append(pair)
    return sorted(set(cleaned))


def lines_from_route_history(
    route_number: str,
    bus_history: dict,
    model_of: Callable[[str], str] | None = None,
    *,
    checked: str = "",
    current_within_days: int = 7,
    today: str = "",
) -> list[LineLink]:
    """「这条线用过哪些车」→ 线路记录（按车型归并，一条线路一个车型一条记录）。

    与"按车辆查"的区别：这里 numbers 会是**一批自编号**（同型号的），
    证据只记一条总体区间——因为一条线可能用过 300 多辆车，逐车写证据会把文件撑爆。
    逐车的日期明细留在界面里看，不落盘。
    """
    stamp = today or datetime.now().strftime("%Y%m%d")
    line = str(route_number).strip()
    grouped: dict[str, dict] = {}
    for bus_id, ranges in (bus_history or {}).items():
        cleaned = _clean_ranges(ranges)
        if not cleaned:
            continue
        model = (model_of(str(bus_id)) if model_of else "") or ""
        entry = grouped.setdefault(model, {"numbers": [], "starts": [], "ends": []})
        entry["numbers"].append(str(bus_id).strip())
        entry["starts"].append(cleaned[0][0])
        entry["ends"].append(cleaned[-1][1])
    links: list[LineLink] = []
    for model, entry in grouped.items():
        if not entry["numbers"]:
            continue
        latest = max(entry["ends"])
        status = "current" if _stamp_within(latest, stamp, current_within_days) else "recorded"
        links.append(
            LineLink(
                line=line,
                model=model,
                numbers=tuple(sorted(entry["numbers"])),
                first=_compact_month(min(entry["starts"])),
                last=_compact_month(latest),
                status=status,
                source="online",
                note=f"网上查的：{line} 用过的车（{len(entry['numbers'])} 辆，{checked or stamp} 查）",
                evidence=(
                    LineEvidence(
                        source="online", bus="", start=min(entry["starts"]), end=latest, checked=checked or stamp
                    ),
                ),
            )
        )
    return sorted(links, key=lambda item: (month_key(item.first), item.model))


# --------------------------------------------------------------------------- #
# 线路资料（首末站 / 站点表；来源同样是联网，属于用户自己的数据）
# --------------------------------------------------------------------------- #

def line_info_path(base_directory: Path) -> Path:
    return base_directory / LINE_INFO_FILENAME


def load_line_info(base_directory: Path) -> dict:
    """读线路资料；没有/坏了都返回空表（缺了不该挡住任何功能）。"""
    path = line_info_path(base_directory)
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    lines = payload.get("lines")
    return lines if isinstance(lines, dict) else {}


def save_line_info(base_directory: Path, lines: dict) -> None:
    """原子写线路资料（自动备份）。"""
    _write_json(line_info_path(base_directory), {"lines": lines}, backup=True)


ONLINE_LOG_FILENAME = "联网日志.txt"
ONLINE_LOG_KEEP = 300


def online_log_path(base_directory: Path) -> Path:
    return base_directory / ONLINE_LOG_FILENAME


def append_online_log(base_directory: Path, line: str) -> None:
    """给联网请求记一行日志（写不进去就悄悄算了，绝不因为日志挡住查询）。"""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path = online_log_path(base_directory)
    try:
        existing = path.read_text(encoding="utf-8").splitlines() if path.is_file() else []
        existing.append(f"{stamp}\t{line}")
        path.write_text("\n".join(existing[-ONLINE_LOG_KEEP:]) + "\n", encoding="utf-8")
    except OSError:
        pass


def read_online_log(base_directory: Path, *, limit: int = 50) -> list[str]:
    path = online_log_path(base_directory)
    if not path.is_file():
        return []
    try:
        return path.read_text(encoding="utf-8").splitlines()[-limit:]
    except OSError:
        return []


def clear_online_log(base_directory: Path) -> bool:
    path = online_log_path(base_directory)
    try:
        path.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def recent_online_errors(base_directory: Path, *, limit: int = 10) -> list[str]:
    return [line for line in read_online_log(base_directory, limit=ONLINE_LOG_KEEP) if "失败" in line][-limit:]


def parse_schedules(payload) -> list[tuple[str, str, str, str, str]]:
    """发车时刻 → [(自编号, 发车时间, 发车站, 到达时间, 到达站)]。"""
    if isinstance(payload, dict):
        payload = payload.get("result") or []
    if not isinstance(payload, list):
        return []
    rows: list[tuple[str, str, str, str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        depart = item.get("depart") or {}
        arrive = item.get("arrive") or {}
        if not isinstance(depart, dict) or not isinstance(arrive, dict):
            continue
        rows.append(
            (
                str(item.get("bus_id", "")).strip(),
                str(depart.get("time", "")).strip(),
                str(depart.get("stop_name", "")).strip(),
                str(arrive.get("time", "")).strip(),
                str(arrive.get("stop_name", "")).strip(),
            )
        )
    return rows


def parse_depots(payload) -> list[tuple[str, str, str, str]]:
    """出车记录（场站）→ [(自编号, 时间, 位置/站名, 方向)]。"""
    if isinstance(payload, dict):
        payload = payload.get("result") or []
    if not isinstance(payload, list):
        return []
    rows: list[tuple[str, str, str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        rows.append(
            (
                str(item.get("bus_id", "")).strip(),
                str(item.get("time", "")).strip(),
                str(item.get("stop_name", "")).strip(),
                str(item.get("direction_to", "")).strip(),
            )
        )
    return rows


def parse_routes_at_stop(payload) -> list[tuple[str, str, str, str]]:
    """「这个站有哪些线」→ [(线路号, 首站, 末站, 这是第几站)]。"""
    if isinstance(payload, dict):
        payload = payload.get("result") or []
    if not isinstance(payload, list):
        return []
    rows: list[tuple[str, str, str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        rows.append(
            (
                str(item.get("route_number", "")).strip(),
                str(item.get("route_first_stop", "")).strip(),
                str(item.get("route_last_stop", "")).strip(),
                str(item.get("current_stop_number", "")).strip(),
            )
        )
    return rows


def parse_routes_between(payload) -> list[tuple[str, str, str, str, str]]:
    """「从 A 到 B 直达」→ [(线路号, 首站, 末站, 上车站序号, 下车站序号)]。"""
    if isinstance(payload, dict):
        payload = payload.get("result") or []
    if not isinstance(payload, list):
        return []
    rows: list[tuple[str, str, str, str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        rows.append(
            (
                str(item.get("route_number", "")).strip(),
                str(item.get("route_first_stop", "")).strip(),
                str(item.get("route_last_stop", "")).strip(),
                str(item.get("from_stop_number", "")).strip(),
                str(item.get("to_stop_number", "")).strip(),
            )
        )
    return rows


def parse_timetable(text: str) -> list[tuple[str, str, str, str]]:
    """把服务端的时刻表（制表符表格字符串）解析成 (车号, 始发, 终点, 站数)。

    原始格式：第一行是表头（车号 + 各站名），之后每行是「车号 + 每站"时间,拥挤度,距离"」。
    表格太宽，界面上只给概要；要细节可以导出原始文本。
    """
    lines = [line for line in str(text).splitlines() if line.strip()]
    if len(lines) < 2:
        return []
    rows: list[tuple[str, str, str, str]] = []
    for line in lines[1:]:
        cells = line.split("\t")
        bus = cells[0].strip()
        times: list[str] = []
        for cell in cells[1:]:
            piece = cell.split(",")[0].strip()
            if piece:
                times.append(piece)
        if not bus or not times:
            continue
        rows.append((bus, times[0], times[-1], str(len(times))))
    return rows


def parse_track_latest(payload) -> dict:
    """某车轨迹（很长）→ 最后一条 = 它现在跑到哪儿了。"""
    if isinstance(payload, dict):
        payload = payload.get("result") or []
    if not isinstance(payload, list) or not payload:
        return {}
    last = payload[-1]
    if not isinstance(last, dict):
        return {}
    return {
        "time": str(last.get("time", "")),
        "line_name": str(last.get("line_name", "")),
        "stop_name": str(last.get("stop_name", "")),
        "distance_m": last.get("distance_m"),
        "arrived": bool(last.get("arrived", False)),
        "occupancy": last.get("occupancy"),
        "points": str(len(payload)),
    }


def parse_arrivals(payload) -> list[tuple[str, str, str]]:
    """到站记录 → [(自编号, 时间, 拥挤度)]。"""
    if isinstance(payload, dict):
        payload = payload.get("result") or []
    if not isinstance(payload, list):
        return []
    rows: list[tuple[str, str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        rows.append(
            (
                str(item.get("bus_id", "")).strip(),
                str(item.get("time", "")).strip(),
                str(item.get("occupancy", "") or ""),
            )
        )
    return rows


def parse_route_list(payload) -> dict[str, tuple[str, str, str]]:
    """服务端全量线路目录 → {内部id: (线路号, 首站, 末站)}。

    注意：服务端返回的**内部 id 只对当前数据集有效**，所以这份目录只用来"查线路号在不在"、
    "首末站是什么"，**不要把 id 当稳定标识存进你自己的数据**。
    """
    if isinstance(payload, dict) and "result" in payload and isinstance(payload["result"], dict):
        payload = payload["result"]
    if not isinstance(payload, dict):
        return {}
    routes: dict[str, tuple[str, str, str]] = {}
    for route_id, item in payload.items():
        if not isinstance(item, dict):
            continue
        number = str(item.get("number", "")).strip()
        if not number:
            continue
        routes[str(route_id)] = (number, str(item.get("from", "")).strip(), str(item.get("to", "")).strip())
    return routes


def route_list_numbers(routes: dict) -> dict[str, list[tuple[str, str]]]:
    """线路号 → [(首站, 末站)]（同一线路号通常有上下行两条）。"""
    index: dict[str, list[tuple[str, str]]] = {}
    for _route_id, (number, first_stop, last_stop) in (routes or {}).items():
        pair = (first_stop, last_stop)
        bucket = index.setdefault(number, [])
        if pair not in bucket:
            bucket.append(pair)
    return {number: sorted(pairs) for number, pairs in index.items()}


def route_list_suggest(index: dict[str, list[tuple[str, str]]], query: str, *, limit: int = 8) -> list[str]:
    """按"前缀 → 包含 → 只差一个字符"猜线路号（打错号时给个提示）。

    为什么要有第三层：用户手打线路号最常见的错是**错一个字符**（887 想打 886），
    只按前缀/包含找不出来，等于没帮上忙。每一层内部按（长度, 名字）排，层与层保持先后。
    """
    text = str(query or "").strip().casefold()
    if not text:
        return []
    names = sorted(index, key=lambda value: (len(value), value))
    tiers = [
        [name for name in names if name.casefold().startswith(text)],
        [name for name in names if text in name.casefold()],
        [
            name for name in names
            if len(text) >= 2 and len(name) == len(text)
            and name.casefold()[:1] == text[:1]          # 首字符得一样，否则单字查询会匹配一大堆
            and sum(a != b for a, b in zip(name.casefold(), text)) == 1
        ],
    ]
    hits: list[str] = []
    for tier in tiers:
        for name in tier:
            if name not in hits:
                hits.append(name)
            if len(hits) >= limit:
                return hits
    return hits


def route_list_path(base_directory: Path) -> Path:
    return base_directory / ROUTE_LIST_FILENAME


def load_route_list(base_directory: Path) -> dict:
    """读本地线路目录（{内部id: [线路号, 首站, 末站]}）；坏了就当没有。"""
    path = route_list_path(base_directory)
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    routes = payload.get("routes")
    checked = str(payload.get("checked", ""))
    if not isinstance(routes, dict):
        return {}
    parsed = {
        str(route_id): (str(values[0]), str(values[1]), str(values[2]))
        for route_id, values in routes.items()
        if isinstance(values, (list, tuple)) and len(values) == 3
    }
    return {"routes": parsed, "checked": checked}


def save_route_list(base_directory: Path, routes: dict, *, checked: str = "") -> None:
    """把线路目录存到本地（**不进备份目录**：它随时能重新拉一份）。"""
    payload = {
        "checked": checked,
        "routes": {route_id: list(values) for route_id, values in routes.items()},
    }
    _write_json(route_list_path(base_directory), payload, backup=False)


def validate_line_numbers(
    links: list[LineLink], index: dict[str, list[tuple[str, str]]]
) -> dict[str, list]:
    """拿本地线路目录校验你的线路记录：哪些线路号服务端不认识（可能记错号）。

    返回 {"unknown": [(线路号, 车型, 辆数)], "suggest": {线路号: [相近的号]}}。
    目录为空时返回空结果——**没目录就别乱报**，否则会把用户自己的记录误判成错的。
    """
    if not index:
        return {"unknown": [], "suggest": {}}
    unknown: dict[str, tuple[str, int]] = {}
    suggest: dict[str, list[str]] = {}
    for link in links:
        key = line_key(link.line)
        if not key or key in index:
            continue
        model, numbers = unknown.get(key, (link.model, 0))
        unknown[key] = (model or link.model, numbers + len(link.numbers))
        if key not in suggest:
            suggest[key] = route_list_suggest(index, key, limit=3)
    rows = [(name, model, count) for name, (model, count) in unknown.items()]
    rows.sort(key=lambda item: (len(item[0]), item[0]))
    return {"unknown": rows, "suggest": suggest}


def format_duration(seconds) -> str:
    """秒 → 「1 小时 6 分」这种人话。"""
    try:
        total = int(float(seconds))
    except (TypeError, ValueError):
        return ""
    if total < 0:
        return ""
    if total < 60:
        return "不到 1 分"
    hours, minutes = divmod(total // 60, 60)
    if hours:
        return f"{hours} 小时 {minutes} 分" if minutes else f"{hours} 小时"
    return f"{minutes} 分"


def parse_journey_time(payload) -> list[tuple[str, str, str, str]]:
    """两站间乘车耗时 → [(自编号, 出发时间, 时长, 到达时间)]（按出发时间排）。"""
    if isinstance(payload, dict):
        payload = payload.get("result") or []
    if not isinstance(payload, list):
        return []
    rows: list[tuple[str, str, str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        rows.append(
            (
                str(item.get("bus_id", "")).strip(),
                str(item.get("start", "")).strip(),
                format_duration(item.get("duration_sec")),
                str(item.get("end", "")).strip(),
            )
        )
    rows.sort(key=lambda row: row[1])
    return rows


def latest_link(links: list[LineLink]) -> LineLink | None:
    """这批记录里"最近还在跑的"那条（按 last 的年月比，同月优先"现在还在跑"）。

    主界面查一辆车时要回答"它最新在哪条线运营"——这句话的依据就是这个函数。
    """
    if not links:
        return None
    return max(links, key=lambda item: (month_key(item.last), -LINE_STATUS_ORDER.get(item.status, 9)))


def update_line_info(
    lines: dict,
    route_number: str,
    *,
    directions: list[tuple[str, str]] | None = None,
    stops: list[tuple[str, str, str, str, str]] | None = None,
    checked: str = "",
) -> dict:
    """更新一条线路的资料：首末站 / 站点表各自记自己的查证日期（没查的那项不动）。

    内部 route id **不落盘**——服务端明确说它只在当前数据集里有效、会变。
    """
    key = line_key(route_number)
    updated = dict(lines)
    record = dict(updated.get(key) or {})
    record["line"] = str(route_number).strip()
    if directions is not None:
        record["directions"] = [{"first_stop": first, "last_stop": last} for first, last in directions]
        record["directions_checked"] = checked
    if stops is not None:
        record["stops"] = [
            {"seq": seq, "name": name, "distance_m": distance, "longitude": lon, "latitude": lat}
            for seq, name, distance, lon, lat in stops
        ]
        record["stops_checked"] = checked
    updated[key] = record
    return updated


def line_info_stops(lines: dict, route_number: str) -> list[tuple[str, str, str, str, str]]:
    """取某条线路的站点表（(序号, 站名, 距上一站米, 经度, 纬度)）。"""
    record = lines.get(line_key(route_number)) or {}
    rows: list[tuple[str, str, str, str, str]] = []
    for item in record.get("stops") or []:
        if not isinstance(item, dict):
            continue
        rows.append(
            (
                str(item.get("seq", "")),
                str(item.get("name", "")),
                str(item.get("distance_m", "")),
                str(item.get("longitude", "")),
                str(item.get("latitude", "")),
            )
        )
    return rows


def line_info_directions(lines: dict, route_number: str) -> list[tuple[str, str]]:
    record = lines.get(line_key(route_number)) or {}
    return [
        (str(item.get("first_stop", "")), str(item.get("last_stop", "")))
        for item in record.get("directions") or []
        if isinstance(item, dict)
    ]


def line_info_lines_without_stops(lines: dict, links: list[LineLink]) -> list[str]:
    """体检用：线路数据里出现过、但线路资料里还没有站点表的线路。"""
    known = {line_key(name) for name in all_line_names(links)}
    return sorted(
        name
        for name in known
        if not (lines.get(name) or {}).get("stops")
    )


def apply_line_plan(current: list[LineLink], plan: LinePlan, strategy: str = "incoming") -> list[LineLink]:
    """按合并计划得到新数据（strategy="current" 时保留旧值，只补新增）。"""
    merged = {(line_key(item.line), item.model): item for item in current}
    if strategy == "current":
        for item in plan.added:
            merged[(line_key(item.line), item.model)] = item
        return list(merged.values())
    for item in plan.added:
        merged[(line_key(item.line), item.model)] = item
    for item, _old in plan.updated:
        merged[(line_key(item.line), item.model)] = item
    for item in plan.removed:
        merged.pop((line_key(item.line), item.model), None)
    return list(merged.values())


# -- 查询索引 --------------------------------------------------------------- #
def lines_for_model(links: list[LineLink], model: str) -> list[LineLink]:
    """这个车型跑过哪些线路（当前还在跑的排前面，其次时间新的在前）。"""
    found = [item for item in links if item.model and item.model == model]
    return sorted(found, key=lambda item: (LINE_STATUS_ORDER.get(item.status, 9), _negated_sort_text(item.last), line_key(item.line)))


def _negated_sort_text(text: str) -> tuple[int, str]:
    """让「时间新的排在前面」也能用一个升序 key 表达。"""
    parts = [int(piece) for piece in re.findall(r"\d+", text or "")][:2]
    return (-parts[0] if parts else 0, text)


def lines_for_number(links: list[LineLink], number: str) -> list[LineLink]:
    """这个自编号被记过跑哪些线路。"""
    wanted = number.strip()
    if not wanted:
        return []
    padded = wanted.zfill(5) if wanted.isdigit() else wanted
    found = [
        item
        for item in links
        if any(value.strip().zfill(5) == padded or value.strip() == wanted for value in item.numbers)
    ]
    return sorted(found, key=lambda item: (LINE_STATUS_ORDER.get(item.status, 9), line_key(item.line)))


def models_for_line(links: list[LineLink], line: str) -> list[str]:
    """这条线路上记录过的车型（去重，按记录时间新→旧）。"""
    key = line_key(line)
    found = [item for item in links if line_key(item.line) == key and item.model]
    found.sort(key=lambda item: (LINE_STATUS_ORDER.get(item.status, 9), _negated_sort_text(item.last)))
    models: list[str] = []
    for item in found:
        if item.model not in models:
            models.append(item.model)
    return models


def numbers_for_line(links: list[LineLink], line: str) -> list[str]:
    key = line_key(line)
    numbers: set[str] = set()
    for item in links:
        if line_key(item.line) == key:
            numbers.update(value.strip() for value in item.numbers if value.strip())
    return sorted(numbers)


def all_line_names(links: list[LineLink]) -> list[str]:
    """所有线路名（去重，保持原有写法）。"""
    names: list[str] = []
    for item in links:
        if item.line and item.line not in names:
            names.append(item.line)
    return names


def line_family_suggestions(links: list[LineLink]) -> list[tuple[str, list[str]]]:
    """疑似同号家族（886路 / 886路区间…），交给用户决定要不要合并。"""
    families: dict[str, list[str]] = {}
    for name in all_line_names(links):
        families.setdefault(line_family(name), []).append(name)
    return sorted(
        ((key, sorted(names)) for key, names in families.items() if len(names) > 1),
        key=lambda item: item[0],
    )


def build_links_from_records(records: list[RideRecord], *, source: str = "import") -> list[LineLink]:
    """把乘车记录整理成线路记录（线路 + 车型 + 涉及的自编号 + 首末时间）。

    这是通用的整理逻辑：任何来源的记录表都能走这一条路，不依赖某个人的文件。
    """
    grouped: dict[tuple[str, str], dict[str, object]] = {}
    for record in records:
        if not record.line:
            continue
        key = (line_key(record.line), record.model)
        bucket = grouped.setdefault(key, {"line": record.line, "model": record.model, "numbers": set(), "months": []})
        if record.number:
            bucket["numbers"].add(record.number.strip())  # type: ignore[union-attr]
        month = _ride_month(record)
        if month:
            bucket["months"].append(month)  # type: ignore[union-attr]
    links: list[LineLink] = []
    for bucket in grouped.values():
        months = sorted(set(bucket["months"]))  # type: ignore[arg-type]
        links.append(
            LineLink(
                line=str(bucket["line"]),
                model=str(bucket["model"]),
                numbers=tuple(sorted(bucket["numbers"])),  # type: ignore[arg-type]
                first=months[0] if months else "",
                last=months[-1] if months else "",
                status="recorded",
                source=source,
            )
        )
    return links


def _ride_month(record: RideRecord) -> str:
    """从「年份 + 日期」拼出 年.月（日期是 1.7 这种 月.日）。"""
    year = record.year.strip()
    if not year:
        return ""
    match = re.match(r"(\d{1,2})", record.date.strip())
    if match:
        month = int(match.group(1))
        if 1 <= month <= 12:
            return f"{year}.{month}"
    return year


def line_discrepancies(links: list[LineLink], ranges: list[BusRange]) -> dict[str, list[tuple[str, ...]]]:
    """线路记录和号段表对不对得上（给数据体检用）。"""
    known_models = set(model_names(ranges))
    unknown: list[tuple[str, ...]] = []
    mismatch: list[tuple[str, ...]] = []
    for link in links:
        if link.model and link.model not in known_models:
            unknown.append((link.line, link.model))
        for value in link.numbers:
            number = validate_number(value)
            if number is None:
                continue
            item = find_range(number, ranges)
            if item is not None and link.model and item.model != link.model:
                mismatch.append((link.line, f"{number:05d}", link.model, item.model))
    return {"unknown_model": sorted(set(unknown)), "number_mismatch": sorted(set(mismatch))}


UNKNOWN_TAG_TEXT = "未标注"  # 与界面层的 UNKNOWN_TAG 一致；数据层不反向依赖界面层


def applicable_suggestions(suggestions: list[TagSuggestion]) -> list[TagSuggestion]:
    """能直接写入的建议（`value` 为空的是"查到待你选"，绝不能被当成默认值写进去）。"""
    return [item for item in suggestions if item.value and item.value != UNKNOWN_TAG_TEXT]


def _write_json(path: Path, payload: object, *, backup: bool = False) -> None:
    """原子写一个 JSON 数据文件（失败时抛 ProfileError，由调用方提示）。"""
    temporary = path.with_suffix(".tmp")
    try:
        if backup:
            backup_file(path)
        temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(path)
    except OSError as error:
        temporary.unlink(missing_ok=True)
        raise ProfileError(f"无法保存 {path.name}，请确认程序文件夹可写入。") from error


def _column_index(cell_reference: str) -> int:
    letters = "".join(character for character in cell_reference if character.isalpha())
    result = 0
    for character in letters:
        result = result * 26 + ord(character.upper()) - ord("A") + 1
    return result - 1


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values: list[str] = []
    for item in root.findall(f"{{{SPREADSHEET_NS}}}si"):
        values.append("".join(node.text or "" for node in item.iter(f"{{{SPREADSHEET_NS}}}t")))
    return values


def _first_sheet_path(archive: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheet = workbook.find(f".//{{{SPREADSHEET_NS}}}sheet")
    if sheet is None:
        raise DataError("Excel 中没有工作表。")
    relationship_id = sheet.get(f"{{{OFFICE_RELATIONSHIPS_NS}}}id")
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for relationship in relationships.findall(f"{{{RELATIONSHIPS_NS}}}Relationship"):
        if relationship.get("Id") == relationship_id:
            target = relationship.get("Target", "")
            return "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
    raise DataError("无法定位 Excel 的第一个工作表。")


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{SPREADSHEET_NS}}}t"))
    value = cell.findtext(f"{{{SPREADSHEET_NS}}}v", default="")
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (IndexError, ValueError):
            raise DataError("Excel 的共享字符串数据无效。") from None
    return value


def read_sheets(path: Path) -> list[tuple[str, list[list[str]]]]:
    """读出 Excel 里的**每一张**工作表（名称 + 矩形数据）。

    号段表只需要第一张表，所以 read_table 一直没做这件事；运转记录这类表有
    6 个年度工作表，对账时要全读。
    """
    if not path.is_file():
        raise DataError(f"未找到 {path.name}。")
    try:
        with zipfile.ZipFile(path) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            targets = {
                node.get("Id"): node.get("Target", "")
                for node in relationships.findall(f"{{{RELATIONSHIPS_NS}}}Relationship")
            }
            shared_strings = _read_shared_strings(archive)
            sheets: list[tuple[str, list[list[str]]]] = []
            for sheet in workbook.iter(f"{{{SPREADSHEET_NS}}}sheet"):
                name = sheet.get("name", "")
                relationship_id = sheet.get(f"{{{OFFICE_RELATIONSHIPS_NS}}}id")
                target = targets.get(relationship_id, "")
                if not target:
                    continue
                member = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
                try:
                    raw = archive.read(member)
                except KeyError:
                    continue
                sheets.append((name, _rows_from_xml(raw, shared_strings)))
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError) as error:
        raise DataError(f"无法读取 {path.name}。请确认它是有效的 .xlsx 文件。") from error
    return sheets


def _rows_from_xml(raw: bytes, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(raw)
    table: list[list[str]] = []
    for row in root.iter(f"{{{SPREADSHEET_NS}}}row"):
        cells: dict[int, str] = {}
        for cell in row.iter(f"{{{SPREADSHEET_NS}}}c"):
            cells[_column_index(cell.get("r", ""))] = _cell_text(cell, shared_strings)
        if cells:
            table.append([cells.get(index, "") for index in range(max(cells) + 1)])
    return table


def _worksheet_table(archive: zipfile.ZipFile) -> list[list[str]]:
    """Every row of the first worksheet as a rectangular list of strings."""
    shared_strings = _read_shared_strings(archive)
    root = ET.fromstring(archive.read(_first_sheet_path(archive)))
    rows: list[dict[int, str]] = []
    width = 0
    for row in root.findall(f".//{{{SPREADSHEET_NS}}}sheetData/{{{SPREADSHEET_NS}}}row"):
        cells: dict[int, str] = {}
        for cell in row.findall(f"{{{SPREADSHEET_NS}}}c"):
            index = _column_index(cell.get("r", ""))
            if index >= 0:
                cells[index] = _cell_text(cell, shared_strings).strip()
        if cells:
            width = max(width, max(cells) + 1)
        rows.append(cells)
    return [[row.get(index, "") for index in range(width)] for row in rows]


def _worksheet_rows(archive: zipfile.ZipFile) -> list[list[str]]:
    return [(row + ["", "", ""])[:3] for row in _worksheet_table(archive)]


def _column_letter(index: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def read_table(path: Path) -> tuple[list[str], list[list[str]]]:
    """Return (header, data rows) of the first worksheet, with every column."""
    if not path.is_file():
        return [], []
    try:
        with zipfile.ZipFile(path) as archive:
            table = _worksheet_table(archive)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError):
        return [], []
    if not table:
        return [], []
    return table[0], table[1:]


def write_table(path: Path, header: list[str], rows: list[list[str]], numeric_columns: set[int] | frozenset[int] = frozenset()) -> None:
    """Write a minimal single-sheet xlsx atomically (numbers stay numbers)."""
    worksheet = ET.Element(f"{{{SPREADSHEET_NS}}}worksheet")
    sheet_data = ET.SubElement(worksheet, f"{{{SPREADSHEET_NS}}}sheetData")
    for row_number, values in enumerate([header] + [list(row) for row in rows], start=1):
        row = ET.SubElement(sheet_data, f"{{{SPREADSHEET_NS}}}row", {"r": str(row_number)})
        for column, value in enumerate(values, start=1):
            reference = f"{_column_letter(column)}{row_number}"
            text = "" if value is None else str(value)
            if row_number > 1 and (column - 1) in numeric_columns and re.fullmatch(r"-?\d+", text):
                cell = ET.SubElement(row, f"{{{SPREADSHEET_NS}}}c", {"r": reference})
                ET.SubElement(cell, f"{{{SPREADSHEET_NS}}}v").text = text
            else:
                cell = ET.SubElement(row, f"{{{SPREADSHEET_NS}}}c", {"r": reference, "t": "inlineStr"})
                inline = ET.SubElement(cell, f"{{{SPREADSHEET_NS}}}is")
                ET.SubElement(inline, f"{{{SPREADSHEET_NS}}}t").text = text
    workbook = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{SPREADSHEET_NS}" xmlns:r="{OFFICE_RELATIONSHIPS_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    relationships = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'''
    root_relationships = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'''
    temporary = path.with_suffix(".tmp")
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", content_types)
            archive.writestr("_rels/.rels", root_relationships)
            archive.writestr("xl/workbook.xml", workbook)
            archive.writestr("xl/_rels/workbook.xml.rels", relationships)
            archive.writestr("xl/worksheets/sheet1.xml", ET.tostring(worksheet, encoding="utf-8", xml_declaration=True))
        temporary.replace(path)
    except OSError:
        temporary.unlink(missing_ok=True)
        raise


def extra_columns_by_row(
    path: Path,
) -> tuple[list[str], dict[tuple[str, str, str], list[str]], dict[str, list[str]], dict[tuple[str, str], tuple[str, list[str]]]]:
    """Collect the columns after the first three, indexed by row, by model and by span."""
    header, rows = read_table(path)
    if len(header) <= 3:
        return [], {}, {}, {}
    extras_header = header[3:]
    by_key: dict[tuple[str, str, str], list[str]] = {}
    by_model: dict[str, list[str]] = {}
    by_span: dict[tuple[str, str], tuple[str, list[str]]] = {}
    for row in rows:
        padded = (row + [""] * len(header))[: len(header)]
        extras = padded[3:]
        by_key.setdefault((padded[0], padded[1], padded[2]), extras)
        by_span.setdefault((padded[0], padded[1]), (padded[2], extras))
        if padded[2]:
            by_model.setdefault(padded[2], extras)
    return extras_header, by_key, by_model, by_span


def load_ranges(path: Path) -> list[BusRange]:
    if not path.is_file():
        raise DataError(f"未找到 {DATA_FILENAME}。请将该文件放在程序同一文件夹中。")
    try:
        with zipfile.ZipFile(path) as archive:
            rows = _worksheet_rows(archive)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError) as error:
        raise DataError("无法读取 Excel 文件。请确认文件未损坏且为有效 .xlsx 格式。") from error
    if not rows or tuple(rows[0]) != EXPECTED_HEADERS:
        raise DataError("Excel 前三列表头必须依次为：号段起始、号段结束、车型。")

    ranges: list[BusRange] = []
    for line_number, row in enumerate(rows[1:], start=2):
        if row == ["", "", ""]:
            continue
        try:
            start, end = int(row[0]), int(row[1])
        except ValueError:
            raise DataError(f"Excel 第 {line_number} 行的号段起始或号段结束不是有效整数。") from None
        model = row[2].strip()
        if start > end or not model:
            raise DataError(f"Excel 第 {line_number} 行的号段或车型无效。")
        ranges.append(BusRange(start, end, model))
    if not ranges:
        raise DataError("Excel 中没有可用的号段数据。")
    ranges.sort(key=lambda item: item.start)
    for previous, current in zip(ranges, ranges[1:]):
        if current.start <= previous.end:
            raise DataError(f"Excel 号段重叠：{previous.start}-{previous.end} 与 {current.start}-{current.end}。")
    return ranges


def load_ranges_lenient(path: Path) -> tuple[list[BusRange], list[str]]:
    """尽量把号段读出来：跳过错行、把重叠行里靠后的那条丢掉。

    数据文件坏了时用来救急——返回 (可用的号段, 问题说明)；写回文件（例如在号段管理里
    保存一次）就会顺手把坏行修掉。正常的 load_ranges 仍然严格校验，不会放水。
    """
    if not path.is_file():
        raise DataError(f"未找到 {DATA_FILENAME}。请将该文件放在程序同一文件夹中。")
    try:
        with zipfile.ZipFile(path) as archive:
            rows = _worksheet_rows(archive)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError) as error:
        raise DataError("无法读取 Excel 文件。请确认文件未损坏且为有效 .xlsx 格式。") from error
    if not rows:
        raise DataError("Excel 是空的。")
    header = tuple(rows[0])
    notes: list[str] = []
    if header[:3] != EXPECTED_HEADERS:
        notes.append(f"表头是「{'、'.join(header[:3])}」，已按前三列当作号段处理。")

    accepted: list[BusRange] = []
    for line_number, row in enumerate(rows[1:], start=2):
        padded = (list(row) + ["", "", ""])[:3]
        if not any(cell.strip() for cell in padded):
            continue
        try:
            start, end = int(padded[0]), int(padded[1])
        except ValueError:
            notes.append(f"第 {line_number} 行号段不是数字，已跳过")
            continue
        model = padded[2].strip()
        if start > end or not model:
            notes.append(f"第 {line_number} 行（{start}-{end}）号段或车型无效，已跳过")
            continue
        accepted.append(BusRange(start, end, model))

    accepted.sort(key=lambda item: item.start)
    kept: list[BusRange] = []
    for item in accepted:
        clash = next((old for old in kept if item.start <= old.end and item.end >= old.start), None)
        if clash is not None:
            notes.append(f"{item.start}-{item.end} 与 {clash.start}-{clash.end} 重叠，已丢弃前者")
            continue
        kept.append(item)
    if not kept:
        raise DataError("Excel 中没有可用的号段数据。")
    return kept, notes


def validate_range(start_text: str, end_text: str, model: str, ranges: list[BusRange], exclude_index: int | None = None) -> BusRange:
    """Validate a user-maintained range and reject all intersections with others."""
    try:
        start, end = int(start_text.strip()), int(end_text.strip())
    except ValueError:
        raise RangeError("号段起始和号段结束必须是五位数字。") from None
    model = model.strip()
    if not (10000 <= start <= 99999 and 10000 <= end <= 99999):
        raise RangeError("号段起始和号段结束必须是五位数字。")
    if start > end:
        raise RangeError("号段起始不能大于号段结束。")
    if not model:
        raise RangeError("请输入对应车型。")
    for index, current in enumerate(ranges):
        if index != exclude_index and start <= current.end and end >= current.start:
            raise RangeError(f"号段与现有范围 {current.start}-{current.end} 重复，无法保存。")
    return BusRange(start, end, model)


def save_ranges(path: Path, ranges: list[BusRange]) -> None:
    """写回号段表。

    原文件里第 4 列之后的额外列会被保留：整行（起止+车型）没变的沿用原值；
    **同一个号段只是改了车型名**（例如改掉写错的型号）的沿用该号段的额外列，
    并把额外列里出现的旧型号字样一起换掉；车型仍在但号段改过的沿用该车型第一次
    出现的额外列；其余留空。
    """
    ordered = sorted(ranges, key=lambda item: item.start)
    extras_header, by_key, by_model, by_span = extra_columns_by_row(path)
    rows: list[list[str]] = []
    for item in ordered:
        extras = by_key.get((str(item.start), str(item.end), item.model))
        if extras is None:
            old_model, span_extras = by_span.get((str(item.start), str(item.end)), ("", None))  # type: ignore[assignment]
            if span_extras is not None:
                extras = [
                    cell.replace(old_model, item.model) if old_model and old_model in cell else cell
                    for cell in span_extras
                ]
        if extras is None:
            extras = by_model.get(item.model)
        if extras is None:
            extras = [""] * len(extras_header)
        rows.append([str(item.start), str(item.end), item.model] + list(extras))
    try:
        backup_file(path)
        write_table(path, list(EXPECTED_HEADERS) + list(extras_header), rows, numeric_columns={0, 1})
    except OSError as error:
        raise RangeError(f"无法保存 {DATA_FILENAME}，请确认文件未被 Excel 占用且程序文件夹可写入。") from error


@dataclass(frozen=True)
class MergePlan:
    """新旧号段表合并后的差异。"""

    added: list[BusRange]
    updated: list[tuple[BusRange, str]]  # (新号段, 原来的车型)
    removed: list[BusRange]
    conflict: list[tuple[BusRange, str]]  # (新号段, 与它重叠的旧号段车型)

    @property
    def total(self) -> int:
        return len(self.added) + len(self.updated) + len(self.removed) + len(self.conflict)

    @property
    def summary(self) -> str:
        return f"新增 {len(self.added)} · 修改 {len(self.updated)} · 删除 {len(self.removed)} · 冲突 {len(self.conflict)}"


def merge_ranges(current: list[BusRange], incoming: list[BusRange]) -> MergePlan:
    """比较当前号段与导入的号段。

    同一个自编号区间（起始+结束完全一样）而车型不同算「修改」；
    与旧号段部分重叠但区间不同的算「冲突」（需要人决定怎么处理）。
    """
    deduped: dict[tuple[int, int], BusRange] = {}
    for item in incoming:
        deduped[(item.start, item.end)] = item
    current_by_key = {(item.start, item.end): item for item in current}
    added: list[BusRange] = []
    updated: list[tuple[BusRange, str]] = []
    conflict: list[tuple[BusRange, str]] = []
    for _, item in sorted(deduped.items()):
        existing = current_by_key.get((item.start, item.end))
        if existing is not None:
            if existing.model != item.model:
                updated.append((item, existing.model))
            continue
        overlapping = next((old for old in current if item.start <= old.end and item.end >= old.start), None)
        if overlapping is not None:
            conflict.append((item, overlapping.model))
        else:
            added.append(item)
    removed = [item for item in current if (item.start, item.end) not in deduped]
    return MergePlan(added=added, updated=updated, removed=removed, conflict=conflict)


def merge_conflict_keys(plan: MergePlan, current: list[BusRange]) -> set[tuple[int, int]]:
    """冲突的新区间会顶掉的旧号段（按区间键）。"""
    replaced: set[tuple[int, int]] = set()
    for item, _old in plan.conflict:
        for old in current:
            if item.start <= old.end and item.end >= old.start:
                replaced.add((old.start, old.end))
    return replaced


def apply_merge(current: list[BusRange], plan: MergePlan, strategy: str = "incoming") -> list[BusRange]:
    """按策略算出合并结果。

    incoming：以新表为准（新增 + 修改 + 删除新表里没有的，冲突区间由新表顶掉）；
    additive：只增不改（保留全部旧号段，只补新表里不与任何旧号段重叠的新增）。

    两种策略都不会产生重叠——否则数据文件会加载失败。
    """
    updates = {(item.start, item.end): item for item, _old in plan.updated}
    if strategy == "additive":
        result = list(current)
        occupied = [(item.start, item.end) for item in result]
        for item in plan.added:
            if any(item.start <= end and item.end >= start for start, end in occupied):
                continue
            result.append(item)
            occupied.append((item.start, item.end))
        return sorted(result, key=lambda item: item.start)

    removed = {(item.start, item.end) for item in plan.removed} | merge_conflict_keys(plan, current)
    result: list[BusRange] = []
    seen: set[tuple[int, int]] = set()
    for old in current:
        key = (old.start, old.end)
        if key in removed:
            continue
        result.append(updates.get(key, old))
        seen.add(key)
    for item in plan.added + [entry for entry, _old in plan.conflict]:
        key = (item.start, item.end)
        if key not in seen:
            result.append(item)
            seen.add(key)
    return sorted(result, key=lambda item: item.start)


def find_range(number: int, ranges: list[BusRange]) -> BusRange | None:
    low, high = 0, len(ranges) - 1
    while low <= high:
        middle = (low + high) // 2
        current = ranges[middle]
        if number < current.start:
            high = middle - 1
        elif number > current.end:
            low = middle + 1
        else:
            return current
    return None


def find_model(number: int, ranges: list[BusRange]) -> str | None:
    item = find_range(number, ranges)
    return item.model if item else None


def validate_number(value: str) -> int | None:
    """Return a five-digit bus number, or None when the user input is invalid."""
    value = value.strip()
    if len(value) != 5 or not value.isdecimal():
        return None
    try:
        return int(value)
    except ValueError:
        return None


def model_names(ranges: list[BusRange]) -> list[str]:
    return sorted({item.model for item in ranges}, key=str.casefold)


def ranges_for_model(model: str, ranges: list[BusRange]) -> list[BusRange]:
    return [item for item in ranges if item.model == model]


def group_missing_numbers(numbers: list[int]) -> list[tuple[int, int]]:
    """把未收录的自编号合并成连续区间。"""
    ordered = sorted({int(number) for number in numbers})
    if not ordered:
        return []
    groups: list[tuple[int, int]] = []
    start = previous = ordered[0]
    for number in ordered[1:]:
        if number == previous + 1:
            previous = number
            continue
        groups.append((start, previous))
        start = previous = number
    groups.append((start, previous))
    return groups


def parse_number_list(text: str) -> tuple[list[int], list[str]]:
    """Split pasted text into unique five-digit numbers plus the parts ignored."""
    numbers: list[int] = []
    ignored: list[str] = []
    seen: set[int] = set()
    for token in re.split(r"[\s,，、;；|/\\]+", text.strip()):
        if not token:
            continue
        number = validate_number(token)
        if number is None:
            ignored.append(token)
        elif number not in seen:
            seen.add(number)
            numbers.append(number)
    return numbers, ignored


def segment_lookup(start: int, end: int, ranges: list[BusRange]) -> list[tuple[int, int, str | None]]:
    """Split an inclusive number range into model segments, gaps included."""
    segments: list[tuple[int, int, str | None]] = []
    cursor = start
    for item in ranges:
        if item.end < cursor:
            continue
        if item.start > end:
            break
        if item.start > cursor:
            segments.append((cursor, min(item.start - 1, end), None))
            cursor = item.start
        if cursor > end:
            break
        segments.append((cursor, min(item.end, end), item.model))
        cursor = min(item.end, end) + 1
        if cursor > end:
            break
    if cursor <= end:
        segments.append((cursor, end, None))
    return segments
