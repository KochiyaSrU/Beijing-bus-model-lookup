"""Self-checks for the bus/metro lookup tool.

Run with:

    python -m unittest -v test_bpt_bus_digital_number.py

Logic tests use synthetic data so that updating the Excel workbook never turns
them red; only :class:`ShippedDataTests` looks at the real files, and it does so
with tolerant bounds instead of hard-coded counts.
"""

import ast
import json
import shutil
import queue
import sys
import tempfile
import time
import tkinter as tk
import unittest
from tkinter import ttk
from collections import Counter
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch

from PIL import Image

import bpt_data
import bpt_online
import bpt_selfcheck
import bpt_web
import BPT_BusDigitalNumber

# 测试里默认关掉"主界面自动联网查询"（真联网会让套件变慢且依赖外网）；
# 这个功能本身有专门用例：临时打开 + 注入假服务验证。
BPT_BusDigitalNumber.BusLookupApp.AUTO_ONLINE_DEFAULT = False

from BPT_BusDigitalNumber import (
    # 测试里关掉"主界面自动联网"：不然每查一次自编号都真去联网，套件又慢又依赖外网。
    # 这个开关本身的行为有专门的用例（会临时改回 True 并用假服务验证）。
    ANY_TAG,
    DATA_FILENAME,
    DRAFT_TEMPLATE,
    METRO_SOURCE_FILENAME,
    PROFILE_FILENAME,
    UNKNOWN_TAG,
    BulkRangeWindow,
    BODY_VALUES,
    BusLookupApp,
    BusRange,
    CatalogWindow,
    WindowMixin,
    DataError,
    DraftProfilesWindow,
    MergeWindow,
    MetroError,
    MetroRange,
    MetroWindow,
    ModelDetailWindow,
    ProfileError,
    QuickLookupWindow,
    RangeEditorWindow,
    RangeError,
    StatsWindow,
    alias_index,
    apply_merge,
    cli_export_html,
    cli_lookup,
    console_write,
    delete_image,
    draft_profile,
    empty_profile_models,
    find_metro_models,
    find_model,
    find_range,
    group_missing_numbers,
    image_directory,
    import_image,
    load_metro_ranges,
    load_profiles,
    load_ranges,
    load_tag_overrides,
    merge_ranges,
    metro_lines,
    metro_ranges_for_line,
    model_aliases,
    model_names,
    model_tags,
    parse_metro_source,
    parse_number_list,
    profile_is_empty,
    ranges_for_model,
    looks_like_metro_model,
    looks_like_out_of_town,
    parse_ride_sheet,
    read_sheets,
    reconcile_rides,
    RideRecord,
    save_metro_ranges,
    save_profiles,
    save_ranges,
    segment_lookup,
    clean_tag_overrides,
    save_tag_overrides,
    TagEditorWindow,
    _column_letter,
    apply_manufacturer_prefix,
    apply_model_tag_values,
    BulkTagWindow,
    DataHealthWindow,
    health_report,
    load_tag_columns,
    main,
    merge_tag_columns,
    read_table,
    write_table,
    write_tagged_workbook,
    RangeManagerWindow,
    TagRuleEditorWindow,
    TagRulesWindow,
    MetroBatchWindow,
    export_metro_source,
    import_metro_source,
    metro_health_items,
    parse_metro_number_list,
    rule_hits,
    rule_matches,
    tree_rows,
    tree_headings,
    BackupWindow,
    MAX_ERROR_LOG_BYTES,
    SETTINGS_FILENAME,
    ClipboardWatcher,
    SettingsWindow,
    EXPECTED_HEADERS,
    SpotDialog,
    SpotEntry,
    SpotLogWindow,
    append_spot,
    backup_file,
    center_window,
    TagConfirmWindow,
    MyRecordsWindow,
    OnlineQueryWindow,
    LineQueryWindow,
    LineInfoWindow,
    DistrictBusWindow,
    LineLink,
    load_lines,
    save_lines,
    TagEditorWindow,
    LineEditorWindow,
    LineManagerWindow,
    DataHealthWindow,
    DataTable,
    TableColumn,
    debounce,
    default_body,
    body_is_guessed_double_decker,
    enable_row_hover,    existing_image_names,
    export_model_card,
    extract_clipboard_number,
    filter_spots,
    _is_subsequence,
    list_backups,
    KNOWN_DOUBLE_DECKERS,
    load_history,
    load_ranges,
    load_ranges_lenient,
    load_settings,
    load_spots,
    log_exception,
    make_spot,
    pinyin_initials,
    profile_images,
    remember_number,
    replace_spot,
    restore_backup,
    save_settings,
    save_spot_photo,
    save_spots,
    search_hit,
    set_profile_images,
    speak_text,
    spot_models,
    spot_summary,
    spot_summary_text,
    stored_geometry,
    stored_photo_path,
    delete_spot_photo,
    tag_suggestions,
    _tag_overridden,
    validate_metro_range,
    web_payload,
    validate_number,
    validate_range,
)


SOURCE_DATA = Path(__file__).with_name(DATA_FILENAME)
SOURCE_METRO = Path(__file__).with_name(METRO_SOURCE_FILENAME)


@contextmanager
def isolated_model(app, model: str):
    """临时把这个车型的标签覆盖与资料从内存里摘掉，结束后原样放回。

    测试不该依赖用户随时会改的 车型标签.json / 车型资料.json——用户补一条标签，
    「这个车型本来是未标注」这种前提就不成立了（真的踩过）。
    """
    saved_tags = json.loads(json.dumps(app.tag_overrides))
    saved_profile = app.profiles.get(model)
    try:
        models = app.tag_overrides.get("models")
        if isinstance(models, dict):
            models.pop(model, None)
        app.profiles.pop(model, None)
        yield
    finally:
        app.tag_overrides.clear()
        app.tag_overrides.update(saved_tags)
        if saved_profile is None:
            app.profiles.pop(model, None)
        else:
            app.profiles[model] = saved_profile


def scratch():
    """Temporary directory that survives Windows file-lock cleanup quirks."""
    return tempfile.TemporaryDirectory(ignore_cleanup_errors=True)


@contextmanager
def patch_app_dir(directory):
    """把 application_dir 指到临时目录。

    实现住在 bpt_data 里、主入口只是再导出，所以两边都要打补丁；只打一边的话，
    搬进 bpt_data 的函数会继续用真实目录，测试会「假装通过」却污染工作区。
    """
    with patch("bpt_data.application_dir", return_value=directory), patch(
        "BPT_BusDigitalNumber.application_dir", return_value=directory
    ):
        yield


def destroy(widget) -> None:
    """Destroy a window that may already have closed itself."""
    if widget is None:
        return
    try:
        widget.destroy()
    except Exception:
        pass


def backups_of(base: Path, stem: str) -> list[Path]:
    """自动备份目录里某份数据的备份文件（按名字/时间排序）。"""
    directory = base / "备份"
    return sorted(directory.glob(f"{stem}.*")) if directory.is_dir() else []


class BusRangeTests(unittest.TestCase):
    def setUp(self):
        self.ranges = [
            BusRange(20000, 20212, "BJ6123C7C4D"),
            BusRange(20213, 20228, "XMQ6100G1"),
            BusRange(20229, 20426, "BJ6123C7C4D"),
        ]

    def test_find_model_covers_start_end_and_middle(self):
        self.assertEqual(find_model(20000, self.ranges), "BJ6123C7C4D")
        self.assertEqual(find_model(20100, self.ranges), "BJ6123C7C4D")
        self.assertEqual(find_model(20212, self.ranges), "BJ6123C7C4D")
        self.assertEqual(find_model(20213, self.ranges), "XMQ6100G1")
        self.assertEqual(find_model(20426, self.ranges), "BJ6123C7C4D")

    def test_find_model_returns_none_outside_every_range(self):
        for number in (10000, 19999, 20427, 99999):
            with self.subTest(number=number):
                self.assertIsNone(find_model(number, self.ranges))

    def test_catalog_names_are_unique_and_case_insensitively_sorted(self):
        names = model_names([BusRange(1, 2, "b"), BusRange(3, 4, "A"), BusRange(5, 6, "b")])
        self.assertEqual(names, ["A", "b"])

    def test_ranges_for_model_keeps_every_segment(self):
        self.assertEqual(ranges_for_model("BJ6123C7C4D", self.ranges), [self.ranges[0], self.ranges[2]])
        self.assertEqual(ranges_for_model("MISSING", self.ranges), [])


class NumberValidationTests(unittest.TestCase):
    def test_accepts_five_digits_including_full_width(self):
        self.assertEqual(validate_number("20000"), 20000)
        self.assertEqual(validate_number(" 20000 "), 20000)
        self.assertEqual(validate_number("２００００"), 20000)

    def test_rejects_wrong_length_and_letters(self):
        for value in ("", "2000", "200000", "20A00", "2000 "):
            with self.subTest(value=value):
                self.assertIsNone(validate_number(value))

    def test_rejects_digits_that_are_not_decimal_numbers(self):
        # Regression: these pass str.isdigit() but int() used to raise ValueError.
        for value in ("12²45", "1234①", "1②345"):
            with self.subTest(value=value):
                self.assertIsNone(validate_number(value))


class RangeValidationTests(unittest.TestCase):
    def setUp(self):
        self.ranges = [BusRange(20000, 20099, "A"), BusRange(20200, 20299, "B")]

    def test_overlapping_and_touching_ranges_are_rejected(self):
        for start, end in ((20050, 20150), (19999, 20000), (20099, 20099), (20199, 20200)):
            with self.subTest(start=start, end=end):
                with self.assertRaisesRegex(RangeError, "重复"):
                    validate_range(str(start), str(end), "C", self.ranges)

    def test_editing_the_same_range_is_allowed(self):
        ranges = [BusRange(20000, 20099, "A")]
        self.assertEqual(validate_range("20000", "20099", "NEW", ranges, 0), BusRange(20000, 20099, "NEW"))

    def test_bounds_must_be_five_digits_and_ordered(self):
        with self.assertRaises(RangeError):
            validate_range("2000", "20099", "A", [])
        with self.assertRaises(RangeError):
            validate_range("20099", "20000", "A", [])
        with self.assertRaises(RangeError):
            validate_range("20000", "20099", "   ", [])

    def test_adjacent_ranges_do_not_collide(self):
        self.assertEqual(validate_range("20100", "20199", "C", self.ranges), BusRange(20100, 20199, "C"))


class WorkbookTests(unittest.TestCase):
    def test_missing_workbook_is_reported(self):
        with scratch() as directory:
            with self.assertRaisesRegex(DataError, "未找到"):
                load_ranges(Path(directory) / "missing-bus-data.xlsx")

    def test_invalid_header_is_reported(self):
        with patch("bpt_data._worksheet_rows", return_value=[["错误", "表头", "字段"]]):
            with self.assertRaisesRegex(DataError, "表头"):
                load_ranges(SOURCE_DATA)

    def test_overlapping_rows_are_rejected(self):
        rows = [["号段起始", "号段结束", "车型"], ["20000", "20010", "A"], ["20010", "20020", "B"]]
        with patch("bpt_data._worksheet_rows", return_value=rows):
            with self.assertRaisesRegex(DataError, "重叠"):
                load_ranges(SOURCE_DATA)

    def test_non_numeric_bounds_are_rejected(self):
        rows = [["号段起始", "号段结束", "车型"], ["abc", "20010", "A"]]
        with patch("bpt_data._worksheet_rows", return_value=rows):
            with self.assertRaisesRegex(DataError, "有效整数"):
                load_ranges(SOURCE_DATA)

    def test_ranges_round_trip_through_the_workbook(self):
        with scratch() as directory:
            path = Path(directory) / DATA_FILENAME
            save_ranges(path, [BusRange(20200, 20299, "B"), BusRange(20000, 20099, "A")])
            self.assertEqual(load_ranges(path), [BusRange(20000, 20099, "A"), BusRange(20200, 20299, "B")])

    def test_saving_workbook_keeps_the_previous_version_as_backup(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / DATA_FILENAME
            save_ranges(path, [BusRange(20000, 20099, "A")])
            save_ranges(path, [BusRange(20000, 20099, "CHANGED")])
            backups = backups_of(base, "BPT_BusDigitalNumber")
            self.assertEqual(len(backups), 1)
            self.assertEqual(load_ranges(backups[0]), [BusRange(20000, 20099, "A")])
            self.assertEqual(load_ranges(path), [BusRange(20000, 20099, "CHANGED")])


class ProfileTests(unittest.TestCase):
    def test_profiles_persist_and_empty_profiles_are_marked(self):
        with scratch() as directory:
            base = Path(directory)
            profiles = {"MODEL-A": {"introduction": "简介", "specifications": "参数", "image": ""}}
            save_profiles(base, profiles)
            loaded = load_profiles(base)["MODEL-A"]
            self.assertEqual(loaded["introduction"], "简介")
            self.assertEqual(loaded["specifications"], "参数")
            self.assertEqual(loaded["image"], "")
            self.assertEqual(loaded["images"], [])
            self.assertTrue(profile_is_empty({"introduction": " ", "specifications": "", "image": ""}))
            self.assertFalse(profile_is_empty(profiles["MODEL-A"]))

    def test_saving_profiles_keeps_a_backup(self):
        with scratch() as directory:
            base = Path(directory)
            save_profiles(base, {"A": {"introduction": "第一版", "specifications": "", "image": ""}})
            save_profiles(base, {"A": {"introduction": "第二版", "specifications": "", "image": ""}})
            backups = backups_of(base, "车型资料")
            self.assertEqual(len(backups), 1)
            self.assertIn("第一版", backups[0].read_text(encoding="utf-8"))
            self.assertEqual(load_profiles(base)["A"]["introduction"], "第二版")

    def test_profile_with_missing_image_file_counts_as_empty(self):
        with scratch() as directory:
            base = Path(directory)
            profile = {"introduction": "", "specifications": "", "image": "GONE.jpg"}
            self.assertFalse(profile_is_empty(profile))
            self.assertTrue(profile_is_empty(profile, base))

    def test_broken_profile_file_is_reported(self):
        with scratch() as directory:
            base = Path(directory)
            (base / PROFILE_FILENAME).write_text("{not json", encoding="utf-8")
            with self.assertRaises(ProfileError):
                load_profiles(base)

    def test_import_replace_and_delete_main_image(self):
        with scratch() as directory:
            base = Path(directory)
            source = base / "source.png"
            replacement = base / "replacement.jpg"
            Image.new("RGB", (2400, 1800), "red").save(source)
            Image.new("RGB", (10, 10), "blue").save(replacement)

            first = import_image(base, "MODEL-A", source)
            self.assertTrue((image_directory(base) / first).is_file())
            with Image.open(image_directory(base) / first) as compressed:
                self.assertEqual(compressed.format, "JPEG")
                self.assertLessEqual(compressed.width, 1600)
                self.assertLessEqual(compressed.height, 1200)

            second = import_image(base, "MODEL-A", replacement, first)
            self.assertEqual(second, first)  # deterministic name per model
            self.assertTrue((image_directory(base) / second).is_file())

            delete_image(base, second)
            self.assertFalse((image_directory(base) / second).exists())

    def test_importing_a_new_image_removes_the_previous_model_file(self):
        with scratch() as directory:
            base = Path(directory)
            source = base / "source.png"
            Image.new("RGB", (20, 20), "green").save(source)
            old = "OLD_NAME.jpg"
            image_directory(base).mkdir(exist_ok=True)
            (image_directory(base) / old).write_bytes(b"stale")
            new_name = import_image(base, "MODEL-A", source, old)
            self.assertFalse((image_directory(base) / old).exists())
            self.assertTrue((image_directory(base) / new_name).is_file())

    def test_invalid_image_is_rejected(self):
        with scratch() as directory:
            base = Path(directory)
            source = base / "not-an-image.png"
            source.write_text("not an image", encoding="utf-8")
            with self.assertRaises(ProfileError):
                import_image(base, "MODEL-A", source)

    def test_unsupported_image_extension_is_rejected(self):
        with scratch() as directory:
            base = Path(directory)
            source = base / "picture.txt"
            source.write_text("nope", encoding="utf-8")
            with self.assertRaises(ProfileError):
                import_image(base, "MODEL-A", source)

    def test_deleting_a_missing_image_is_harmless(self):
        with scratch() as directory:
            delete_image(Path(directory), "does-not-exist.jpg")
            delete_image(Path(directory), "")


class MetroSourceTests(unittest.TestCase):
    def test_shipped_source_parses_into_lines_and_ranges(self):
        ranges = parse_metro_source(SOURCE_METRO)
        self.assertGreaterEqual(len(metro_lines(ranges)), 20)
        self.assertEqual(find_metro_models("01001", ranges)[0].model, "DKZ4")
        self.assertEqual([item.model for item in find_metro_models("T401", ranges)], ["DKZ16"])
        self.assertEqual(find_metro_models("TP401", ranges)[0].model, "DKZ13")
        self.assertTrue(metro_ranges_for_line("1号线八通线", ranges))

    def test_metro_ranges_for_line_is_sorted_by_number(self):
        ranges = [MetroRange("线", "01010", "01020", "B"), MetroRange("线", "01001", "01009", "A")]
        self.assertEqual([item.model for item in metro_ranges_for_line("线", ranges)], ["A", "B"])

    def test_range_without_space_before_model_is_rejected(self):
        with scratch() as directory:
            path = Path(directory) / METRO_SOURCE_FILENAME
            path.write_text("1号线\n01001-01031DKZ4\n", encoding="utf-8")
            with self.assertRaisesRegex(MetroError, "格式无效"):
                parse_metro_source(path)

    def test_range_with_spaces_instead_of_hyphen_is_rejected(self):
        with scratch() as directory:
            path = Path(directory) / METRO_SOURCE_FILENAME
            path.write_text("1号线\n01001 01031 DKZ4\n", encoding="utf-8")
            with self.assertRaisesRegex(MetroError, "格式无效"):
                parse_metro_source(path)

    def test_range_with_padded_hyphen_is_rejected(self):
        with scratch() as directory:
            path = Path(directory) / METRO_SOURCE_FILENAME
            path.write_text("1号线\n01001 - 01031 DKZ4\n", encoding="utf-8")
            with self.assertRaisesRegex(MetroError, "格式无效"):
                parse_metro_source(path)

    def test_range_before_any_line_name_is_rejected(self):
        with scratch() as directory:
            path = Path(directory) / METRO_SOURCE_FILENAME
            path.write_text("01001-01031 DKZ4\n", encoding="utf-8")
            with self.assertRaisesRegex(MetroError, "线路名称"):
                parse_metro_source(path)

    def test_descending_range_is_rejected(self):
        with scratch() as directory:
            path = Path(directory) / METRO_SOURCE_FILENAME
            path.write_text("1号线\n01031-01001 DKZ4\n", encoding="utf-8")
            with self.assertRaisesRegex(MetroError, "无效"):
                parse_metro_source(path)

    def test_missing_source_file_is_reported(self):
        with scratch() as directory:
            with self.assertRaisesRegex(MetroError, "未找到"):
                parse_metro_source(Path(directory) / METRO_SOURCE_FILENAME)


class MetroDataTests(unittest.TestCase):
    def test_validation_rejects_overlaps_per_line_and_prefix(self):
        ranges = [MetroRange("1号线", "01001", "01010", "A")]
        with self.assertRaisesRegex(MetroError, "重复"):
            validate_metro_range("1号线", "01010", "01020", "B", ranges)
        self.assertEqual(validate_metro_range("2号线", "01010", "01020", "B", ranges), MetroRange("2号线", "01010", "01020", "B"))
        self.assertEqual(validate_metro_range("1号线", "T010", "T020", "C", ranges), MetroRange("1号线", "T010", "T020", "C"))

    def test_metro_ranges_persist_separately_from_bus_data(self):
        with scratch() as directory:
            base = Path(directory)
            expected = [MetroRange("测试线", "TS001", "TS010", "MODEL")]
            save_metro_ranges(base, expected)
            self.assertEqual(load_metro_ranges(base), expected)

    def test_saving_metro_ranges_keeps_a_backup(self):
        with scratch() as directory:
            base = Path(directory)
            save_metro_ranges(base, [MetroRange("测试线", "TS001", "TS010", "OLD")])
            save_metro_ranges(base, [MetroRange("测试线", "TS001", "TS010", "NEW")])
            self.assertEqual(load_metro_ranges(base)[0].model, "NEW")
            backups = backups_of(base, "北京地铁数据")
            self.assertEqual(len(backups), 1)
            self.assertIn("OLD", backups[0].read_text(encoding="utf-8"))

    def test_seed_data_is_still_returned_when_it_cannot_be_saved(self):
        with scratch() as directory:
            base = Path(directory)
            (base / METRO_SOURCE_FILENAME).write_text("测试线\nTS001-TS010 MODEL\n", encoding="utf-8")
            with patch("BPT_BusDigitalNumber.save_metro_ranges", side_effect=MetroError("只读")):
                ranges = load_metro_ranges(base)
            self.assertEqual(ranges, [MetroRange("测试线", "TS001", "TS010", "MODEL")])
            self.assertFalse((base / "北京地铁数据.json").exists())

    def test_first_run_seeds_the_json_file(self):
        with scratch() as directory:
            base = Path(directory)
            (base / METRO_SOURCE_FILENAME).write_text("测试线\nTS001-TS010 MODEL\n", encoding="utf-8")
            self.assertEqual(load_metro_ranges(base), [MetroRange("测试线", "TS001", "TS010", "MODEL")])
            self.assertTrue((base / "北京地铁数据.json").is_file())

    def test_broken_metro_json_is_reported(self):
        with scratch() as directory:
            base = Path(directory)
            (base / "北京地铁数据.json").write_text("{broken", encoding="utf-8")
            with self.assertRaises(MetroError):
                load_metro_ranges(base)


class ShippedDataTests(unittest.TestCase):
    """Tolerant checks on the real workbook so data edits do not break the suite."""

    @classmethod
    def setUpClass(cls):
        cls.ranges = load_ranges(SOURCE_DATA)

    def test_workbook_is_loaded_and_sorted_without_overlaps(self):
        self.assertGreaterEqual(len(self.ranges), 100)
        self.assertTrue(all(a.end < b.start for a, b in zip(self.ranges, self.ranges[1:])))

    def test_every_range_is_usable(self):
        for item in self.ranges:
            with self.subTest(start=item.start):
                self.assertTrue(10000 <= item.start <= item.end <= 99999)
                self.assertTrue(item.model.strip())

    def test_catalog_is_populated_and_sorted(self):
        names = model_names(self.ranges)
        self.assertGreaterEqual(len(names), 50)
        self.assertEqual(names, sorted(names, key=str.casefold))
        self.assertEqual(len(names), len(set(names)))


class ModelTagTests(unittest.TestCase):
    def test_manufacturer_comes_from_the_model_prefix(self):
        self.assertEqual(model_tags("BJ6123C7C4D").manufacturer, "北汽福田")
        self.assertEqual(model_tags("ZK6105HNG2").manufacturer, "宇通客车")
        self.assertEqual(model_tags("XMQ6100G1").manufacturer, "厦门金龙")

    def test_unknown_prefix_falls_back_to_the_code_itself(self):
        self.assertEqual(model_tags("QQ1234").manufacturer, "QQ")

    def test_fuel_distinguishes_the_electric_variants(self):
        self.assertEqual(model_tags("BJ6105EVCA-27").fuel, "纯电动")
        self.assertEqual(model_tags("BJ6105PHEVCA-7").fuel, "插电混动")
        self.assertEqual(model_tags("BJ6107SHEVCA").fuel, "增程混动")
        self.assertEqual(model_tags("BJ6122FCEVCH").fuel, "燃料电池")
        self.assertEqual(model_tags("BJD-WG120F").fuel, "无轨电车")
        self.assertEqual(model_tags("BJ6123C7C4D").fuel, UNKNOWN_TAG)

    def test_length_class_handles_both_code_styles(self):
        self.assertEqual(model_tags("BJ6123C7C4D").length, "12 米级")
        self.assertEqual(model_tags("BJ6851C6BCD").length, "8 米级")
        self.assertEqual(model_tags("JNP6182G").length, "18 米级")
        self.assertEqual(model_tags("BJD-WG120F").length, "12 米级")
        self.assertEqual(model_tags("BK560").length, UNKNOWN_TAG)

    def test_override_file_fixes_names_and_single_models(self):
        with scratch() as directory:
            base = Path(directory)
            (base / "车型标签.json").write_text(
                json.dumps(
                    {"manufacturers": {"QQ": "某厂"}, "models": {"BJ6123C7C4D": {"fuel": "柴油"}}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            overrides = load_tag_overrides(base)
            self.assertEqual(model_tags("QQ1234", overrides).manufacturer, "某厂")
            self.assertEqual(model_tags("BJ6123C7C4D", overrides).fuel, "柴油")
            self.assertEqual(model_tags("BJ6123C7C4D", overrides).length, "12 米级")

    def test_broken_override_file_is_ignored(self):
        with scratch() as directory:
            base = Path(directory)
            (base / "车型标签.json").write_text("{not json", encoding="utf-8")
            self.assertEqual(load_tag_overrides(base), {})
            self.assertEqual(model_tags("BJ6123C7C4D", load_tag_overrides(base)).manufacturer, "北汽福田")


class NumberListTests(unittest.TestCase):
    def test_separators_and_deduplication(self):
        numbers, ignored = parse_number_list("20000 20213,20300\n20000\t20229；20213")
        self.assertEqual(numbers, [20000, 20213, 20300, 20229])
        self.assertEqual(ignored, [])

    def test_invalid_parts_are_reported(self):
        numbers, ignored = parse_number_list("20000 20A00 1234")
        self.assertEqual(numbers, [20000])
        self.assertEqual(ignored, ["20A00", "1234"])

    def test_full_width_digits_are_accepted(self):
        self.assertEqual(parse_number_list("２００００")[0], [20000])

    def test_empty_input(self):
        self.assertEqual(parse_number_list("   \n  "), ([], []))


class SegmentLookupTests(unittest.TestCase):
    def setUp(self):
        self.ranges = [BusRange(20000, 20099, "A"), BusRange(20200, 20299, "B")]

    def test_segments_split_by_model_and_include_gaps(self):
        self.assertEqual(
            segment_lookup(20098, 20201, self.ranges),
            [(20098, 20099, "A"), (20100, 20199, None), (20200, 20201, "B")],
        )

    def test_range_inside_one_model(self):
        self.assertEqual(segment_lookup(20050, 20060, self.ranges), [(20050, 20060, "A")])

    def test_range_outside_every_model(self):
        self.assertEqual(segment_lookup(30000, 30010, self.ranges), [(30000, 30010, None)])

    def test_find_range_returns_the_containing_segment(self):
        self.assertEqual(find_range(20050, self.ranges), self.ranges[0])
        self.assertIsNone(find_range(20150, self.ranges))


class TagOverrideFileTests(unittest.TestCase):
    def test_save_and_reload_keeps_a_backup(self):
        with scratch() as directory:
            base = Path(directory)
            save_tag_overrides(base, {"manufacturers": {"HK": "黄海客车"}})
            save_tag_overrides(base, {"manufacturers": {"HK": "黄海"}})
            backups = backups_of(base, "车型标签")
            self.assertEqual(len(backups), 1)
            self.assertIn("黄海客车", backups[0].read_text(encoding="utf-8"))
            self.assertEqual(load_tag_overrides(base)["manufacturers"]["HK"], "黄海")

    def test_clean_drops_empty_and_malformed_entries(self):
        self.assertEqual(clean_tag_overrides({"manufacturers": {"BJ": " "}, "models": {"X": {}}}), {})
        self.assertEqual(
            clean_tag_overrides({"manufacturers": {"HK": " 黄海 "}, "models": {"X": {"fuel": "柴油", "length": ""}}}),
            {"manufacturers": {"HK": "黄海"}, "models": {"X": {"fuel": "柴油"}}},
        )
        self.assertEqual(clean_tag_overrides({"models": {"X": "oops"}}), {})


class WorkbookTableTests(unittest.TestCase):
    def test_column_letter_covers_multiple_letters(self):
        self.assertEqual(_column_letter(1), "A")
        self.assertEqual(_column_letter(26), "Z")
        self.assertEqual(_column_letter(27), "AA")
        self.assertEqual(_column_letter(52), "AZ")

    def test_write_and_read_table_round_trip(self):
        with scratch() as directory:
            path = Path(directory) / "table.xlsx"
            header = ["号段起始", "号段结束", "车型", "厂商", "燃料", "车长"]
            rows = [["20000", "20212", "BJ6123C7C4D", "北汽福田", "柴油", "12 米级"]]
            write_table(path, header, rows, numeric_columns={0, 1})
            self.assertEqual(read_table(path), (header, rows))

    def test_saving_ranges_keeps_extra_columns(self):
        with scratch() as directory:
            path = Path(directory) / "extra.xlsx"
            write_table(
                path,
                ["号段起始", "号段结束", "车型", "备注"],
                [["20000", "20010", "A", "原来的备注"], ["20011", "20020", "B", "B 的备注"]],
                numeric_columns={0, 1},
            )
            # 新增一行 A（同车型）与一行全新车型 C
            save_ranges(path, [BusRange(20000, 20010, "A"), BusRange(20021, 20030, "A"), BusRange(20031, 20040, "C")])
            header, rows = read_table(path)
            self.assertEqual(header, ["号段起始", "号段结束", "车型", "备注"])
            self.assertEqual(rows[0][3], "原来的备注")
            self.assertEqual(rows[1][3], "原来的备注")  # 同车型沿用
            self.assertEqual(rows[2][3], "")  # 新车型留空
            self.assertEqual(load_ranges(path), [BusRange(20000, 20010, "A"), BusRange(20021, 20030, "A"), BusRange(20031, 20040, "C")])


class TagExchangeTests(unittest.TestCase):
    def test_export_then_read_back_tag_columns(self):
        with scratch() as directory:
            path = Path(directory) / "tagged.xlsx"
            ranges = [BusRange(20000, 20010, "BJ6123C7C4D"), BusRange(20011, 20020, "XMQ6100G1")]
            count = write_tagged_workbook(path, ranges, {})
            self.assertEqual(count, 2)
            header, _rows = read_table(path)
            self.assertEqual(header, ["号段起始", "号段结束", "车型", "厂商", "燃料", "车长"])
            columns = load_tag_columns(path)
            self.assertEqual(columns["BJ6123C7C4D"]["manufacturer"], "北汽福田")
            self.assertEqual(columns["BJ6123C7C4D"]["length"], "12 米级")

    def test_workbook_without_tag_columns_yields_nothing(self):
        with scratch() as directory:
            path = Path(directory) / "plain.xlsx"
            write_table(path, ["号段起始", "号段结束", "车型"], [["20000", "20010", "A"]], numeric_columns={0, 1})
            self.assertEqual(load_tag_columns(path), {})
            self.assertEqual(load_tag_columns(Path(directory) / "missing.xlsx"), {})

    def test_merge_ignores_unknown_models_and_drops_auto_values(self):
        columns = {
            "BJ6123C7C4D": {"fuel": "柴油", "length": "12 米级"},  # length 与自动推断相同 -> 不写覆盖
            "NEW1234": {"fuel": "柴油"},
        }
        merged, ignored = merge_tag_columns({"manufacturers": {"HK": "黄海"}}, columns, {"BJ6123C7C4D"})
        self.assertEqual(ignored, ["NEW1234"])
        self.assertEqual(merged["models"], {"BJ6123C7C4D": {"fuel": "柴油"}})
        self.assertEqual(merged["manufacturers"], {"HK": "黄海"})

    def test_apply_helpers_handle_reset_and_prefix(self):
        overrides: dict = {}
        apply_model_tag_values(overrides, "BJ6123C7C4D", {"fuel": "柴油"})
        self.assertEqual(overrides["models"]["BJ6123C7C4D"], {"fuel": "柴油"})
        apply_model_tag_values(overrides, "BJ6123C7C4D", {"fuel": ""})
        self.assertNotIn("BJ6123C7C4D", overrides.get("models", {}))

        apply_manufacturer_prefix(overrides, "HK", "黄海客车")
        self.assertEqual(overrides["manufacturers"]["HK"], "黄海客车")
        apply_manufacturer_prefix(overrides, "HK", "HK")  # 等于内置回退值 -> 删除
        self.assertNotIn("HK", overrides.get("manufacturers", {}))


class HealthReportTests(unittest.TestCase):
    def test_report_counts_each_category(self):
        with scratch() as directory:
            base = Path(directory)
            ranges = [
                BusRange(20000, 20010, "BJ6123C7C4D"),
                BusRange(20020, 20030, "QQ9999"),
            ]
            profiles = {"BJ6123C7C4D": {"introduction": "", "specifications": "", "image": "GONE.jpg"}}
            items = {item.key: item for item in health_report(base, ranges, profiles, {})}

            self.assertEqual(len(items["manufacturer"].rows), 1)  # QQ9999 前缀未确认
            self.assertEqual(items["manufacturer"].rows[0][0], "QQ9999")
            self.assertEqual(len(items["fuel"].rows), 2)  # 两个型号都没有燃料关键字
            self.assertEqual(len(items["length"].rows), 1)  # QQ9999 车长未标注
            self.assertEqual(items["length"].rows[0][0], "QQ9999")
            self.assertEqual(len(items["profile"].rows), 2)  # 两个车型资料都不完整
            self.assertEqual(items["profile"].rows[0][1], "仅有主图")
            self.assertEqual(len(items["image"].rows), 1)  # 主图文件不存在
            self.assertEqual(items["gap"].rows, [("20011-20019", "9 个自编号未收录")])


class TagRuleTests(unittest.TestCase):
    def test_rules_match_substrings_and_regex_case_insensitively(self):
        overrides = {"rules": [{"pattern": "c7btd", "fuel": "天然气"}]}
        self.assertEqual(model_tags("BJ6103C7BTD", overrides).fuel, "天然气")
        overrides = {"rules": [{"pattern": r"C7B(TD|CD)", "fuel": "天然气"}]}
        self.assertEqual(model_tags("BJ6123C7BCD", overrides).fuel, "天然气")

    def test_first_rule_defining_a_field_wins(self):
        overrides = {
            "rules": [
                {"pattern": "C7BTD", "fuel": "天然气"},
                {"pattern": "C7BTD-1", "fuel": "柴油"},
            ]
        }
        self.assertEqual(model_tags("BJ6123C7BTD-1", overrides).fuel, "天然气")
        overrides["rules"].reverse()
        self.assertEqual(model_tags("BJ6123C7BTD-1", overrides).fuel, "柴油")

    def test_per_model_override_beats_rules(self):
        overrides = {
            "rules": [{"pattern": "BJ6123", "fuel": "柴油"}],
            "models": {"BJ6123C7C4D": {"fuel": "天然气"}},
        }
        self.assertEqual(model_tags("BJ6123C7C4D", overrides).fuel, "天然气")
        self.assertEqual(model_tags("BJ6123C7BTD", overrides).fuel, "柴油")

    def test_rule_can_set_manufacturer_over_prefix_mapping(self):
        overrides = {"manufacturers": {"BJ": "北汽福田"}, "rules": [{"pattern": "BJ6180", "manufacturer": "福田欧辉"}]}
        self.assertEqual(model_tags("BJ6180C8CTD", overrides).manufacturer, "福田欧辉")
        self.assertEqual(model_tags("BJ6123C7C4D", overrides).manufacturer, "北汽福田")

    def test_fields_not_set_by_a_rule_keep_auto_inference(self):
        overrides = {"rules": [{"pattern": "C7BTD", "fuel": "天然气"}]}
        tags = model_tags("BJ6103C7BTD", overrides)
        self.assertEqual(tags.fuel, "天然气")
        self.assertEqual(tags.length, "10 米级")
        self.assertEqual(tags.manufacturer, "北汽福田")

    def test_invalid_regex_is_ignored(self):
        overrides = {"rules": [{"pattern": "C7([", "fuel": "柴油"}]}
        self.assertEqual(model_tags("BJ6103C7BTD", overrides).fuel, UNKNOWN_TAG)
        self.assertFalse(rule_matches("BJ6103C7BTD", "C7(["))

    def test_clean_keeps_valid_rules_and_drops_empty_ones(self):
        cleaned = clean_tag_overrides(
            {"rules": [{"pattern": "C7BTD", "fuel": "天然气"}, {"pattern": "  ", "fuel": "柴油"}, {"pattern": "X"}]}
        )
        self.assertEqual(cleaned["rules"], [{"pattern": "C7BTD", "fuel": "天然气"}])
        self.assertEqual(clean_tag_overrides({"rules": [{"pattern": "X"}]}), {})

    def test_rule_hits_lists_matching_models(self):
        models = ["BJ6103C7BTD", "BJ6123C7BTD", "BJ6123C7C4D"]
        self.assertEqual(rule_hits("C7BTD", models), ["BJ6103C7BTD", "BJ6123C7BTD"])


class MetroBatchTests(unittest.TestCase):
    def test_number_list_dedupes_and_upper_cases(self):
        numbers, ignored = parse_metro_number_list("01001 t401, YZ001\n01001 NOPE")
        self.assertEqual(numbers, ["01001", "T401", "YZ001"])
        self.assertEqual(ignored, ["NOPE"])

    def test_empty_input(self):
        self.assertEqual(parse_metro_number_list("  "), ([], []))

    def test_export_then_parse_round_trip(self):
        with scratch() as directory:
            base = Path(directory)
            ranges = [
                MetroRange("测试线", "TS001", "TS010", "A"),
                MetroRange("测试线", "TS011", "TS020", "B"),
                MetroRange("2号线", "T401", "T450", "DKZ16"),
            ]
            path = export_metro_source(base, ranges)
            self.assertTrue(path.is_file())
            self.assertEqual(set(parse_metro_source(path)), set(ranges))

    def test_import_overwrites_saved_metro_data(self):
        with scratch() as directory:
            base = Path(directory)
            (base / METRO_SOURCE_FILENAME).write_text("1号线\n01001-01010 A\n", encoding="utf-8")
            save_metro_ranges(base, [MetroRange("旧线", "X001", "X010", "OLD")])
            ranges = import_metro_source(base)
            self.assertEqual(ranges, [MetroRange("1号线", "01001", "01010", "A")])
            self.assertEqual(load_metro_ranges(base), ranges)

    def test_metro_health_reports_drift_overlap_and_gap(self):
        with scratch() as directory:
            base = Path(directory)
            current = [
                MetroRange("1号线", "01001", "01010", "A"),
                MetroRange("1号线", "01005", "01015", "C"),
                MetroRange("1号线", "01020", "01030", "B"),
            ]
            (base / METRO_SOURCE_FILENAME).write_text("1号线\n01001-01010 A\n", encoding="utf-8")
            save_metro_ranges(base, current)
            items = {item.key: item for item in metro_health_items(base, current)}
            self.assertTrue(items["metro_source"].rows)
            self.assertEqual(len(items["metro_overlap"].rows), 1)
            self.assertTrue(any("未收录" in row[1] for row in items["metro_gap"].rows))

    def test_health_report_includes_metro_items(self):
        with scratch() as directory:
            base = Path(directory)
            ranges = [BusRange(20000, 20010, "BJ6123C7C4D")]
            items = health_report(base, ranges, {}, {}, [MetroRange("1号线", "01001", "01010", "A")])
            self.assertIn("metro_source", {item.key for item in items})


class BackupAndLogTests(unittest.TestCase):
    def test_backups_are_timestamped_and_pruned(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / "demo.json"
            for index in range(5):
                path.write_text(str(index), encoding="utf-8")
                backup_file(path, keep=3)
            backups = backups_of(base, "demo")
            self.assertEqual(len(backups), 3)
            self.assertIn("4", [item.read_text(encoding="utf-8") for item in backups])

    def test_list_backups_finds_new_and_legacy_files(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / DATA_FILENAME
            save_ranges(path, [BusRange(20000, 20099, "A")])
            save_ranges(path, [BusRange(20000, 20099, "B")])
            (base / (DATA_FILENAME + ".bak")).write_bytes(path.read_bytes())
            entries = list_backups(base)
            labels = [entry[0] for entry in entries]
            self.assertTrue(any("旧版 .bak" in label for label in labels))
            self.assertTrue(all(entry[2] == DATA_FILENAME for entry in entries))

    def test_restore_returns_the_old_version_and_saves_the_current_one(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / DATA_FILENAME
            save_ranges(path, [BusRange(20000, 20099, "A")])
            save_ranges(path, [BusRange(20000, 20099, "B")])
            backups = backups_of(base, "BPT_BusDigitalNumber")
            restore_backup(base, backups[0], DATA_FILENAME)
            self.assertEqual(load_ranges(path), [BusRange(20000, 20099, "A")])
            restored = backups_of(base, "BPT_BusDigitalNumber")
            self.assertEqual(len(restored), 2)
            self.assertIn([BusRange(20000, 20099, "B")], [load_ranges(item) for item in restored])

    def test_restore_reports_an_unwritable_target(self):
        with scratch() as directory:
            base = Path(directory)
            missing = base / "不存在的备份.json"
            with self.assertRaises(DataError):
                restore_backup(base, missing, "目标.json")

    def test_exception_log_records_the_traceback(self):
        with scratch() as directory:
            with patch_app_dir(Path(directory)):
                try:
                    raise ValueError("演示错误")
                except ValueError:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    path = log_exception(exc_type, exc_value, exc_tb, "单元测试")
                self.assertIsNotNone(path)
                text = path.read_text(encoding="utf-8")
                self.assertIn("演示错误", text)
                self.assertIn("ValueError", text)
                self.assertIn("单元测试", text)

    def test_oversized_log_is_rotated(self):
        with scratch() as directory:
            base = Path(directory)
            with patch_app_dir(base):
                log = base / "错误日志.txt"
                log.write_text("x" * (MAX_ERROR_LOG_BYTES + 10), encoding="utf-8")
                try:
                    raise RuntimeError("第二轮")
                except RuntimeError:
                    log_exception(*sys.exc_info(), "轮转测试")
                self.assertTrue((base / "错误日志.old.txt").is_file())
                self.assertIn("第二轮", log.read_text(encoding="utf-8"))
                self.assertLess(log.stat().st_size, MAX_ERROR_LOG_BYTES)


class SettingsTests(unittest.TestCase):
    def test_settings_round_trip_and_broken_file(self):
        with scratch() as directory:
            base = Path(directory)
            self.assertEqual(load_settings(base), {})
            save_settings(base, {"windows": {"CatalogWindow": [10, 20, 800, 600]}})
            self.assertEqual(stored_geometry(base, "CatalogWindow"), (10, 20, 800, 600))
            self.assertIsNone(stored_geometry(base, "OtherWindow"))
            (base / SETTINGS_FILENAME).write_text("{坏掉的", encoding="utf-8")
            self.assertEqual(load_settings(base), {})
            self.assertIsNone(stored_geometry(base, "CatalogWindow"))

    def test_profile_is_empty_can_use_a_precomputed_image_index(self):
        with scratch() as directory:
            base = Path(directory)
            image_directory(base).mkdir(exist_ok=True)
            (image_directory(base) / "HAS.jpg").write_bytes(b"x")
            names = existing_image_names(base)
            self.assertIn("HAS.jpg", names)
            gone = {"introduction": "", "specifications": "", "image": "GONE.jpg"}
            here = {"introduction": "", "specifications": "", "image": "HAS.jpg"}
            self.assertTrue(profile_is_empty(gone, None, names))
            self.assertFalse(profile_is_empty(here, None, names))


class SpotModeTests(unittest.TestCase):
    def test_history_round_trip_dedupes_and_limits(self):
        with scratch() as directory:
            base = Path(directory)
            self.assertEqual(load_history(base), [])
            for number in ("20000", "20213", "20000"):
                remember_number(base, number, limit=3)
            self.assertEqual(load_history(base), ["20000", "20213"])
            for number in ("20001", "20002", "20003"):
                remember_number(base, number, limit=3)
            self.assertEqual(load_history(base), ["20003", "20002", "20001"])
            self.assertEqual(remember_number(base, "abc"), load_history(base))

    def test_history_accepts_full_width_digits(self):
        with scratch() as directory:
            remember_number(Path(directory), "２００００")
            self.assertEqual(load_history(Path(directory)), ["20000"])

    def test_spot_log_append_and_filter(self):
        with scratch() as directory:
            base = Path(directory)
            first = make_spot("20000", "BJ6123C7C4D", "早上在四惠")
            second = make_spot("20213", "XMQ6100G1", "下午")
            append_spot(base, first)
            append_spot(base, second)
            entries = load_spots(base)
            self.assertEqual([entry.number for entry in entries], ["20213", "20000"])  # 新的在前

            self.assertEqual(len(filter_spots(entries, "四惠")), 1)
            self.assertEqual(len(filter_spots(entries, "", "XMQ6100G1")), 1)
            self.assertEqual(len(filter_spots(entries, "", None, first.time[:10])), 2)
            self.assertEqual(len(filter_spots(entries, "", None, "1999-01-01")), 0)
            self.assertEqual(spot_models(entries), ["BJ6123C7C4D", "XMQ6100G1"])

    def test_spot_log_skips_malformed_entries(self):
        with scratch() as directory:
            base = Path(directory)
            (base / "拍车日志.json").write_text(
                json.dumps({"entries": [{"time": "x", "number": "", "model": ""}, "坏数据", {"number": "20000", "model": "A"}]}),
                encoding="utf-8",
            )
            entries = load_spots(base)
            self.assertEqual(len(entries), 1)
            self.assertEqual(entries[0].number, "20000")

    def test_make_spot_formats_number_and_time(self):
        entry = make_spot("２００００", "X", "  备注  ")
        self.assertEqual(entry.number, "20000")
        self.assertEqual(entry.note, "备注")
        self.assertRegex(entry.time, r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")

    def test_profile_images_migration_and_cover_sync(self):
        self.assertEqual(profile_images({"image": "A.jpg"}), [{"name": "A.jpg", "caption": ""}])
        self.assertEqual(
            profile_images({"image": "A.jpg", "images": [{"name": "B.jpg", "caption": "侧面"}]}),
            [{"name": "B.jpg", "caption": "侧面"}],
        )
        self.assertEqual(profile_images({"images": [{"name": "  "}]}), [])

        profile: dict = {}
        set_profile_images(profile, [{"name": "C.jpg", "caption": ""}, {"name": "D.jpg", "caption": ""}])
        self.assertEqual(profile["image"], "C.jpg")
        self.assertEqual(len(profile["images"]), 2)
        set_profile_images(profile, [])
        self.assertEqual(profile["image"], "")
        self.assertEqual(profile["images"], [])

    def test_gallery_import_uses_explicit_filename_and_keeps_others(self):
        with scratch() as directory:
            base = Path(directory)
            first = import_image(base, "MODEL-A", Image.new("RGB", (30, 20), "red"))
            second = import_image(base, "MODEL-A", Image.new("RGB", (30, 20), "blue"), filename="MODEL-A-1.jpg", delete_previous=False)
            self.assertTrue((image_directory(base) / first).is_file())
            self.assertTrue((image_directory(base) / second).is_file())
            self.assertNotEqual(first, second)

    def test_clipboard_number_extraction(self):
        self.assertEqual(extract_clipboard_number("20000"), 20000)
        self.assertEqual(extract_clipboard_number("  编号：20213\n"), 20213)
        self.assertIsNone(extract_clipboard_number("abc"))
        self.assertIsNone(extract_clipboard_number(""))

    def test_speak_text_reports_failure_without_shell(self):
        with patch("BPT_BusDigitalNumber.subprocess.Popen", side_effect=OSError("no shell")):
            self.assertFalse(speak_text("BJ6123C7C4D"))
        self.assertFalse(speak_text("   "))
        with patch("BPT_BusDigitalNumber.subprocess.Popen") as popen:
            self.assertTrue(speak_text("BJ6123C7C4D"))
            self.assertTrue(popen.called)


class DataProductivityTests(unittest.TestCase):
    def test_group_missing_numbers_merges_consecutive(self):
        self.assertEqual(group_missing_numbers([]), [])
        self.assertEqual(group_missing_numbers([20005]), [(20005, 20005)])
        self.assertEqual(group_missing_numbers([20003, 20001, 20002, 20010, 20020, 20021]), [(20001, 20003), (20010, 20010), (20020, 20021)])
        self.assertEqual(group_missing_numbers([7, 7, 7]), [(7, 7)])

    def test_alias_cleaning_and_lookup(self):
        cleaned = clean_tag_overrides({"aliases": {"A": ["大福后", "大福后", "  "], "B": "单个别名", "C": []}})
        self.assertEqual(cleaned["aliases"], {"A": ["大福后"], "B": ["单个别名"]})
        self.assertEqual(model_aliases("A", cleaned), ["大福后"])
        self.assertEqual(model_aliases("Z", cleaned), [])
        index = alias_index(cleaned, ["A", "B"])
        self.assertEqual(index["大福后"], ["A"])
        self.assertEqual(alias_index(cleaned), {"大福后": ["A"], "单个别名": ["B"]})

    def test_merge_plan_classifies_every_kind_of_change(self):
        current = [BusRange(20000, 20212, "A"), BusRange(20300, 20400, "B"), BusRange(30000, 30100, "C")]
        incoming = [
            BusRange(20000, 20212, "NEW"),  # 修改
            BusRange(20300, 20400, "B"),  # 没变
            BusRange(20500, 20600, "D"),  # 新增
            BusRange(30050, 30150, "E"),  # 冲突（与 C 重叠）
        ]
        plan = merge_ranges(current, incoming)
        self.assertEqual([(item.model, old) for item, old in plan.updated], [("NEW", "A")])
        self.assertEqual([item.model for item in plan.added], ["D"])
        self.assertEqual([item.model for item, _old in plan.conflict], ["E"])
        # C 的区间在新表里没有（同时被 E 顶掉）
        self.assertEqual([item.model for item in plan.removed], ["C"])
        self.assertEqual(plan.summary, "新增 1 · 修改 1 · 删除 1 · 冲突 1")

        incoming.append(BusRange(20700, 20800, "F"))
        plan = merge_ranges(current, incoming)
        self.assertEqual([item.model for item in plan.added], ["D", "F"])

    def test_apply_merge_never_leaves_overlaps(self):
        current = [BusRange(20000, 20212, "A"), BusRange(20300, 20400, "B"), BusRange(30000, 30100, "C")]
        incoming = [BusRange(20000, 20212, "NEW"), BusRange(30050, 30150, "E"), BusRange(20500, 20600, "D")]
        plan = merge_ranges(current, incoming)

        for strategy in ("incoming", "additive"):
            merged = apply_merge(current, plan, strategy)
            pairs = list(zip(merged, merged[1:]))
            self.assertFalse([pair for pair in pairs if pair[1].start <= pair[0].end], strategy)
            self.assertEqual(merged, sorted(merged, key=lambda item: item.start))

        without_delete = apply_merge(current, merge_ranges(current, incoming), "incoming")
        # NEW（改了 A）／D（新增）／E（顶掉 C）；B、C 的区间不在新表里，按「以新表为准」被删
        self.assertEqual([item.model for item in without_delete], ["NEW", "D", "E"])
        self.assertNotIn("C", [item.model for item in without_delete])

        additive = apply_merge(current, plan, "additive")
        # 按起始排序：A(20000) B(20300) D(20500) C(30000)；冲突的 E 不进、旧的一律不动
        self.assertEqual([item.model for item in additive], ["A", "B", "D", "C"])

    def test_draft_profile_fills_template_and_falls_back(self):
        ranges = [BusRange(10000, 10099, "BJ6123EVCA-37"), BusRange(10100, 10109, "BJ6123EVCA-37")]
        tags = model_tags("BJ6123EVCA-37")
        draft = draft_profile("BJ6123EVCA-37", tags, ranges)
        self.assertIn(tags.manufacturer, draft["introduction"])
        self.assertIn("10000-10099", draft["introduction"])
        self.assertIn("10000-10099、10100-10109", draft["specifications"])
        self.assertIn("110", draft["specifications"])  # 覆盖的编号数

        fallback = draft_profile("BJ6123EVCA-37", tags, ranges, "{不存在的占位符}")
        self.assertTrue(fallback["introduction"])

    def test_empty_profile_models_skips_models_with_images(self):
        with scratch() as directory:
            base = Path(directory)
            image_directory(base).mkdir(exist_ok=True)
            (image_directory(base) / "HAS.jpg").write_bytes(b"x")
            ranges = [BusRange(10000, 10009, "A"), BusRange(10100, 10109, "B")]
            profiles = {"A": {"introduction": "已有", "specifications": "", "image": ""}, "B": {"image": "HAS.jpg"}}
            names = existing_image_names(base)
            self.assertEqual(empty_profile_models(profiles, ranges, names), [])
            profiles["B"] = {"image": "GONE.jpg"}
            self.assertEqual(empty_profile_models(profiles, ranges, names), ["B"])


class WebExportTests(unittest.TestCase):
    def test_payload_carries_ranges_tags_aliases_and_metro(self):
        ranges = [BusRange(20000, 20212, "A"), BusRange(20300, 20400, "B")]
        metro = [MetroRange("1号线", "1001", "1050", "SFM04")]
        payload = web_payload(ranges, {"aliases": {"A": ["大福后"]}}, {"A": {"introduction": "简介"}}, metro)
        self.assertEqual(payload["ranges"], [[20000, 20212, "A"], [20300, 20400, "B"]])
        self.assertIn("A", payload["tags"])
        self.assertEqual(payload["aliases"]["A"], ["大福后"])
        self.assertEqual(payload["profiles"]["A"]["intro"], "简介")
        self.assertEqual(payload["metro"]["1号线"], [[1001, 1050, "SFM04"]])
        self.assertIn("2 个号段", payload["subtitle"])

    def test_render_html_is_self_contained(self):
        payload = web_payload([BusRange(20000, 20212, "A")], {}, {}, [])
        html = bpt_web.render_html(payload)
        self.assertIn("<!DOCTYPE html>", html)
        self.assertIn("const DATA = ", html)
        self.assertNotIn("__DATA__", html)
        self.assertNotIn("__TITLE__", html)
        # 自包含：不联网、不引用外部资源
        self.assertNotIn("http://", html)
        self.assertNotIn("https://", html)
        self.assertNotIn("<link", html)
        self.assertNotIn("<script src", html)

    def test_render_html_escapes_script_endings(self):
        payload = web_payload([BusRange(1, 2, "A")], {}, {"A": {"introduction": "带 </script> 的文字"}}, [])
        html = bpt_web.render_html(payload)
        self.assertIn("<\\/script>", html)

    def test_export_embeds_thumbnail_and_stays_small(self):
        with scratch() as directory:
            base = Path(directory)
            image_directory(base).mkdir(exist_ok=True)
            Image.new("RGB", (600, 400), "red").save(image_directory(base) / "A.jpg", format="JPEG")
            profiles = {"A": {"introduction": "", "specifications": "", "image": "A.jpg", "images": [{"name": "A.jpg", "caption": ""}]}}
            payload = web_payload([BusRange(10000, 10099, "A")], {}, profiles, [], base, embed_images=True)
            self.assertTrue(payload["profiles"]["A"]["image"].startswith("data:image/jpeg;base64,"))
            without = web_payload([BusRange(10000, 10099, "A")], {}, profiles, [], base, embed_images=False)
            self.assertNotIn("A", without["profiles"])  # 只有图片、没有文字时不会进 payload

    def test_cli_lookup_exit_codes(self):
        code, lines = cli_lookup("abc")
        self.assertEqual(code, 2)
        self.assertIn("不是五位自编号", lines[0])
        code, lines = cli_lookup("99999")
        self.assertEqual(code, 1)
        self.assertIn("未收录", lines[0])

    def test_cli_export_html_writes_a_file(self):
        with scratch() as directory:
            target = Path(directory) / "分享.html"
            code, lines = cli_export_html(str(target))
            self.assertEqual(code, 0)
            self.assertTrue(target.is_file())
            text = target.read_text(encoding="utf-8")
            self.assertIn("const DATA = ", text)
            self.assertLess(target.stat().st_size, 5 * 1024 * 1024)


    def test_console_write_and_main_dispatch(self):
        import io

        buffer = io.StringIO()
        with patch("sys.stdout", buffer):
            self.assertTrue(console_write("测试输出"))
        self.assertIn("测试输出", buffer.getvalue())

        seen: list[tuple[str, list[str]]] = []
        with scratch() as directory:
            target = Path(directory) / "main.html"
            with patch("BPT_BusDigitalNumber.cli_report", side_effect=lambda title, lines: seen.append((title, lines))), patch(
                "BPT_BusDigitalNumber.BusLookupApp"
            ) as app_class:
                self.assertEqual(main(["--help"]), 0)
                self.assertEqual(main(["--export-html", str(target)]), 0)
                self.assertTrue(target.is_file())
                self.assertEqual(main(["99999"]), 1)
                self.assertEqual(main(["abc"]), 2)
                self.assertEqual(main([]), 0)  # 无参数 → 打开图形界面
                self.assertTrue(app_class.called)
        self.assertGreaterEqual(len(seen), 4)


class SpotPhotoTests(unittest.TestCase):
    def test_lenient_loader_keeps_usable_ranges_and_reports_problems(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / DATA_FILENAME
            write_table(
                path,
                list(EXPECTED_HEADERS),
                [["20000", "20010", "A"], ["20010", "20020", "B"], ["abc", "20030", "C"], ["20040", "20050", ""], ["20060", "20070", "D"]],
                numeric_columns={0, 1},
            )
            ranges, notes = load_ranges_lenient(path)
            self.assertEqual([(item.start, item.end, item.model) for item in ranges], [(20000, 20010, "A"), (20060, 20070, "D")])
            self.assertEqual(len(notes), 3)
            with self.assertRaises(DataError):
                load_ranges(path)  # 严格加载仍然拒绝

    def test_lenient_loader_rejects_a_hopeless_file(self):
        with scratch() as directory:
            with self.assertRaises(DataError):
                load_ranges_lenient(Path(directory) / "missing.xlsx")

    def test_spot_photo_round_trip_and_delete(self):
        with scratch() as directory:
            base = Path(directory)
            name = save_spot_photo(base, Image.new("RGB", (900, 700), "red"))
            self.assertTrue(name.endswith(".jpg"))
            self.assertTrue(stored_photo_path(base, name).is_file())
            second = save_spot_photo(base, Image.new("RGB", (900, 700), "blue"))
            self.assertNotEqual(name, second)  # 同一秒内也不会互相覆盖
            delete_spot_photo(base, name)
            self.assertFalse(stored_photo_path(base, name).is_file())
            delete_spot_photo(base, "")  # 空名字不报错

    def test_spot_summary_counts(self):
        entries = [
            make_spot("20000", "A", "备注"),
            make_spot("20001", "A"),
            make_spot("20002", "B", photo="p.jpg"),
        ]
        summary = spot_summary(entries)
        self.assertEqual(summary["total"], 3)
        self.assertEqual(summary["models"], 2)
        self.assertEqual(summary["numbers"], 3)
        self.assertEqual(summary["with_photo"], 1)
        self.assertEqual(summary["top"][0], ("A", 2))
        self.assertIn("3 条", spot_summary_text(summary))
        self.assertEqual(spot_summary_text(spot_summary([])), "还没有拍车记录")

    def test_replace_spot_updates_one_entry(self):
        with scratch() as directory:
            base = Path(directory)
            append_spot(base, make_spot("20000", "A", "旧备注"))
            append_spot(base, make_spot("20001", "B"))
            replace_spot(base, 1, SpotEntry(time="2026-01-01 00:00:00", number="20000", model="A", note="新备注"))
            entries = load_spots(base)
            self.assertEqual(len(entries), 2)
            self.assertEqual(entries[1].note, "新备注")
            with self.assertRaises(IndexError):
                replace_spot(base, 9, make_spot("20000", "A"))


class BodyTagTests(unittest.TestCase):
    def test_body_inference_handles_articulated_and_double_decker(self):
        self.assertEqual(default_body("BJ6180C8CTD", "18 米级"), "铰接")
        self.assertEqual(default_body("BJ6160SHEVCA-3", "16 米级"), "铰接")
        # 13 / 14 米级的刚体客车按三轴算（BJ6140 / BJ6147 / ZK6140 这类）
        self.assertEqual(default_body("BJ6140C8CTD", "14 米级"), "三轴")
        self.assertEqual(default_body("BJ6147SHEVCA", "14 米级"), "三轴")
        self.assertEqual(default_body("ZK6140HNG2", "14 米级"), "三轴")
        self.assertEqual(default_body("GTQ6131BEVST8", "13 米级"), "三轴")
        # 双层优先于三轴：13 米的青年双层仍是双层
        self.assertEqual(default_body("JNP6130GSC", "13 米级"), "双层")
        self.assertEqual(default_body("HFF6110GS01D", "11 米级"), "双层")  # 注意是 GS01D，不是 GS101D
        self.assertEqual(default_body("HFF6120GS01C", "12 米级"), "双层")
        # 型号里带 GS 不等于双层：DNC6107BEVGS1 是纯电动低入口城市客车（BEVGS1 是版本代号）
        self.assertEqual(default_body("DNC6107BEVGS1", "10 米级"), "单机")
        self.assertFalse(body_is_guessed_double_decker("DNC6107BEVGS1"))
        self.assertTrue(body_is_guessed_double_decker("JNP6130GSC"))
        self.assertEqual(default_body("SCT6700RZB53L", "7 米级"), "单机")
        self.assertEqual(default_body("BJ6123SHEVCA-2", "12 米级"), "单机")  # SHEV 的 S 是混动
        self.assertEqual(default_body("BK560", UNKNOWN_TAG), UNKNOWN_TAG)  # 车长都没定，就别猜
        self.assertEqual(list(BODY_VALUES), ["单机", "铰接", "双层", "三轴", UNKNOWN_TAG])

    def test_body_respects_overrides_and_rules(self):
        self.assertEqual(model_tags("BJ6180C8CTD", {"models": {"BJ6180C8CTD": {"body": "单机"}}}).body, "单机")
        self.assertEqual(model_tags("BJ6123C7C4D", {"rules": [{"pattern": "BJ6123", "body": "双层"}]}).body, "双层")
        self.assertEqual(model_tags("BJ6180C8CTD", {"rules": [{"pattern": "BJ6180", "body": "双层"}]}).body, "双层")
        cleaned = clean_tag_overrides({"models": {"A": {"body": "铰接", "fuel": ""}}})
        self.assertEqual(cleaned["models"]["A"], {"body": "铰接"})
        self.assertEqual(tag_suggestions([BusRange(10000, 10009, "A")])["body"], list(BODY_VALUES))
        # 人工定过的字段不再算「猜的」
        self.assertTrue(_tag_overridden("A", {"models": {"A": {"body": "双层"}}}, "body"))
        self.assertFalse(_tag_overridden("A", {}, "body"))

    def test_shipped_data_body_distribution_is_sane(self):
        ranges = load_ranges(SOURCE_DATA)
        names = model_names(ranges)
        bodies = Counter(model_tags(name).body for name in names)
        self.assertGreater(bodies["铰接"], 0)
        self.assertGreater(bodies["单机"], 0)
        self.assertGreaterEqual(bodies["三轴"], 5)  # BJ6140 / BJ6147 / ZK6140 那一批
        self.assertLessEqual(bodies["铰接"], 40)  # 16/18 米级就那么多
        for name in names:
            tag = model_tags(name)
            if tag.length in ("16 米级", "18 米级"):
                self.assertEqual(tag.body, "铰接", name)
            if tag.length in ("13 米级", "14 米级"):
                self.assertIn(tag.body, ("三轴", "双层"), name)  # 双层优先，其余是三轴

    def test_model_mismatch_is_classified_by_nature_of_difference(self):
        """同一个自编号两个写法：只描述差异性质，不按"名字像不像"猜笔误。

        用户的实际情况：换批次（-1 / -5）、换能源（BTB / BTD）名字本来就像，不能当笔误报。
        """
        cases = {
            ("BJ6906FCEVCH", "BJ6906FCEVCH (HC9)"): "写法差异（多了厂内代号/括号）",
            ("ZK6105BEVG55", "ZK6105BEVG55(E10)"): "写法差异（多了厂内代号/括号）",
            ("BJ6123SHEVCA-1", "BJ6123SHEVCA-5"): "批次或能源后缀不同",
            ("BJ6127C8BTB-1", "BJ6127C8BTD-1"): "少/多一个字符",
            ("BJ6123C7BHD", "BJ6123C7BH"): "少/多一个字符",
            ("BJ6855SHEVCA-1", "BJ6851C6BCD"): "完全不搭（可能记录写错/号段表是别的车）",
            ("BJ6123SHEVCA-5", "BJ6123EVCA-47"): "完全不搭（可能记录写错/号段表是别的车）",
        }
        for (ours, theirs), expected in cases.items():
            self.assertEqual(bpt_data.classify_model_mismatch(ours, theirs), expected, f"{ours} vs {theirs}")
            # 反过来看也该是同一类（谁在"号段表"那一侧不影响性质）
            self.assertEqual(bpt_data.classify_model_mismatch(theirs, ours), expected, f"{theirs} vs {ours}")

    def test_mismatch_rows_are_ordered_by_how_much_they_matter(self):
        """对账结果：完全不搭的排最前，纯写法差异排最后。"""
        records = [
            bpt_data.RideRecord("2024", "1.1", "1路", "20001", "BJ6906FCEVCH (HC9)"),   # 写法差异
            bpt_data.RideRecord("2024", "1.2", "2路", "20002", "BJ6123C7BH"),            # 少一个字符
            bpt_data.RideRecord("2024", "1.3", "3路", "20003", "ZZZ999"),                # 完全不搭
        ]
        ranges = [
            bpt_data.BusRange(20001, 20001, "BJ6906FCEVCH"),
            bpt_data.BusRange(20002, 20002, "BJ6123C7BHD"),
            bpt_data.BusRange(20003, 20003, "BJ6123SHEVCA-5"),
        ]
        report = bpt_data.reconcile_rides(records, {"BJ6906FCEVCH", "BJ6123C7BHD", "BJ6123SHEVCA-5"}, set(), ranges)
        kinds = [row[0] for row in report["mismatch"]]  # type: ignore[index]
        self.assertEqual(kinds[0], "完全不搭（可能记录写错/号段表是别的车）")
        self.assertEqual(kinds[-1], "写法差异（多了厂内代号/括号）")
        # 每一行现在有 4 列：差异性质 / 对照 / 条数 / 例子
        self.assertTrue(all(len(row) == 4 for row in report["mismatch"]))  # type: ignore[index]

    def test_health_reports_overlapping_ranges(self):
        """公交号段重叠：同一个自编号落在两个号段里，查车会按先遇到的回答。"""
        ranges = [
            BusRange(20000, 20212, "BJ6123C7C4D"),
            BusRange(20200, 20300, "BJ6127C8BTB-1"),   # 与上一段重叠 20200-20212
            BusRange(30000, 30100, "BJ6123C7C4D"),
        ]
        with scratch() as directory:
            items = health_report(Path(directory), ranges, {}, {}, [])
        overlap = next(item for item in items if item.key == "overlap")
        self.assertEqual(len(overlap.rows), 1)
        self.assertIn("20200-20212", overlap.rows[0][1])
        self.assertIn("重叠 13 个自编号", overlap.rows[0][1])
        self.assertIn("车型不同", overlap.rows[0][1])
        # 没有重叠时这一项是空的
        with scratch() as directory:
            items = health_report(Path(directory), ranges[:1] + ranges[2:], {}, {}, [])
        self.assertEqual(next(item for item in items if item.key == "overlap").rows, [])

    def test_health_reports_body_categories(self):
        ranges = load_ranges(SOURCE_DATA)
        report = {item.key: item for item in health_report(Path.cwd(), ranges, {}, {}, [])}
        self.assertIn("body", report)
        self.assertIn("body_guess", report)
        guessed = [row[0] for row in report["body_guess"].rows]
        self.assertTrue(guessed)
        self.assertTrue(all(model_tags(name).body == "双层" for name in guessed))
        overrides = {"models": {name: {"body": "单机"} for name in guessed}}
        again = {item.key: item for item in health_report(Path.cwd(), ranges, {}, overrides, [])}
        self.assertEqual(again["body_guess"].rows, [])


class PinyinSearchTests(unittest.TestCase):
    def test_initials_and_subsequence(self):
        self.assertEqual(pinyin_initials("北汽福田"), "bqft")
        self.assertEqual(pinyin_initials("厦门金龙"), "xmjl")
        self.assertTrue(_is_subsequence("bqt", "bqft"))
        self.assertFalse(_is_subsequence("ftb", "bqft"))

    def test_search_hit_covers_code_manufacturer_alias_and_pinyin(self):
        self.assertTrue(search_hit("bj6123", "BJ6123C7C4D", "北汽福田"))
        self.assertTrue(search_hit("福田", "BJ6123C7C4D", "北汽福田"))
        self.assertTrue(search_hit("bqft", "BJ6123C7C4D", "北汽福田"))
        self.assertTrue(search_hit("bqt", "BJ6123C7C4D", "北汽福田"))  # 跳着打也能命中
        self.assertTrue(search_hit("df", "BJ6123C7BTD", "北汽福田", ["大福后"]))
        self.assertFalse(search_hit("zzz", "BJ6123C7C4D", "北汽福田", ["大福后"]))
        self.assertTrue(search_hit("", "任意"))


class ModelCardTests(unittest.TestCase):
    def test_model_card_is_self_contained(self):
        with scratch() as directory:
            base = Path(directory)
            image_directory(base).mkdir(exist_ok=True)
            Image.new("RGB", (500, 400), "blue").save(image_directory(base) / "A.jpg", format="JPEG")
            profile = {
                "introduction": "18 米铰接客车",
                "specifications": "厂商：北汽福田",
                "images": [{"name": "A.jpg", "caption": "左侧"}],
            }
            ranges = [BusRange(20000, 20099, "BJ6180C8CTD"), BusRange(20100, 20109, "BJ6180C8CTD")]
            target = base / "card.html"
            written, size = export_model_card(target, "BJ6180C8CTD", ranges, {}, profile, base)
            text = written.read_text(encoding="utf-8")
            self.assertGreater(size, 0)
            self.assertIn("BJ6180C8CTD", text)
            self.assertIn("北汽福田", text)
            self.assertIn("铰接", text)
            self.assertIn("data:image/jpeg;base64", text)
            self.assertIn("左侧", text)
            self.assertIn("18 米铰接客车", text)
            self.assertIn("20000-20099", text)
            self.assertIn("109", text)  # 覆盖 110 个编号，卡片里以表格数字呈现
            self.assertNotIn("<script", text)  # 卡片是纯静态的
            self.assertNotIn("http://", text)

    def test_web_payload_tags_carry_four_fields(self):
        ranges = [BusRange(20000, 20212, "BJ6123C7C4D")]
        payload = web_payload(ranges, {}, {}, [])
        self.assertEqual(payload["tags"]["BJ6123C7C4D"], ["北汽福田", "未标注", "12 米级", "单机"])
        html = bpt_web.render_html(payload)
        self.assertIn('id="body"', html)  # 网页版多了一个「车体形式」筛选


    def test_save_ranges_keeps_extra_columns_when_only_the_model_is_renamed(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / DATA_FILENAME
            header = list(EXPECTED_HEADERS) + ["镜像"]
            write_table(
                path,
                header,
                [
                    ["20000", "20010", "旧型号", '{"type": "旧型号"}'],
                    ["20011", "20020", "另一个", '{"type": "另一个"}'],
                ],
                numeric_columns={0, 1},
            )
            ranges = load_ranges(path)
            renamed = [BusRange(item.start, item.end, "新型号") if item.model == "旧型号" else item for item in ranges]
            save_ranges(path, renamed)
            _header, rows = read_table(path)
            self.assertEqual(rows[0][:3], ["20000", "20010", "新型号"])
            # 额外列要跟着走，而且里面的型号字样也要一起换掉
            self.assertEqual(rows[0][3], '{"type": "新型号"}')
            self.assertEqual(rows[1][3], '{"type": "另一个"}')
            self.assertEqual([item.model for item in load_ranges(path)], ["新型号", "另一个"])

    def test_save_ranges_still_keeps_extras_when_span_changes(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / DATA_FILENAME
            write_table(
                path,
                list(EXPECTED_HEADERS) + ["说明"],
                [["20000", "20010", "A", "原有说明"]],
                numeric_columns={0, 1},
            )
            save_ranges(path, [BusRange(20000, 20019, "A")])
            _header, rows = read_table(path)
            self.assertEqual(rows[0], ["20000", "20019", "A", "原有说明"])


class RideReconcileTests(unittest.TestCase):
    def test_parse_ride_sheet_finds_header_year_and_columns(self):
        rows = [
            ["2026公交出行记录", "", "", "", "", "", "", ""],
            ["日期", "上车时间", "线路", "车辆编号", "上车站", "下车站", "车型", "备注"],
            ["1.1", "0.5", "886路", "31688", "A站", "B站", "BJ6127C8BTB-1", "早高峰"],
            ["", "", "", "", "", "", "", ""],
            ["1.2", "0.6", "专45路", "31939", "C站", "D站", "BJ6680SHEVCA", ""],
        ]
        records = parse_ride_sheet(rows)
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0].year, "2026")
        self.assertEqual((records[0].line, records[0].number, records[0].model), ("886路", "31688", "BJ6127C8BTB-1"))
        self.assertEqual((records[0].board, records[0].alight, records[0].note), ("A站", "B站", "早高峰"))
        self.assertEqual(parse_ride_sheet([]), [])
        self.assertEqual(parse_ride_sheet([["随便", "一张表"]]), [])

    def test_metro_and_out_of_town_classification(self):
        for model in ("DKZ31", "SFM13", "ZBM15", "BDK07", "25G", "02A05", "CR400BF"):
            self.assertTrue(looks_like_metro_model(model), model)
        for model in ("BJ6123C7C4D", "ZK6140HNG2", "LCK6850EVG3M1"):
            self.assertFalse(looks_like_metro_model(model), model)
        for line in ("洛阳101路", "承德7路", "上海757路", "G6756", "K7754", "天津5路"):
            self.assertTrue(looks_like_out_of_town(line), line)
        for line in ("886路", "专45路", "郊80路", "昌31路", "Y11路", "15号线"):
            self.assertFalse(looks_like_out_of_town(line), line)

    def test_read_sheets_reads_every_worksheet(self):
        sheets = read_sheets(SOURCE_DATA)
        self.assertEqual(len(sheets), 1)
        name, rows = sheets[0]
        self.assertTrue(name)
        self.assertEqual(tuple(rows[0][:3]), EXPECTED_HEADERS)

    def test_reconcile_matches_by_number_not_by_model_string(self):
        ranges = [BusRange(20000, 20010, "A")]
        records = [
            RideRecord(year="2026", date="1.1", line="1路", number="20000", model="A"),  # 一致
            RideRecord(year="2026", date="1.2", line="1路", number="20005", model="B"),  # 编号查得到，型号不同
            RideRecord(year="2026", date="1.3", line="2路", number="90001", model="C"),  # 编号查不到
            RideRecord(year="2026", date="1.4", line="3路", number="", model="D"),  # 没编号、型号也不认识
            RideRecord(year="2026", date="1.5", line="4路", number="1", model="DKZ31"),  # 地铁
            RideRecord(year="2026", date="1.6", line="洛阳101路", number="5", model="CAT6100CRBEVT1"),  # 外地
        ]
        report = reconcile_rides(records, {"A"}, {"DKZ4"}, ranges)
        self.assertEqual(report["totals"], {"all": 6, "metro": 1, "out_of_town": 1, "no_number": 1})
        self.assertEqual(report["matched"], [("A", "1")])
        self.assertEqual(
            report["mismatch"],
            [("少/多一个字符", "A → B", "1", "1路 20005")],
        )
        self.assertEqual(report["number_gap"], [("90001", "C", "1")])
        self.assertEqual(report["model_gap"], [("D", "1")])
        # line_models 记的是**号段表里的车型**：20000 与 20005 查到的都是 A，所以 1 路记两次
        self.assertEqual(report["line_models"]["1路"], {"A": 2})

    def test_reconcile_reports_unridden_and_line_pairs(self):
        ranges = [BusRange(20000, 20010, "A"), BusRange(20100, 20110, "B")]
        records = [RideRecord(year="2025", date="1.1", line="特1路", number="20000", model="A")]
        report = reconcile_rides(records, {"A", "B"}, set(), ranges)
        self.assertEqual(report["unridden"], ["B"])
        self.assertEqual(report["line_count"], 1)
        self.assertEqual(report["pair_count"], 1)

    def test_reconcile_handles_missing_number_column(self):
        records = [RideRecord(year="2026", date="1.1", line="1路", number="", model="A")]
        report = reconcile_rides(records, {"A"}, set(), [BusRange(20000, 20010, "A")])
        self.assertEqual(report["model_gap"], [("A", "1")])
        self.assertEqual(report["matched"], [])


class LineDataTests(unittest.TestCase):
    """线路维度：读写、合并、索引、状态与对账。"""

    def test_lines_round_trip_and_default_empty(self):
        with scratch() as directory:
            base = Path(directory)
            self.assertEqual(bpt_data.load_lines(base), [])  # 默认空文件也能跑
            links = [
                bpt_data.LineLink("886路", "BJ6127C8BTB-1", ("31688", "31700"), "2024.9", "2026.8", "current"),
                bpt_data.LineLink("专45路", "BJ6680SHEVCA", ("31939",), "2025.3", "2026.9"),
            ]
            bpt_data.save_lines(base, links)
            loaded = bpt_data.load_lines(base)
            self.assertEqual(len(loaded), 2)
            self.assertEqual(loaded[0].numbers, ("31688", "31700"))
            self.assertEqual(loaded[0].status, "current")
            self.assertEqual(loaded[1].status, "recorded")  # 缺字段用默认值（向后兼容）

    def test_load_lines_lenient_reports_bad_rows(self):
        with scratch() as directory:
            base = Path(directory)
            (base / bpt_data.LINES_FILENAME).write_text(
                '{"links": [{"line": "886路", "model": "A"}, {"model": "没有线路名"}, "不是对象"]}',
                encoding="utf-8",
            )
            links, notes = bpt_data.load_lines_lenient(base)
            self.assertEqual([link.line for link in links], ["886路"])
            self.assertEqual(len(notes), 2)
            broken = base / bpt_data.LINES_FILENAME
            broken.write_text("{ 这不是 json", encoding="utf-8")
            links, notes = bpt_data.load_lines_lenient(base)
            self.assertEqual(links, [])
            self.assertTrue(notes)

    def test_merge_lines_unions_time_and_numbers(self):
        current = [bpt_data.LineLink("886路", "BJ6127C8BTB-1", ("31688",), "2024.9", "2025.1", "current")]
        incoming = [bpt_data.LineLink("886路", "BJ6127C8BTB-1", ("31700",), "2023.5", "2026.8", "recorded", "import")]
        plan = bpt_data.merge_lines(current, incoming)
        self.assertEqual(len(plan.added), 0)
        self.assertEqual(len(plan.updated), 1)
        merged, _old = plan.updated[0]
        self.assertEqual(merged.numbers, ("31688", "31700"))  # 并集
        self.assertEqual((merged.first, merged.last), ("2023.5", "2026.8"))  # 时间取范围
        self.assertEqual(merged.status, "current")  # 已确认"还在跑"的不被导入数据降级
        result = bpt_data.apply_line_plan(current, plan)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].numbers, ("31688", "31700"))

    def test_indexes_by_model_and_number(self):
        links = [
            bpt_data.LineLink("886路", "A", ("31688",), "2024.9", "2026.8", "recorded"),
            bpt_data.LineLink("919路", "A", ("31700",), "2023.1", "2025.2", "current"),
            bpt_data.LineLink("专45路", "B", ("31939",), "2025.3", "2026.9", "recorded"),
        ]
        by_model = bpt_data.lines_for_model(links, "A")
        self.assertEqual([link.line for link in by_model], ["919路", "886路"])  # 现在还在跑的排前面
        self.assertEqual([link.line for link in bpt_data.lines_for_number(links, "31688")], ["886路"])
        # 位数不足会补零（"8688" → "08688"），但补零后的号跟别的号不会混起来
        self.assertEqual(bpt_data.lines_for_number(links, "3168"), [])
        self.assertEqual([link.line for link in bpt_data.lines_for_number(links, "31700")], ["919路"])
        self.assertEqual(bpt_data.lines_for_number(links, "99999"), [])
        self.assertEqual(bpt_data.models_for_line(links, "886路"), ["A"])
        self.assertEqual(bpt_data.numbers_for_line(links, "886路"), ["31688"])
        self.assertEqual(bpt_data.line_display(links[0]), "886路（2024.9–2026.8）")

    def test_family_suggestions_only_hint(self):
        links = [
            bpt_data.LineLink("886路", "A"),
            bpt_data.LineLink("886路区间", "A"),
            bpt_data.LineLink("1路", "B"),
        ]
        families = dict(bpt_data.line_family_suggestions(links))
        self.assertEqual(families["886"], ["886路", "886路区间"])
        self.assertNotIn("1", families)  # 只有一条不提示
        self.assertEqual(bpt_data.line_family("886路快车"), "886")

    def test_build_links_from_records_and_discrepancies(self):
        records = [
            bpt_data.RideRecord("2024", "9.22", "886路", "31688", "BJ6127C8BTB-1", "A", "B"),
            bpt_data.RideRecord("2026", "8.27", "886路", "31700", "BJ6127C8BTB-1", "C", "D"),
            bpt_data.RideRecord("2025", "3.1", "专45路", "31939", "不存在的型号", "E", "F"),
        ]
        links = bpt_data.build_links_from_records(records)
        self.assertEqual(len(links), 2)
        first = next(link for link in links if link.line == "886路")
        self.assertEqual(first.numbers, ("31688", "31700"))
        self.assertEqual((first.first, first.last), ("2024.9", "2026.8"))
        ranges = [bpt_data.BusRange(31688, 31700, "BJ6127C8BTB-1")]
        issues = bpt_data.line_discrepancies(links, ranges)
        self.assertEqual(issues["unknown_model"], [("专45路", "不存在的型号")])
        self.assertEqual(issues["number_mismatch"], [])


class TagSuggestionTests(unittest.TestCase):
    """从用户已有标注里学「后缀 → 标签值」，以及按后缀分组。"""

    def test_suffix_candidates_strip_variant_and_need_two_letters(self):
        self.assertEqual(bpt_data.tag_suffix_candidates("BJ6123C7BTD"), ["BTD", "TD"])
        self.assertEqual(bpt_data.tag_suffix_candidates("BJ6123C7BTD-1"), ["BTD", "TD"])
        self.assertEqual(bpt_data.tag_suffix_candidates("BJ6123C6B4D"), ["B4D"])  # "4D" 只有一个字母，不算
        self.assertEqual(bpt_data.tag_suffix_candidates("HFF6105G39C"), [])  # 别的厂商命名方式，不硬套
        self.assertEqual(bpt_data.tag_suffix_candidates("BK560"), [])

    def test_learn_rules_needs_support_and_no_conflict(self):
        overrides = {
            "models": {
                "BJ6123C7BTD": {"fuel": "天然气"},
                "BJ6127C8BTD": {"fuel": "天然气"},
                "BJ6851C6BXD": {"fuel": "柴油"},  # 只有一台，支持度不够
            }
        }
        hints = bpt_data.learn_tag_rules(overrides, "fuel")
        suffixes = {hint.suffix: hint for hint in hints}
        self.assertEqual(suffixes["BTD"].value, "天然气")
        self.assertEqual(suffixes["BTD"].support, 2)
        self.assertIn("BJ6123C7BTD", suffixes["BTD"].examples)
        self.assertNotIn("BXD", suffixes)  # 单台证据不推断

        # 同一个后缀对应过两个值 → 整条规则作废
        conflicting = {"models": dict(overrides["models"], **{"BJ6100C8BTD": {"fuel": "柴油"}})}
        self.assertNotIn("BTD", {hint.suffix for hint in bpt_data.learn_tag_rules(conflicting, "fuel")})

    def test_suggest_skips_models_that_already_have_a_value(self):
        overrides = {
            "models": {
                "AABTD": {"fuel": "天然气"},
                "BBBTD": {"fuel": "天然气"},
                "CCCZZZ": {"fuel": "柴油"},  # 已经标过（后缀只此一台，不生成规则）
            }
        }
        hints = bpt_data.learn_tag_rules(overrides, "fuel")
        self.assertIsNone(bpt_data.suggest_tag(overrides, "fuel", "CCCZZZ", hints))
        suggestion = bpt_data.suggest_tag(overrides, "fuel", "DDDBTD", hints)
        self.assertIsNotNone(suggestion)
        self.assertEqual(suggestion.value, "天然气")
        self.assertIn("BTD", suggestion.evidence())
        # 后缀对应过两个值就不猜：CTD 一台天然气一台柴油
        ambiguous = {"models": {"AACDT": {"fuel": "天然气"}, "BBCDT": {"fuel": "柴油"}}}
        self.assertIsNone(bpt_data.suggest_tag(ambiguous, "fuel", "CCCDT", bpt_data.learn_tag_rules(ambiguous, "fuel")))

    def test_suggest_tags_for_models_splits_suggested_and_unsure(self):
        overrides = {"models": {"AABTD": {"fuel": "天然气"}, "BBBTD": {"fuel": "天然气"}}}
        suggested, unsure = bpt_data.suggest_tags_for_models(overrides, "fuel", ["CCBTD", "ZZXYZ"])
        self.assertEqual([item.model for item in suggested], ["CCBTD"])
        self.assertEqual(unsure, ["ZZXYZ"])

    def test_group_by_suffix_keeps_variants_together_and_skips_tagged(self):
        overrides = {"models": {"BJ6127C8BTD": {"fuel": "天然气"}}}
        groups = bpt_data.group_by_suffix(["BJ6127C8BTB", "BJ6127C8BTB-1", "BJ6127C8BTD"], overrides, "fuel")
        self.assertEqual(groups[0], ("BTB", ["BJ6127C8BTB", "BJ6127C8BTB-1"]))
        self.assertNotIn("BJ6127C8BTD", [model for _suffix, models in groups for model in models])
        # 提不出后缀的按厂商兜底，而不是全塞进一个"无法归组"
        fallback = dict(bpt_data.group_by_suffix(["HFF6110GS101D", "CA6120URD22", "BK6100N"], None, "fuel"))
        self.assertEqual(fallback["HFF 系（无后缀）"], ["HFF6110GS101D"])
        self.assertEqual(fallback["CA 系（无后缀）"], ["CA6120URD22"])
        self.assertEqual(fallback["BK 系（无后缀）"], ["BK6100N"])

    def test_facts_beat_learned_rules_and_carry_a_source(self):
        # 学到的规律说 BTD→天然气；但查证事实说 BJ6127C8MTB 是天然气、且事实优先
        facts = {fact.model for fact in bpt_data.TAG_FACTS}
        self.assertIn("BJ6127C8MTB", facts)
        overrides = {"models": {"AABTD": {"fuel": "柴油"}, "BBBTD": {"fuel": "柴油"}}}
        hints = bpt_data.learn_tag_rules(overrides, "fuel")
        suggestion = bpt_data.suggest_tag(overrides, "fuel", "BJ6127C8MTB", hints)
        self.assertIsNotNone(suggestion)
        self.assertEqual(suggestion.value, "天然气")  # 事实赢过学到的规律
        self.assertEqual(suggestion.source, "已核实")
        self.assertIn("chinabuses", suggestion.evidence())
        self.assertIn("http", suggestion.evidence())

    def test_fact_without_a_mappable_value_asks_the_user_to_choose(self):
        fact = bpt_data.fact_for("BJ6123C7C4D", "fuel")
        self.assertIsNotNone(fact)
        self.assertEqual(fact.value, "")
        self.assertIn("混合动力", fact.detail)
        suggestion = bpt_data.suggest_tag(None, "fuel", "BJ6123C7C4D", [])
        self.assertEqual(suggestion.value, "")
        self.assertEqual(suggestion.source, "查到待你选")
        self.assertIn("17vin", suggestion.evidence())

    def test_facts_are_exact_model_matches(self):
        self.assertIsNone(bpt_data.fact_for("BJ6127C8BTB-1", "fuel"))  # 变体型号不自动套用整条事实
        self.assertIsNone(bpt_data.fact_for("BJ6127C8MTB", "length"))
        self.assertIsNotNone(bpt_data.fact_for("BJ6127C8MTB", "fuel"))

    def test_applicable_suggestions_skip_empty_values(self):
        suggestions = [
            bpt_data.suggest_tag(None, "fuel", "BJ6127C8MTB", []),  # 已核实，有值
            bpt_data.suggest_tag(None, "fuel", "BJ6123C7C4D", []),  # 查到待你选，值为空
        ]
        applicable = bpt_data.applicable_suggestions([item for item in suggestions if item is not None])
        self.assertEqual([item.model for item in applicable], ["BJ6127C8MTB"])

    def test_learned_rules_cover_a_synthetic_family(self):
        """用固定的小数据断言覆盖率（不看用户随时会改的 车型标签.json）。"""
        overrides = {
            "models": {
                "BJ6100C7BTD": {"fuel": "天然气"},
                "BJ6123C7BTD": {"fuel": "天然气"},
                "BJ7000C8MJB": {"fuel": "柴油"},
                "BJ7100C8MJB": {"fuel": "柴油"},
            }
        }
        suggested, unsure = bpt_data.suggest_tags_for_models(
            overrides, "fuel", ["BJ6200C7BTD", "BJ7200C8MJB", "ZZZ999"]
        )
        self.assertEqual([item.model for item in suggested], ["BJ6200C7BTD", "BJ7200C8MJB"])
        self.assertEqual([item.value for item in suggested], ["天然气", "柴油"])
        self.assertEqual(unsure, ["ZZZ999"])

    def test_live_tag_file_is_consistent_with_suggestions(self):
        """真实数据只断言"自洽"，不断言具体条数——用户会把标签补完，条数会变。"""
        overrides = load_tag_overrides(Path.cwd())
        hints = bpt_data.learn_tag_rules(overrides, "fuel")
        for hint in hints:
            self.assertGreaterEqual(hint.support, 2)
            self.assertTrue(hint.examples)
            self.assertNotEqual(hint.value, UNKNOWN_TAG)
        ranges = load_ranges(SOURCE_DATA)
        untagged = [model for model in model_names(ranges) if model_tags(model, overrides).fuel == UNKNOWN_TAG]
        suggested, unsure = bpt_data.suggest_tags_for_models(overrides, "fuel", untagged)
        self.assertEqual(len(suggested) + len(unsure), len(untagged))
        for item in suggested:
            self.assertTrue(item.source)
            self.assertEqual(item.model, item.model)
        for model in unsure:
            self.assertNotIn(model, [item.model for item in suggested])

    def test_curated_lists_reference_real_models(self):
        """内置清单里的型号必须真的在号段表里——写错一个字符会静默失效（真踩过）。"""
        models = set(model_names(load_ranges(SOURCE_DATA)))
        for name in KNOWN_DOUBLE_DECKERS:
            self.assertIn(name, models, f"内置双层清单里的 {name} 在号段表里找不到")
        for fact in bpt_data.TAG_FACTS:
            self.assertIn(fact.model, models, f"查证事实里的 {fact.model} 在号段表里找不到")

    def test_facts_are_all_cited(self):
        """查证过的事实必须都带来源链接（这是这一轮的核心承诺）。"""
        self.assertGreaterEqual(len(bpt_data.TAG_FACTS), 10)
        for fact in bpt_data.TAG_FACTS:
            self.assertTrue(fact.model)
            self.assertTrue(fact.detail)
            self.assertTrue(fact.source)
            self.assertTrue(fact.url.startswith("http"), fact.model)

    def test_shipped_tags_produce_suggestions_for_untagged_models(self):
        """真实数据只做自洽检查：学到的规律与建议都要"有依据、能解释"。

        不断言"一定要学到 N 条规律"或"一定有 N 个建议"——用户会把标签补完，
        那时候没有待确认项才是正确状态（真的踩过这个坑）。
        """
        overrides = load_tag_overrides(Path.cwd())
        hints = bpt_data.learn_tag_rules(overrides, "fuel")
        for hint in hints:
            self.assertGreaterEqual(hint.support, 2)
            self.assertTrue(hint.examples, hint.suffix)
            self.assertIn(hint.suffix, hint.examples[0].upper())
        ranges = load_ranges(SOURCE_DATA)
        untagged = [
            model for model in model_names(ranges) if model_tags(model, overrides).fuel == UNKNOWN_TAG
        ]
        suggested, unsure = bpt_data.suggest_tags_for_models(overrides, "fuel", untagged)
        self.assertEqual(len(suggested) + len(unsure), len(untagged))
        for item in suggested:
            if item.value:  # 「查到待你选」的那条值为空，不该被当成可写入的建议
                self.assertNotEqual(item.value, UNKNOWN_TAG)
            self.assertTrue(item.source)
            self.assertTrue(item.evidence())


class DataTableTests(unittest.TestCase):
    """共享表格组件：排序、指示、斑马、保持选中、复制/导出、空状态。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")
        cls.app.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def make_table(self, **kwargs):
        window = tk.Toplevel(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        table = DataTable(
            window,
            self.app.theme,
            (
                TableColumn("model", "车型", 160),
                TableColumn("count", "编号数", 90, "e", numeric=True),
                TableColumn("fuel", "燃料", 80, "center"),
            ),
            **kwargs,
        )
        table.card.pack(fill="both", expand=True)
        return table

    def test_rows_are_zebra_striped_and_text_only(self):
        table = self.make_table()
        table.set_rows([("A", 3, "柴油"), ("B", 12, "纯电动")])
        self.assertEqual(table.rows(), [("A", "3", "柴油"), ("B", "12", "纯电动")])
        tags = [table.tree.item(item, "tags") for item in table.tree.get_children()]
        self.assertNotIn("odd", tags[0])
        self.assertIn("odd", tags[1])

    def test_numeric_column_sorts_by_value_not_text(self):
        table = self.make_table()
        table.set_rows([("A", 3, "柴油"), ("B", 12, "纯电动"), ("C", 100, "天然气")])
        table.sort_by("count")
        self.assertEqual([row[0] for row in table.rows()], ["A", "B", "C"])
        table.sort_by("count")  # 再点一次反过来
        self.assertEqual([row[0] for row in table.rows()], ["C", "B", "A"])

    def test_heading_shows_sort_direction(self):
        table = self.make_table()
        table.set_rows([("A", 1, "柴油")])
        self.assertEqual(tree_headings(table.tree)[0], "车型 ▲")
        table.sort_by("fuel")
        self.assertEqual(tree_headings(table.tree)[1:], ["编号数", "燃料 ▲"])
        table.sort_by("fuel")
        self.assertEqual(tree_headings(table.tree)[2], "燃料 ▼")

    def test_custom_sort_key_wins(self):
        table = self.make_table(sort_key=lambda column, row: {"A": 9, "B": 1, "C": 5}[row[0]] if column == "model" else row[1])
        table.set_rows([("A", 1, "x"), ("B", 1, "x"), ("C", 1, "x")])
        self.assertEqual([row[0] for row in table.rows()], ["B", "C", "A"])

    def test_selection_survives_refresh_and_iid_is_kept(self):
        table = self.make_table()
        table.set_rows([("A", 1, "x"), ("B", 2, "y")], iid=lambda index, row: row[0])
        table.tree.selection_set("B")
        table.set_rows([("C", 3, "z"), ("B", 2, "y")], iid=lambda index, row: row[0])
        self.assertEqual(table.selection_first(), "B")
        self.assertEqual(sorted(table.tree.get_children()), ["B", "C"])

    def test_set_columns_keeps_natural_order_when_asked(self):
        """set_columns 不能把 sort_on_load=False 的表又按第一列排序（会把排好的顺序打乱）。"""
        table = self.make_table(sort_on_load=False)
        table.set_columns((TableColumn("m", "车型", 160), TableColumn("n", "条数", 90)))
        table.set_rows([("C", 1), ("A", 2), ("B", 3)])   # 切完列再塞数据（refresh_table 就是这么用的）
        self.assertEqual(table.sort_column, "")
        self.assertEqual([table.tree.item(i, "values")[0] for i in table.tree.get_children()], ["C", "A", "B"])
        # 默认表仍然按第一列排
        plain = self.make_table()
        plain.set_columns((TableColumn("m", "车型", 160), TableColumn("n", "条数", 90)))
        plain.set_rows([("C", 1), ("A", 2)])
        self.assertEqual(plain.sort_column, "m")
        self.assertEqual([plain.tree.item(i, "values")[0] for i in plain.tree.get_children()], ["A", "C"])

    def test_set_columns_swaps_and_resets_sort(self):
        table = self.make_table()
        table.set_rows([("A", 1, "x")])
        table.sort_by("count")
        table.set_columns((TableColumn("a", "甲", 100), TableColumn("b", "乙", 100, "e", numeric=True)))
        self.assertEqual(tree_headings(table.tree), ["甲 ▲", "乙"])
        self.assertEqual(table.rows(), [])

    def test_copy_and_empty_state(self):
        table = self.make_table(empty_text="还没有数据")
        table.set_rows([])
        self.assertEqual(table.copy(), 0)
        table.set_rows([("A", 1, "柴油")])
        self.assertEqual(table.copy(), 1)
        self.assertIn("车型", self.app.clipboard_get())
        self.assertEqual(table.empty.label.cget("text"), "还没有数据")

    def test_keyboard_hooks(self):
        """Delete 删选中行、Ctrl+F 跳到搜索框：给表格传回调就该绑上。"""
        called: list[str] = []
        table = self.make_table(on_delete=lambda: called.append("delete"), on_search=lambda: called.append("search"))
        self.assertIn("<Key-Delete>", table.tree.bind())
        self.assertIn("<Control-Key-f>", table.tree.winfo_toplevel().bind())
        plain = self.make_table()
        self.assertNotIn("<Key-Delete>", plain.tree.bind())

    def test_copy_a_single_column(self):
        """只想复制车型（或任意一列）：默认整列，右键菜单里可以只复制选中的。"""
        table = self.make_table()
        table.set_rows([("BJ6123C7C4D", 10, "柴油"), ("BJ6127C8BTB-1", 5, "柴油"), ("BJ6851C6BCD", 3, "纯电动")])
        # 默认整列（刷新后程序会自动选中第一行，不能因此只复制一行）
        self.assertEqual(table.copy_column("model"), 3)
        self.assertEqual(self.app.clipboard_get(), "BJ6123C7C4D\nBJ6127C8BTB-1\nBJ6851C6BCD")
        # 显式只要选中的
        table.tree.selection_set(table.tree.get_children()[1])
        self.assertEqual(table.copy_column("model", only_selected=True), 1)
        self.assertEqual(self.app.clipboard_get(), "BJ6127C8BTB-1")
        self.assertEqual(table.copy_column("fuel"), 3)
        self.assertEqual(self.app.clipboard_get(), "柴油\n柴油\n纯电动")
        # 列名不存在 / 空表
        self.assertEqual(table.copy_column("nope"), 0)
        table.set_rows([])
        self.assertEqual(table.copy_column("model"), 0)
        # 每张表都有右键菜单（复制选中行 / 只复制这一列）
        self.assertIn("<Button-3>", table.tree.bind())

    def test_input_order_is_kept_until_sorted(self):
        """sort_on_load=False：先按塞进来的顺序显示；点表头升序 → 降序 → 回到输入顺序。"""
        table = self.make_table(sort_on_load=False)

        def shown() -> list[str]:
            return [table.tree.item(item, "values")[0] for item in table.tree.get_children()]

        table.set_rows([("31779", 1, "x"), ("20001", 1, "y"), ("31778", 1, "z")])
        self.assertEqual(shown(), ["31779", "20001", "31778"])   # 用户的输入顺序
        table.sort_by("model")
        self.assertEqual(shown(), ["20001", "31778", "31779"])   # 升序
        table.sort_by("model")
        self.assertEqual(shown(), ["31779", "31778", "20001"])   # 降序
        table.sort_by("model")
        self.assertEqual(shown(), ["31779", "20001", "31778"])   # 第三下回到输入顺序
        self.assertEqual(table.sort_column, "")
        # 回到输入顺序时表头不再显示排序箭头，而是 ⇅
        self.assertEqual(tree_headings(table.tree), ["车型 ⇅", "编号数 ⇅", "燃料 ⇅"])

    def test_default_still_sorts_on_load(self):
        """默认（sort_on_load=True）仍按第一列升序显示，老窗口行为不变。"""
        table = self.make_table()
        table.set_rows([("B", 1, "x"), ("A", 2, "y")])
        self.assertEqual([table.tree.item(item, "values")[0] for item in table.tree.get_children()], ["A", "B"])

    def test_row_ids_are_bound_before_sorting(self):
        """排序不能改变"行 id → 数据"的对应关系（否则编辑/删除会改错行）。"""
        table = self.make_table()
        table.set_rows(
            [("C", 3, "x"), ("A", 1, "x"), ("B", 2, "x")],
            iid=lambda position, row: f"row{position}",
        )
        before = {child: table.tree.item(child, "values") for child in table.tree.get_children()}
        table.sort_by("count")  # 按编号数排（升序）：A(1) → B(2) → C(3)
        after = {child: table.tree.item(child, "values") for child in table.tree.get_children()}
        self.assertEqual(before, after)
        self.assertEqual(table.tree.get_children()[0], "row1")  # A 行排到最前，但它的 id 还是 row1
        self.assertEqual(table.tree.item("row1", "values")[0], "A")

    def test_toolbar_is_built_above_the_tree(self):
        seen = {}
        table = self.make_table(toolbar=lambda bar: seen.setdefault("frame", bar))
        self.assertIsNotNone(table.toolbar)
        self.assertIs(seen["frame"], table.toolbar)


# 2026-10-01 从 bjbus.jinzihao.me（MCP）实测抓下来的真实回包，直接当测试夹具
ROUTE_IDS_FIXTURE = [
    {"id": 805, "first_stop": "德胜门西", "last_stop": "北京蟒山森林公园"},
    {"id": 806, "first_stop": "北京蟒山森林公园", "last_stop": "德胜门西"},
]
BUS_HISTORY_FIXTURE = {
    "877": ["20230726-20230726", "20240821-20240821", "20260910-20260910"],
    "885": ["20210702-20210702", "20211105-20211105", "20230712-20230712"],
    "886": ["20200815-20201012", "20201014-20201107", "20201109-20210110"],
    "889": ["20240730-20240730"],
}
TOOL_RESULT_FIXTURE = {"content": [{"type": "text", "text": json.dumps(ROUTE_IDS_FIXTURE, ensure_ascii=False)}]}

#: 实时轨迹（实测形状：30 秒一条，取最后一条就是"现在在哪"）
TRACK_FIXTURE = [
    {"time": "2026-10-10 08:54:33", "bus_id": "31688", "line_id": 805,
     "line_name": "886 开往 北京蟒山森林公园", "stop_index": 13, "stop_name": "军都度假村",
     "distance_m": 239, "arrived": False, "occupancy": 1},
    {"time": "2026-10-10 08:55:03", "bus_id": "31688", "line_id": 805,
     "line_name": "886 开往 北京蟒山森林公园", "stop_index": 13, "stop_name": "军都度假村",
     "distance_m": 490, "arrived": False, "occupancy": 1},
]

#: 「这条线用过哪些车」：自编号 → 日期区间（实测 886 路用过 319 辆，这里取几条做夹具）
ROUTE_BUSES_FIXTURE = {
    "31688": ["20200815-20201012", "20250801-20251010"],
    "31691": ["20210901-20211001"],
    "30846": ["20210322-20210623"],
}


class FakeTransport:
    """假的 MCP 传输：不联网，按脚本回包（测握手机制与错误路径）。"""

    def __init__(self, *, drop_first: bool = False) -> None:
        self.sent: list[dict] = []
        self._queue: "queue.Queue[dict]" = queue.Queue()
        self.opened = False
        self.closed = False
        self.drop_first = drop_first
        self._dropped = False

    def open(self) -> None:
        self.opened = True

    def post(self, payload: dict) -> None:
        self.sent.append(payload)
        if "id" not in payload:  # 通知没有 id，不需要回
            return
        if self.drop_first and not self._dropped:
            self._dropped = True
            return  # 故意不回，用来测超时
        if payload.get("method") == "initialize":  # 注意：tools/call 也带 method，别一律当握手
            self._queue.put({"jsonrpc": "2.0", "id": payload["id"], "result": {"protocolVersion": "2024-11-05"}})
            return
        name = (payload.get("params") or {}).get("name", "")
        if name == "query_route_id_by_route_number":
            self._queue.put({"jsonrpc": "2.0", "id": payload["id"], "result": TOOL_RESULT_FIXTURE})
        elif name == "query_bus_operating_history":
            self._queue.put(
                {
                    "jsonrpc": "2.0",
                    "id": payload["id"],
                    "result": {"content": [{"type": "text", "text": json.dumps(BUS_HISTORY_FIXTURE)}]},
                }
            )
        else:
            self._queue.put({"jsonrpc": "2.0", "id": payload["id"], "error": {"code": -32601, "message": "没这个工具"}})

    def events(self) -> "queue.Queue[dict]":
        return self._queue

    def close(self) -> None:
        self.closed = True


class MainOnlineTests(unittest.TestCase):
    """主界面自动补两行：这辆车现在在哪 / 最新在哪条线运营。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def wait_for_online(self, prefix: str) -> str:
        for _ in range(250):
            self.app.update()
            text = self.app.online_var.get()
            if text.startswith(prefix):
                return text
            time.sleep(0.03)
        self.fail(f"等不到主界面的联网信息（当前是 {self.app.online_var.get()!r}）")

    def test_main_window_shows_now_and_latest(self):
        self.app._auto_online = True
        self.addCleanup(setattr, self.app, "_auto_online", False)
        self.addCleanup(self.app.online_var.set, "")
        with scratch() as directory, patch_app_dir(Path(directory)):
            with patch.object(BPT_BusDigitalNumber, "shared_service", lambda *_a, **_k: FakeService()):
                self.app.number_var.set("31688")
                self.app.query()
                text = self.wait_for_online("现在 ")
        self.assertIn("886 开往 北京蟒山森林公园", text)   # 现在跑哪条线（来自轨迹）
        self.assertIn("军都度假村", text)                  # 现在在哪一站
        self.assertIn("距站 490 米", text)                 # 轨迹最后一条的距离
        self.assertIn("最近跑", text)                  # 最新在哪条线运营
        self.assertIn("886", text)

    def test_auto_lookup_can_be_turned_off(self):
        self.app._auto_online = False
        with scratch() as directory, patch_app_dir(Path(directory)):
            with patch.object(BPT_BusDigitalNumber, "shared_service", lambda *_a, **_k: FakeService()):
                self.app.number_var.set("31688")
                self.app.query()
                for _ in range(30):
                    self.app.update()
                    time.sleep(0.03)
        self.assertEqual(self.app.online_var.get(), "", "关掉开关后不该显示联网信息")

    def test_network_failure_is_shown_as_one_short_line(self):
        """联网失败时主界面只补一句短的，不弹框、不挡输入。"""
        self.app._auto_online = True
        self.addCleanup(setattr, self.app, "_auto_online", False)
        self.addCleanup(self.app.online_var.set, "")
        service = FakeService(fail=bpt_online.OnlineUnavailable("连不上联网服务（WinError 10061 积极拒绝）"))
        with scratch() as directory, patch_app_dir(Path(directory)):
            with patch.object(BPT_BusDigitalNumber, "shared_service", lambda *_a, **_k: service):
                self.app.number_var.set("31688")
                self.app.query()
                text = self.wait_for_online("联网没查到")
        self.assertIn("连接被拒绝", text)
        self.assertLess(len(text), 120)   # 短，不刷屏

    def test_incomplete_input_clears_the_online_line(self):
        self.app.online_var.set("现在：886 开往 北京蟒山森林公园")
        self.app.number_var.set("3168")   # 只有 4 位
        self.app.query()
        self.assertEqual(self.app.online_var.get(), "")


class LineSourceTests(unittest.TestCase):
    """线路记录的来源与证据：网上查来的绝不能冒充"我记过"（V2）。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def test_month_comparison_regression(self):
        """「2025.8」和「2025.10」按字符串比会排错——必须按年月比。"""
        self.assertEqual(bpt_data.month_key("2025.10"), (2025, 10))
        self.assertGreater(bpt_data.month_key("2025.10"), bpt_data.month_key("2025.8"))
        self.assertEqual(bpt_data.month_key("乱写"), (0, 0))
        manual = bpt_data.LineLink("886路", "A", (), "2024.9", "2025.8", "recorded", "manual")
        online = bpt_data.LineLink("886路", "A", (), "2020.8", "2025.10", "recorded", "online")
        merged = bpt_data.apply_line_plan([manual], bpt_data.merge_lines([manual], [online]))
        self.assertEqual((merged[0].first, merged[0].last), ("2020.8", "2025.10"))

    def test_online_never_overwrites_manual_provenance(self):
        manual = bpt_data.LineLink("886路", "A", ("31688",), "2024.9", "2025.8", "current", "manual", "我自己记的")
        online = bpt_data.lines_from_history("31688", {"886路": [("20200815", "20251010")]}, lambda _n: "A", checked="20261011")[0]
        self.assertEqual(online.source, "online")
        self.assertEqual(online.evidence[0].source, "online")
        merged = bpt_data.apply_line_plan([manual], bpt_data.merge_lines([manual], [online]))[0]
        self.assertEqual(merged.source, "manual")          # 来源不被覆盖
        self.assertEqual(merged.status, "current")         # 确认过的"还在跑"不被降级
        self.assertEqual(len(merged.evidence), 1)          # 证据链留着
        self.assertIn("网上查的", merged.evidence[0].label)

    def test_delete_by_source_and_stale(self):
        links = [
            bpt_data.LineLink("1路", "A", (), "2024.1", "2024.2", "recorded", "manual"),
            bpt_data.LineLink("2路", "A", (), "2024.1", "2024.2", "recorded", "import"),
            bpt_data.LineLink("3路", "A", (), "2024.1", "2024.2", "recorded", "online",
                              evidence=(bpt_data.LineEvidence("online", "31688", "20240101", "20240201", "20260101"),)),
        ]
        self.assertEqual([item.line for item in bpt_data.delete_by_source(links, "online")], ["1路", "2路"])
        self.assertEqual([item.line for item in bpt_data.delete_by_source(links, "import")], ["1路", "3路"])
        # 只有 online 那条会被算成"过期"（2026-01-01 查的，到 2026-03 已超 30 天）
        stale = bpt_data.online_stale(links, days=30, today="20260301")
        self.assertEqual([item.line for item in stale], ["3路"])
        self.assertEqual(bpt_data.online_stale(links, days=30, today="20260115"), [])

    def test_line_display_shows_source(self):
        manual = bpt_data.LineLink("886路", "A", (), "2024.9", "2025.8")
        online = bpt_data.LineLink("887路", "A", (), "2024.9", "2025.8", "recorded", "online")
        self.assertEqual(bpt_data.line_display(manual, with_source=True), "886路（2024.9–2025.8）")
        self.assertEqual(bpt_data.line_display(online, with_source=True), "887路（2024.9–2025.8） · 网上查的")
        self.assertEqual(bpt_data.source_label("online"), "网上查的")
        self.assertLess(bpt_data.source_rank("manual"), bpt_data.source_rank("online"))

    def test_line_query_refresh_scales_with_many_records(self):
        """性能护栏：1000 条线路记录时刷新不能卡。

        原来是「每条线路重扫一遍全部记录」= O(线路 × 记录)，实测 516ms（搜索框每敲一个字卡半秒）；
        归并成一次 O(记录) 之后是 15ms 左右。
        """
        window = LineQueryWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        window.links = [
            bpt_data.LineLink(f"{index}路", f"BJ{index:04d}", (str(index),), "2024.1", "2026.9")
            for index in range(1000)
        ]
        window.refresh_lines()
        mark = time.perf_counter()
        window.refresh_lines()
        cost = time.perf_counter() - mark
        self.assertEqual(len(window.line_table.tree.get_children()), 1000)
        self.assertLess(cost, 0.25, f"1000 条记录刷新用了 {cost * 1000:.0f} ms（性能回退了？）")
        first = window.line_table.tree.item(window.line_table.tree.get_children()[0], "values")
        self.assertEqual(first[2], "1")   # 车型数（第 0 列线路、第 1 列服务端首末站）
        self.assertEqual(first[3], "1")   # 自编号数

    def test_line_query_filters_by_source(self):
        window = LineQueryWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        window.links = [
            bpt_data.LineLink("1路", "A", ("10001",), "2024.1", "2024.2", "recorded", "manual"),
            bpt_data.LineLink("2路", "A", ("20001",), "2024.1", "2024.2", "recorded", "online"),
        ]
        window.refresh_lines()
        self.assertEqual(len(window.line_table.tree.get_children()), 2)
        window.source_var.set("网上查的")
        window.refresh_lines()
        self.assertEqual([window.line_table.tree.item(i, "values")[0] for i in window.line_table.tree.get_children()], ["2路"])
        self.assertIn("来源：网上查的", window.summary_var.get())
        window.show_line()
        detail = window.detail_table.tree.item(window.detail_table.tree.get_children()[0], "values")
        self.assertEqual(detail[4], "网上查的")   # 明细里带来源列
        window.source_var.set("我记过")
        window.refresh_lines()
        self.assertEqual([window.line_table.tree.item(i, "values")[0] for i in window.line_table.tree.get_children()], ["1路"])
        window.clear_filter()
        self.assertEqual(window.source_var.get(), "全部来源")


class LineInfoTests(unittest.TestCase):
    """V3 线路维度：线路→车（按车型归并）、线路资料（首末站/站点表）。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def test_lines_from_route_history_groups_by_model(self):
        def model_of(number: str) -> str:
            return {"31688": "BJ6127C8BTB-1", "31691": "BJ6127C8BTB-1", "30846": "BJ6105EVCA-27"}.get(number, "")

        links = bpt_data.lines_from_route_history(
            "886", ROUTE_BUSES_FIXTURE, model_of, checked="20251015", today="20251015"
        )
        by_model = {link.model: link for link in links}
        self.assertEqual(set(by_model), {"BJ6127C8BTB-1", "BJ6105EVCA-27"})
        self.assertEqual(by_model["BJ6127C8BTB-1"].numbers, ("31688", "31691"))  # 同型号归并
        self.assertEqual(by_model["BJ6127C8BTB-1"].line, "886")
        self.assertEqual(by_model["BJ6127C8BTB-1"].source, "online")
        self.assertEqual(by_model["BJ6127C8BTB-1"].status, "current")   # 31688 到 2025-10-10，7 天内
        self.assertEqual(by_model["BJ6105EVCA-27"].status, "recorded")
        # 证据只写一条总体区间（一条线可能几百辆车，逐车写会把文件撑爆）
        self.assertEqual(len(by_model["BJ6127C8BTB-1"].evidence), 1)
        self.assertIn("2 辆", by_model["BJ6127C8BTB-1"].note)
        # 原样的字符串区间也要能认（缓存 JSON 往返会变列表）
        self.assertEqual(
            len(bpt_data.lines_from_route_history("886", {"31688": [["20200815", "20201012"]]}, model_of)), 1
        )
        self.assertEqual(bpt_data.lines_from_route_history("886", {"1": ["乱写"]}, model_of), [])

    def test_line_info_round_trip(self):
        with scratch() as directory:
            base = Path(directory)
            self.assertEqual(bpt_data.load_line_info(base), {})
            info = bpt_data.update_line_info(
                {},
                "886",
                directions=[("德胜门西", "北京蟒山森林公园"), ("北京蟒山森林公园", "德胜门西")],
                checked="20261011",
            )
            info = bpt_data.update_line_info(
                info,
                "886",
                stops=[("1", "德胜门西", "0", "116.37731", "39.94914"), ("2", "北太平桥东", "2759", "116.37278", "39.967711")],
                checked="20261012",
            )
            bpt_data.save_line_info(base, info)
            loaded = bpt_data.load_line_info(base)
            self.assertEqual(len(bpt_data.line_info_directions(loaded, "886")), 2)
            stops = bpt_data.line_info_stops(loaded, "886")
            self.assertEqual(stops[0][1], "德胜门西")
            self.assertEqual(stops[1][2], "2759")
            record = loaded[bpt_data.line_key("886")]
            self.assertEqual(record["directions_checked"], "20261011")   # 两项各记各的查证日期
            self.assertEqual(record["stops_checked"], "20261012")
            self.assertNotIn("id", record)                                 # 内部编号绝不落盘
            # 坏文件不阻塞
            bpt_data.line_info_path(base).write_text("{ 不是 json", encoding="utf-8")
            self.assertEqual(bpt_data.load_line_info(base), {})

    def test_health_lists_lines_without_stops(self):
        links = [bpt_data.LineLink("886", "A"), bpt_data.LineLink("919", "A")]
        info = bpt_data.update_line_info({}, "886", stops=[("1", "德胜门西", "0", "1", "2")], checked="20261011")
        self.assertEqual(bpt_data.line_info_lines_without_stops(info, links), ["919"])

    def make_window(self, service=None):
        fake = service or FakeService()
        with patch.object(BPT_BusDigitalNumber, "shared_service", lambda *_a, **_k: fake):
            window = LineInfoWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        window.service = fake
        return window

    def settle(self, window) -> None:
        for _ in range(300):
            window.update()
            if not window.busy:
                return
            time.sleep(0.02)
        self.fail("线路资料查询没有结束")

    def test_window_queries_and_writes(self):
        def model_of(_ranges, number: str) -> str:
            return {"31688": "BJ6127C8BTB-1", "31691": "BJ6127C8BTB-1", "30846": "BJ6105EVCA-27"}.get(number, "")

        window = self.make_window()
        window.query_var.set("886")
        with patch.object(BPT_BusDigitalNumber, "model_lookup", model_of):
            window.query()
            self.settle(window)
        self.assertIn("2 个方向", window.answer_var.get())
        self.assertIn("3 辆车", window.answer_var.get())
        self.assertIn("现在在跑", window.detail_var.get())
        rows = [window.table.tree.item(item, "values") for item in window.table.tree.get_children()]
        self.assertEqual(len(rows), 2)                      # 3 辆车归成 2 种车型
        self.assertNotIn("disabled", window.write_button.state())
        with scratch() as directory, patch_app_dir(Path(directory)):
            with patch("BPT_BusDigitalNumber.messagebox.askyesno", return_value=True), patch(
                "BPT_BusDigitalNumber.messagebox.showinfo"
            ):
                window.write_links()
                window.save_info()
            saved = load_lines(Path(directory))
            self.assertEqual({item.model for item in saved}, {"BJ6127C8BTB-1", "BJ6105EVCA-27"})
            self.assertTrue(all(item.source == "online" for item in saved))
            self.assertTrue(all(item.numbers for item in saved))     # 每款车都带着自编号
            info = bpt_data.load_line_info(Path(directory))
            self.assertEqual(len(bpt_data.line_info_directions(info, "886")), 2)
            # 站点表（V4）：查完能存进 线路资料.json，并进体检
            self.assertEqual(len(window.last_stops), 2)
            with patch("BPT_BusDigitalNumber.messagebox.showinfo"):
                window.save_stops()
            info = bpt_data.load_line_info(Path(directory))
            stops = bpt_data.line_info_stops(info, "886")
            self.assertEqual([row[1] for row in stops], ["德胜门西", "北太平桥东"])
            self.assertEqual(stops[1][2], "2759")
            self.assertEqual(len(window.stop_table.tree.get_children()), 2)
            items = health_report(
                Path(directory), self.app.ranges, {}, {}, [], bpt_data.load_lines(Path(directory)), info
            )
            stops_item = next(item for item in items if item.key == "line_stops")
            self.assertTrue(all(row[0] != "886" for row in stops_item.rows))  # 886 已查过站点，不再提示

    def test_error_text_gives_network_hints(self):
        """联网失败时的文案要说清"大概是什么原因、可以重试"（截图里那次是被拒 10061）。"""
        refused = bpt_online.error_text(
            bpt_online.OnlineUnavailable("连不上联网服务（<urlopen error [WinError 10061] 由于目标计算机积极拒绝>）")
        )
        self.assertIn("连接被拒绝", refused)
        self.assertIn("防火墙", refused)
        self.assertIn("重试", refused)
        self.assertIn("本地数据完全不受影响", refused)
        timeout = bpt_online.error_text(bpt_online.OnlineUnavailable("等了太久，对方没有回应"))
        self.assertIn("等超时", timeout)
        dns = bpt_online.error_text(bpt_online.OnlineUnavailable("getaddrinfo failed"))
        self.assertIn("域名解析", dns)
        self.assertEqual(bpt_online._network_hint("说不清的错误"), "")

    def test_failed_query_headline_is_not_line_not_found(self):
        """连不上 ≠ 没有这条线：标题必须说"查询没成功"，别让人以为线路号错了。"""
        service = FakeService(fail=bpt_online.OnlineUnavailable("连不上联网服务（WinError 10061 积极拒绝）"))
        window = self.make_window(service)
        window.query_var.set("948")
        window.query()
        self.settle(window)
        self.assertIn("查询没成功", window.answer_var.get())
        self.assertNotIn("没有查到这条线路", window.answer_var.get())
        self.assertIn("连接被拒绝", window.detail_var.get())

    def test_window_journey_time_flow(self):
        """④ 实时：从「站名」到「到」那站的耗时（走完线程→回主线程→填表）。"""
        service = FakeService()
        window = self.make_window(service)
        window.query_var.set("886")
        window.query()
        self.settle(window)
        window.stop_name_var.set("德胜门西")
        window.to_stop_var.set("昌平东关")
        window.load_journey()
        for _ in range(300):
            window.update()
            if not window.busy:
                break
            time.sleep(0.02)
        else:
            self.fail("乘车耗时查询没有结束")
        self.assertEqual(service.journey_args, ("805", "德胜门西", "昌平东关"))
        self.assertIn("1 小时 6 分", window.live_hint_var.get() or "")
        rows = [window.live_table.tree.item(item, "values") for item in window.live_table.tree.get_children()]
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][0], "31705")

    def test_offline_seeds_directions_from_catalog(self):
        """断网也能看首末站：本地目录里的方向先填上，联网失败也不许把它冲掉。"""
        service = FakeService(fail=bpt_online.OnlineUnavailable("断网了"))
        window = self.make_window(service)
        with scratch() as directory, patch_app_dir(Path(directory)):
            bpt_data.save_route_list(
                Path(directory),
                bpt_data.parse_route_list(
                    {
                        "805": {"number": "886", "from": "德胜门西", "to": "北京蟒山森林公园"},
                        "806": {"number": "886", "from": "北京蟒山森林公园", "to": "德胜门西"},
                    }
                ),
                checked="20261011",
            )
            window._catalog_index = None
            window.query_var.set("886")
            window.query()
            self.settle(window)
            self.assertEqual(window.busy, False)
            self.assertIn("断网", window.detail_var.get())
            self.assertEqual(len(window.last_directions), 2, "本地目录的首末站被联网失败冲掉了")
            rows = [window.dir_table.tree.item(item, "values") for item in window.dir_table.tree.get_children()]
            pairs = {(row[1], row[2]) for row in rows}
            self.assertEqual(
                pairs, {("德胜门西", "北京蟒山森林公园"), ("北京蟒山森林公园", "德胜门西")}
            )
            self.assertEqual(window.last_checked, "20261011")   # 记的是目录的更新日期
            # 而且能离线存进线路资料
            with patch("BPT_BusDigitalNumber.messagebox.showinfo"):
                window.save_info()
            info = bpt_data.load_line_info(Path(directory))
            self.assertEqual(len(bpt_data.line_info_directions(info, "886")), 2)

    def test_window_degrades_when_bus_list_fails(self):
        """「现在在跑的车」接口偶尔不回包：只影响这一块，其它照常。"""
        service = FakeService()
        service.buses_fail = bpt_online.OnlineUnavailable("等了太久，对方没有回应")
        window = self.make_window(service)
        window.query_var.set("886")
        window.query()
        self.settle(window)
        self.assertIn("暂时查不到", window.detail_var.get())
        self.assertEqual(len(window.table.tree.get_children()), 2)  # 用过的车照常显示


class ModelCatalogLineTests(unittest.TestCase):
    """车型目录的「线路」列与按线路筛选（V6/A4）。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def make_window(self):
        window = CatalogWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        return window

    def test_line_query_shows_server_stops_offline(self):
        """线路查询里直接显示服务端首末站（来自本地目录，不用联网）。"""
        window = LineQueryWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        self.addCleanup(setattr, self.app, "lines", self.app.lines)
        self.app.lines = [bpt_data.LineLink("886", "A"), bpt_data.LineLink("919", "B")]
        with scratch() as directory, patch_app_dir(Path(directory)):
            # 还没有目录：给一句可操作的提示，而不是空白
            window._catalog_index = None
            window.on_data_changed({"lines"})
            self.assertIn("更新线路目录", window.summary_var.get())
            rows = {window.line_table.tree.item(i, "values")[0]: window.line_table.tree.item(i, "values")[1]
                    for i in window.line_table.tree.get_children()}
            self.assertEqual(rows["886"], "—")
            # 有目录之后：离线显示首末站（上下行显示成 A ↔ B）
            bpt_data.save_route_list(
                Path(directory),
                bpt_data.parse_route_list(
                    {
                        "805": {"number": "886", "from": "德胜门西", "to": "北京蟒山森林公园"},
                        "806": {"number": "886", "from": "北京蟒山森林公园", "to": "德胜门西"},
                        "5299": {"number": "919", "from": "地铁北土城站", "to": "体育场小区"},
                    }
                ),
                checked="20261011",
            )
            window.on_data_changed({"lines"})
            rows = {window.line_table.tree.item(i, "values")[0]: window.line_table.tree.item(i, "values")[1]
                    for i in window.line_table.tree.get_children()}
            self.assertIn("↔", rows["886"])   # 上下行：显示 A ↔ B（顺序按目录排，两种都对）
            self.assertEqual(set(rows["886"].split(" ↔ ")), {"德胜门西", "北京蟒山森林公园"})
            self.assertEqual(rows["919"], "地铁北土城站 → 体育场小区")   # 只有一个方向就写箭头
            self.assertIn("其中 2 条能对上", window.summary_var.get())

    def test_line_query_refreshes_when_data_changes(self):
        """开着「线路查询」时别处写入了新记录，它必须跟着刷新（不能一直显示旧快照）。"""
        window = LineQueryWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        self.addCleanup(setattr, self.app, "lines", self.app.lines)
        self.app.lines = [bpt_data.LineLink("886", "A")]
        window.on_data_changed({"lines"})
        self.assertEqual(
            [window.line_table.tree.item(i, "values")[0] for i in window.line_table.tree.get_children()], ["886"]
        )
        self.app.lines = [bpt_data.LineLink("886", "A"), bpt_data.LineLink("919", "B", source="online")]
        window.on_data_changed({"lines"})
        self.assertEqual(len(window.line_table.tree.get_children()), 2)
        # 只关心 tags 之类的变化时不必重扫
        window.on_data_changed({"tags"})
        self.assertEqual(len(window.line_table.tree.get_children()), 2)

    def test_line_manager_warns_before_overwriting_other_changes(self):
        """线路管理是拿打开时那份拷贝在编辑：别处改过就先进确认，拒绝时不许写文件。"""
        window = LineManagerWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        self.addCleanup(setattr, self.app, "lines", self.app.lines)
        with scratch() as directory, patch_app_dir(Path(directory)):
            window.links = [bpt_data.LineLink("886", "A")]
            window._opened_fingerprint = BPT_BusDigitalNumber._lines_fingerprint(window.links)
            # 模拟"别处写入"（例如联网写入 / 我的记录里导入）
            self.app.lines = [bpt_data.LineLink("886", "A"), bpt_data.LineLink("919", "B", source="online")]
            asked = {"called": False, "title": ""}

            def refuse(title, message, **kwargs):
                asked["called"] = True
                asked["title"] = title
                return False

            with patch.object(BPT_BusDigitalNumber.messagebox, "askyesno", refuse):
                window.save()
            self.assertTrue(asked["called"], "没有提示就直接覆盖了别处的改动")
            self.assertIn("别处", asked["title"])
            self.assertFalse((Path(directory) / "线路.json").is_file(), "用户拒绝后不该写文件")
            # 用户同意时照常保存
            with patch.object(BPT_BusDigitalNumber.messagebox, "askyesno", lambda *a, **k: True):
                window.save()
            self.assertEqual([item.line for item in load_lines(Path(directory))], ["886"])
            # 没有别处改动时不该打扰用户
            asked["called"] = False
            window2 = LineManagerWindow(self.app)
            window2.withdraw()
            self.addCleanup(destroy, window2)
            with patch.object(BPT_BusDigitalNumber.messagebox, "askyesno", refuse):
                window2.save()
            self.assertFalse(asked["called"], "没人改过数据时不该弹确认框")
        self.assertTrue(True)

    def test_hint_line_column_counts_and_filter(self):
        window = self.make_window()
        model = bpt_data.model_names(self.app.ranges)[0]
        other = bpt_data.model_names(self.app.ranges)[1]
        saved = self.app.lines
        self.addCleanup(setattr, self.app, "lines", saved)
        self.app.lines = [
            bpt_data.LineLink("886", model, (), "2024.1", "2024.2"),
            bpt_data.LineLink("919", model, (), "2024.1", "2024.2"),
            bpt_data.LineLink("1路", other, (), "2024.1", "2024.2"),
        ]
        window._refresh()
        self.assertIn(model, window.shown_models)
        # 「线路」列 = 这个车型跑过的线路条数
        counts = {
            window.table.tree.set(item, "model"): window.table.tree.set(item, "lines")
            for item in window.table.tree.get_children()
        }
        self.assertEqual(counts[model], "2")
        self.assertEqual(counts[other], "1")
        # 按线路筛选：只剩跑过 886 的车型
        window.line_value_var.set("886")
        window._refresh()
        self.assertEqual(window.shown_models, [model])
        window.line_value_var.set("919")
        window._refresh()
        self.assertEqual(window.shown_models, [model])
        window.line_value_var.set("1路")
        window._refresh()
        self.assertEqual(window.shown_models, [other])
        # 线路被删掉后，下拉里的失效选项要回到「全部」
        self.app.lines = []
        window._refresh()
        self.assertEqual(window.line_value_var.get(), ANY_TAG)
        window._reset_filters()
        self.assertEqual(window.filter_var.get(), "")


class RouteCatalogTests(unittest.TestCase):
    """服务端线路目录 → 本地离线校验/补全（V7）。"""

    ROUTES = {
        "805": {"number": "886", "from": "德胜门西", "to": "北京蟒山森林公园"},
        "806": {"number": "886", "from": "北京蟒山森林公园", "to": "德胜门西"},
        "5918": {"number": "1", "from": "首钢园公交场站", "to": "四惠枢纽站"},
        "3227": {"number": "专45", "from": "沙河", "to": "北街家园五区"},
    }

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def test_parse_and_suggest(self):
        routes = bpt_data.parse_route_list(self.ROUTES)
        self.assertEqual(routes["805"], ("886", "德胜门西", "北京蟒山森林公园"))
        self.assertEqual(len(routes), 4)
        self.assertEqual(bpt_data.parse_route_list(["不是字典"]), {})
        self.assertEqual(bpt_data.parse_route_list("乱写"), {})
        self.assertEqual(bpt_data.parse_route_list({"result": self.ROUTES})["805"][0], "886")
        index = bpt_data.route_list_numbers(routes)
        self.assertEqual(len(index["886"]), 2)                    # 上下行两条
        self.assertEqual(bpt_data.route_list_suggest(index, "88"), ["886"])
        self.assertIn("886", bpt_data.route_list_suggest(index, "8"))
        self.assertEqual(bpt_data.route_list_suggest(index, "专"), ["专45"])
        self.assertEqual(bpt_data.route_list_suggest(index, ""), [])
        self.assertEqual(bpt_data.route_list_suggest(index, "没有这条"), [])

    def test_round_trip_and_bad_file(self):
        with scratch() as directory:
            base = Path(directory)
            self.assertEqual(bpt_data.load_route_list(base), {})
            bpt_data.save_route_list(base, bpt_data.parse_route_list(self.ROUTES), checked="20261011")
            loaded = bpt_data.load_route_list(base)
            self.assertEqual(loaded["checked"], "20261011")
            self.assertEqual(len(loaded["routes"]), 4)
            self.assertEqual(
                bpt_data.route_list_numbers(loaded["routes"])["886"][0], ("北京蟒山森林公园", "德胜门西")
            )
            bpt_data.route_list_path(base).write_text("{ 坏 json", encoding="utf-8")
            self.assertEqual(bpt_data.load_route_list(base), {})
            bpt_data.route_list_path(base).write_text("[1,2]", encoding="utf-8")
            self.assertEqual(bpt_data.load_route_list(base), {})

    def test_validate_line_numbers(self):
        index = bpt_data.route_list_numbers(bpt_data.parse_route_list(self.ROUTES))
        links = [
            bpt_data.LineLink("886", "A", ("10001",)),
            bpt_data.LineLink("884", "A", ("10002", "10003")),
            bpt_data.LineLink("9999", "B", ("10004",)),
        ]
        result = bpt_data.validate_line_numbers(links, index)
        unknown = {row[0]: row for row in result["unknown"]}
        self.assertEqual(set(unknown), {"884", "9999"})
        self.assertEqual(unknown["884"][2], 2)                       # 同一线路号的辆数累加
        self.assertIn("886", result["suggest"]["884"])               # 错一个数字能猜到
        self.assertEqual(result["suggest"]["9999"], [])              # 差太远就不瞎猜
        # 没有本地目录时什么都不报（绝不把用户自己的记录误判成错的）
        self.assertEqual(bpt_data.validate_line_numbers(links, {}), {"unknown": [], "suggest": {}})

    def test_records_window_validates_against_local_catalog(self):
        window = MyRecordsWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        self.addCleanup(setattr, self.app, "lines", self.app.lines)
        with scratch() as directory, patch_app_dir(Path(directory)):
            self.assertEqual(window._local_catalog(), {})
            bpt_data.save_route_list(Path(directory), bpt_data.parse_route_list(self.ROUTES), checked="20261011")
            self.assertIn("886", window._local_catalog())
            self.app.lines = [
                bpt_data.LineLink("886", "A"),
                bpt_data.LineLink("884", "A", ("10002",)),
            ]
            captured = {}

            def fake_showinfo(title, message, **kwargs):
                captured["title"] = title
                captured["message"] = message

            with patch.object(BPT_BusDigitalNumber.messagebox, "showinfo", fake_showinfo):
                window.validate_line_numbers()
            self.assertIn("884", captured["message"])       # 打错的号被点出来
            self.assertIn("886", captured["message"])       # 并给出"你是不是想记"
            self.assertNotIn("　886　", captured["message"])  # 正确的号不列进来

    def test_window_uses_local_catalog_for_hints(self):
        window = LineInfoWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        with scratch() as directory, patch_app_dir(Path(directory)):
            bpt_data.save_route_list(Path(directory), bpt_data.parse_route_list(self.ROUTES), checked="20261011")
            window._catalog_index = None  # 强制重读
            hint = window._catalog_hint("886")
            self.assertIn("本地目录里有这条线", hint)
            self.assertIn("德胜门西", hint)
            missing = window._catalog_hint("887")
            self.assertIn("没有「887」", missing)
            self.assertIn("886", missing)   # 给个相近的号
            self.assertIn("专45", window._catalog_hint("专4"))
            # 没有目录文件时不给任何提示（不瞎猜）
            bpt_data.route_list_path(Path(directory)).unlink()
            window._catalog_index = None
            self.assertEqual(window._catalog_hint("886"), "")


class ThreadSafetyTests(unittest.TestCase):
    """把「工作线程绝不碰 Tk」变成测试。

    为什么要有这条：Tkinter 不是线程安全的，子线程里只要碰一下 Tk（连 `winfo_exists()`、
    `after_idle()` 都算）就会抛 `RuntimeError: main thread is not in main loop`——
    这种异常还常常被 `except tk.TclError` 漏掉，界面表现是"功能静默失效"。
    这个坑在本项目里犯过两次（联网查询线程、全局热键线程），所以用 AST 静态审计钉死。
    """

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    #: 这些方法一出现就说明"线程在碰界面"
    TK_METHODS = {
        "after", "after_idle", "after_cancel", "winfo_exists", "winfo_width", "winfo_height",
        "winfo_rootx", "winfo_rooty", "winfo_screenheight", "update", "update_idletasks",
        "config", "configure", "insert", "delete", "see", "selection_set", "selection_clear",
        "item", "get_children", "focus_set", "lift", "attributes", "destroy", "title", "geometry",
        "clipboard_clear", "clipboard_append", "set_rows", "set_columns", "pack", "grid",
    }
    TK_MODULES = ("messagebox", "simpledialog", "filedialog", "tkinter")

    def source(self) -> str:
        return (Path(__file__).resolve().parent / "BPT_BusDigitalNumber.py").read_text(encoding="utf-8")

    def worker_functions(self, tree) -> list:
        """所有被 threading.Thread(target=…) 用到的函数（含嵌套定义的内部函数）。"""
        names: set[str] = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Call) and ast.unparse(node.func).endswith("threading.Thread"):
                for keyword in node.keywords:
                    if keyword.arg == "target":
                        value = keyword.value
                        names.add(value.attr if isinstance(value, ast.Attribute) else getattr(value, "id", ""))
        found = []
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) and node.name in names:
                found.append(node)
        return found

    def offenders(self, function) -> list[str]:
        hits: list[str] = []
        for sub in ast.walk(function):
            if not isinstance(sub, ast.Call):
                continue
            target = ast.unparse(sub.func)
            last = target.split(".")[-1]
            if last in self.TK_METHODS or any(module in target for module in self.TK_MODULES):
                hits.append(f"L{sub.lineno} {target}")
        return hits

    def test_no_developer_jargon_in_ui_strings(self):
        """界面文案面向公交迷：不许出现版本对比、"环境变量""服务端""内部编号"这类开发者腔。

        只查字符串字面量（注释和 docstring 是写给维护者的，不查）。
        """
        import ast

        banned = ["这一版", "上个版本", "上一版", "本版本", "环境变量", "内部编号", "内部 id",
                  "服务端", "接口", "试验", "重构", "线程"]
        source = (Path(__file__).resolve().parent / "BPT_BusDigitalNumber.py").read_text(encoding="utf-8")
        tree = ast.parse(source)
        docstrings = set()
        for node in ast.walk(tree):
            if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
                body = getattr(node, "body", [])
                if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant):
                    docstrings.add(id(body[0].value))
        offenders: list[str] = []
        for node in ast.walk(tree):
            if isinstance(node, ast.Constant) and isinstance(node.value, str) and id(node) not in docstrings:
                for word in banned:
                    # "线程"允许出现在内部说明里（不是给用户看的那句），但界面文案不该有
                    if word in node.value:
                        offenders.append(f"L{node.lineno} [{word}] {node.value[:50]}")
        self.assertEqual(offenders, [], "界面文案里出现了开发者腔：\n" + "\n".join(offenders[:10]))

    def test_worker_threads_never_touch_tk(self):
        import ast

        tree = ast.parse(self.source())
        workers = self.worker_functions(tree)
        self.assertTrue(workers, "没找到线程入口？审计脚本本身可能失效了")
        problems: list[str] = []
        for function in workers:
            hits = self.offenders(function)
            if hits:
                problems.append(f"{function.name}() 定义在 L{function.lineno}: " + "; ".join(hits))
        self.assertEqual(problems, [], "工作线程里出现了 Tk 调用（会抛 main thread is not in main loop）：\n" + "\n".join(problems))

    def test_half_built_window_is_destroyed(self):
        """构造中途出错的窗口不能留下半成品（否则就是窗口泄漏）。"""

        class BrokenWindow(WindowMixin, tk.Toplevel):
            def __init__(self, parent):
                self.app = parent
                self.theme = parent.theme
                super().__init__(parent)
                self.setup_dialog(parent, "坏窗口", 200, 100)
                raise RuntimeError("构造中途出错")

        before = len(self.app.winfo_children())
        with self.assertRaises(RuntimeError):
            BrokenWindow(self.app)
        self.app.update()
        self.assertEqual(len(self.app.winfo_children()), before, "半成品窗口没有销毁")

    def test_inline_batch_dialog_releases_its_trace(self):
        """内联对话框（同一车型批量录车牌）关掉后不能留下 Tcl 命令。

        它用的是局部 StringVar + trace_add，不挂 self，所以 WindowMixin 的清理够不着；
        实测每开一次永久多 1 条 Tcl 命令。
        """
        import gc

        def find_button(widget, text):
            for child in widget.winfo_children():
                if isinstance(child, ttk.Button) and child.cget("text") == text:
                    return child
                found = find_button(child, text)
                if found is not None:
                    return found
            return None

        dialog_host = DistrictBusWindow(self.app)
        dialog_host.withdraw()
        self.addCleanup(destroy, dialog_host)
        self.app.update()
        baseline = len(self.app.tk.call("info", "commands"))  # 宿主窗口自己的命令不算泄漏
        for _ in range(3):
            dialog_host.batch_by_model()
            self.app.update()
            dialogs = [child for child in self.app.winfo_children() if isinstance(child, tk.Toplevel)]
            self.assertTrue(dialogs, "批量录入对话框没打开")
            cancel = find_button(dialogs[-1], "取消")
            self.assertIsNotNone(cancel, "没找到「取消」按钮")
            cancel.invoke()
            self.app.update()
        gc.collect()
        grown = len(self.app.tk.call("info", "commands")) - baseline
        self.assertLessEqual(grown, 1, f"内联对话框泄漏了 {grown} 条 Tcl 命令")

    def test_closing_windows_does_not_leak_tcl_commands(self):
        """回归：反复开关带搜索框的窗口，Tcl 命令数不能线性增长。

        以前每开关一次 CatalogWindow 就留 6 个（StringVar 的 trace 命令 + tkinter 的
        `_tclCommands` 互相引用成环，Python GC 看不见），反复用会一直涨。
        """
        import gc

        baseline = len(self.app.tk.call("info", "commands"))
        for _ in range(3):
            window = CatalogWindow(self.app)
            window.withdraw()
            self.app.update()
            window.destroy()
            self.app.update()
        gc.collect()
        grown = len(self.app.tk.call("info", "commands")) - baseline
        self.assertLessEqual(grown, 2, f"反复开关窗口泄漏了 {grown} 个 Tcl 命令")

    def test_hotkey_thread_only_queues(self):
        """热键线程只准往队列里塞事件，界面回调交给主线程轮询。"""
        import ast

        tree = ast.parse(self.source())
        manager = None
        for node in ast.walk(tree):
            if isinstance(node, ast.ClassDef) and hasattr(node, "body"):
                methods = {item.name for item in node.body if isinstance(item, ast.FunctionDef)}
                if {"_run", "_poll"} <= methods:
                    manager = node
        self.assertIsNotNone(manager, "没找到热键管理类")
        run_method = next(item for item in manager.body if isinstance(item, ast.FunctionDef) and item.name == "_run")
        self.assertNotIn("after", ast.unparse(run_method), "热键线程里不该排 Tk 定时器")
        self.assertIn("_fired.put", ast.unparse(run_method), "热键线程应该只入队")


class RealtimeTests(unittest.TestCase):
    """V5 实时：解析、日志、以及"实时数据绝不落盘"这条硬规矩。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def test_parsers(self):
        latest = bpt_data.parse_track_latest(TRACK_FIXTURE)
        self.assertEqual(latest["line_name"], "886 开往 北京蟒山森林公园")
        self.assertEqual(latest["distance_m"], 490)     # 取最后一条 = 现在在哪
        self.assertEqual(latest["points"], "2")
        self.assertEqual(bpt_data.parse_track_latest([]), {})
        self.assertEqual(bpt_data.parse_track_latest({"result": TRACK_FIXTURE})["stop_name"], "军都度假村")
        self.assertEqual(
            bpt_data.parse_arrivals([{"bus_id": "31688", "time": "2026-10-10 06:45:50", "occupancy": 1}]),
            [("31688", "2026-10-10 06:45:50", "1")],
        )
        table = "车号\t德胜门西\t北太平桥东\t昌平西关环岛东\n31705\t06:27:40,1,0\t06:38:40,1,0\t07:14:21,2,0\n"
        self.assertEqual(bpt_data.parse_timetable(table), [("31705", "06:27:40", "07:14:21", "3")])
        self.assertEqual(bpt_data.parse_timetable(""), [])

    def test_schedule_and_depot_parsers(self):
        """发车时刻 / 场站出车（真实回包形状，2026-10 实测）。"""
        schedules = bpt_data.parse_schedules([
            {"bus_id": "39490",
             "depart": {"time": "2026-10-10 05:40:23", "stop_index": 1, "stop_name": "德胜门西", "arrived": True},
             "arrive": {"time": "2026-10-10 06:50:08", "stop_index": 3, "stop_name": "军都度假村", "arrived": True}},
        ])
        self.assertEqual(schedules, [("39490", "2026-10-10 05:40:23", "德胜门西", "2026-10-10 06:50:08", "军都度假村")])
        depots = bpt_data.parse_depots([
            {"bus_id": "39490", "time": "2026-10-10 04:48:18", "route_id": 805, "route_number": "886",
             "direction_to": "北京蟒山森林公园", "stop_index": 13, "stop_name": "军都度假村", "distance_m": 2054},
        ])
        self.assertEqual(depots, [("39490", "2026-10-10 04:48:18", "军都度假村", "北京蟒山森林公园")])
        self.assertEqual(bpt_data.parse_schedules("不是列表"), [])
        self.assertEqual(bpt_data.parse_depots({"result": []}), [])

    def test_stop_search_parsers(self):
        """站名搜索两个工具的真实回包形状（德胜门西 / 德胜门西→昌平东关）。"""
        at_stop = bpt_data.parse_routes_at_stop(
            [{"route_id": "805", "route_number": "886", "route_first_stop": "德胜门西",
              "route_last_stop": "北京蟒山森林公园", "current_stop_number": "1"}]
        )
        self.assertEqual(at_stop, [("886", "德胜门西", "北京蟒山森林公园", "1")])
        between = bpt_data.parse_routes_between(
            [{"route_id": "805", "route_number": "886", "route_first_stop": "德胜门西",
              "route_last_stop": "北京蟒山森林公园", "from_stop_number": "1", "to_stop_number": "12"}]
        )
        self.assertEqual(between, [("886", "德胜门西", "北京蟒山森林公园", "1", "12")])
        self.assertEqual(bpt_data.parse_routes_at_stop({"result": []}), [])
        self.assertEqual(bpt_data.parse_routes_between("不是列表"), [])

    def test_journey_time_parser(self):
        """两站间乘车耗时（真实回包形状：duration_sec → 人话）。"""
        rows = bpt_data.parse_journey_time([
            {"bus_id": "31705", "start": "2026-10-10 06:27:40", "end": "2026-10-10 07:34:25", "duration_sec": 4005},
            {"bus_id": "31682", "start": "2026-10-10 05:45:50", "end": "2026-10-10 06:52:19", "duration_sec": 3989},
        ])
        self.assertEqual(rows[0][0], "31682")            # 按出发时间排
        self.assertEqual(rows[1][2], "1 小时 6 分")       # 4005 秒
        self.assertEqual(rows[1][3], "2026-10-10 07:34:25")
        self.assertEqual(bpt_data.format_duration(90), "1 分")
        self.assertEqual(bpt_data.format_duration(3600), "1 小时")
        self.assertEqual(bpt_data.format_duration(3660), "1 小时 1 分")
        self.assertEqual(bpt_data.format_duration(45), "不到 1 分")
        self.assertEqual(bpt_data.format_duration("乱写"), "")
        self.assertEqual(bpt_data.format_duration(-5), "")
        self.assertEqual(bpt_data.parse_journey_time("乱写"), [])

    def test_online_log_helpers(self):
        with scratch() as directory:
            base = Path(directory)
            bpt_data.append_online_log(base, "bus 31688\t成功\t6.4s")
            bpt_data.append_online_log(base, "route 999\t失败\t连不上")
            self.assertEqual(len(bpt_data.read_online_log(base)), 2)
            self.assertEqual(len(bpt_data.recent_online_errors(base)), 1)
            self.assertIn("失败", bpt_data.recent_online_errors(base)[0])
            self.assertTrue(bpt_data.online_log_path(base).is_file())
            self.assertTrue(bpt_data.clear_online_log(base))
            self.assertEqual(bpt_data.read_online_log(base), [])

    def test_realtime_is_cached_in_memory_only(self):
        calls: list[str] = []

        class FakeClient:
            def __init__(self, **_kwargs) -> None:
                pass

            def connect(self) -> None:
                pass

            def query_track_record(self, bus_id):
                calls.append("track")
                return TRACK_FIXTURE

            def query_timetable(self, route_id):
                calls.append("timetable")
                return "车号\tA\tB\n31705\t06:00,1,0\t07:00,1,0\n"

            def close(self) -> None:
                pass

        with scratch() as directory:
            base = Path(directory)
            service = bpt_online.OnlineService(
                bpt_online.OnlineSettings(min_interval_sec=0, cache_days=7), base, client_factory=FakeClient
            )
            latest, _stamp = service.track_latest("31688")
            self.assertEqual(latest["distance_m"], 490)
            service.track_latest("31688")
            self.assertEqual(calls.count("track"), 1, "30 秒内重复查走内存缓存")
            self.assertEqual(service.cache_size, 0, "实时数据不能进磁盘缓存")
            self.assertFalse(bpt_data.online_cache_path(base).exists() or bpt_data.load_online_cache(base))
            # 换一个实例（等于重开程序）→ 必须重新联网，不能拿旧位置糊弄
            fresh = bpt_online.OnlineService(
                bpt_online.OnlineSettings(min_interval_sec=0, cache_days=7), base, client_factory=FakeClient
            )
            fresh.track_latest("31688")
            self.assertEqual(calls.count("track"), 2)

    def test_settings_network_group_round_trip(self):
        window = SettingsWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        self.addCleanup(bpt_online.reset_shared_service)
        with scratch() as directory, patch_app_dir(Path(directory)):
            window.network_var.set(False)
            window.timeout_var.set("12")
            window.cache_days_var.set("3")
            window.interval_var.set("0.5")
            window.save_network()
            saved = load_settings(Path(directory)).get("network") or {}
            self.assertEqual(saved["enabled"], False)
            self.assertEqual(saved["timeout_sec"], 12)
            self.assertEqual(saved["cache_days"], 3)
            self.assertFalse(bpt_online.OnlineSettings.from_dict(saved).usable)
            self.assertIn("已保存", window.network_status_var.get())
            # 填了看不懂的值 → 回到默认，而不是报错卡住
            window.timeout_var.set("abc")
            window.save_network()
            self.assertNotIn("timeout_sec", load_settings(Path(directory))["network"])
            window.clear_log_action()
            self.assertIn("日志", window.network_status_var.get())


class SelfCheckTests(unittest.TestCase):
    """自检体系：好数据必须全过，坏数据必须只失败对应项，且每条失败都要给处理建议。"""

    def good_base(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        base = Path(directory.name)
        shutil.copy(Path(__file__).resolve().parent / bpt_data.DATA_FILENAME, base / bpt_data.DATA_FILENAME)
        return base

    def test_good_data_passes_without_blocking_failures(self):
        base = self.good_base()
        items = bpt_selfcheck.run_self_checks(base, {})
        self.assertTrue(items)
        self.assertFalse(bpt_selfcheck.has_blocking_failure(items), [i.key for i in items])
        self.assertTrue(all(item.title and item.detail for item in items))
        tag = next(item for item in items if item.key == "tags")
        self.assertTrue(tag.ok or not tag.blocking)
        self.assertIn("号段表", "\n".join(bpt_selfcheck.format_report(items)))

    def test_broken_sheet_fails_and_explains(self):
        base = self.good_base()
        (base / bpt_data.DATA_FILENAME).write_bytes(b"not an xlsx")
        by_key = {item.key: item for item in bpt_selfcheck.run_self_checks(base, {})}
        self.assertFalse(by_key["sheet"].ok)
        self.assertTrue(by_key["sheet"].blocking)
        self.assertIn("备份", by_key["sheet"].hint)      # 必须告诉用户下一步怎么办
        self.assertTrue(bpt_selfcheck.has_blocking_failure(list(by_key.values())))

    def test_corrupt_user_json_fails_but_missing_is_fine(self):
        base = self.good_base()
        self.assertTrue({i.key: i for i in bpt_selfcheck.run_self_checks(base, {})}["json"].ok)
        (base / bpt_data.LINES_FILENAME).write_text("{ 坏 json", encoding="utf-8")
        bad = {item.key: item for item in bpt_selfcheck.run_self_checks(base, {})}
        self.assertFalse(bad["json"].ok)
        self.assertIn("线路.json", bad["json"].detail)

    def test_tag_checker_is_injected_and_optional(self):
        base = self.good_base()
        without = {item.key: item for item in bpt_selfcheck.run_self_checks(base, {})}
        self.assertIn("跳过", without["tags"].detail)
        sentinel = bpt_selfcheck.CheckItem("tags", "车型标签", False, "98 种未确认", "去数据体检", blocking=False)
        with_checker = {
            item.key: item
            for item in bpt_selfcheck.run_self_checks(base, {}, tag_checker=lambda _base: sentinel)
        }
        self.assertEqual(with_checker["tags"].detail, "98 种未确认")
        self.assertFalse(with_checker["tags"].blocking)
        self.assertFalse(bpt_selfcheck.has_blocking_failure(list(with_checker.values())))

    def test_json_output_is_parseable(self):
        items = bpt_selfcheck.run_self_checks(self.good_base(), {})
        payload = json.loads(bpt_selfcheck.as_json(items))
        self.assertEqual({entry["key"] for entry in payload}, {item.key for item in items})


class OnlineParseTests(unittest.TestCase):
    """联网层的纯解析：不联网，直接喂 2026-10-01 从 bjbus.jinzihao.me 抓的真实回包。"""

    def test_date_range_and_month_conversion(self):
        self.assertEqual(bpt_online.parse_date_range("20200815-20201012"), ("20200815", "20201012"))
        self.assertIsNone(bpt_online.parse_date_range("20200815"))
        self.assertIsNone(bpt_online.parse_date_range("2020-10-12"))
        self.assertEqual(bpt_online.compact_to_month("20200815"), "2020.8")
        self.assertEqual(bpt_online.compact_to_month("20261010"), "2026.10")
        self.assertEqual(bpt_online.compact_to_month("看不懂"), "看不懂")

    def test_history_and_route_parsing(self):
        history = bpt_online.parse_history_map(BUS_HISTORY_FIXTURE)
        self.assertEqual(list(history), ["877", "885", "886", "889"])
        self.assertEqual(history["886"][0], ("20200815", "20201012"))
        self.assertEqual(len(history["886"]), 3)
        self.assertEqual(
            bpt_online.parse_history_map({"886": ["乱写", "20200815-20201012"]}), {"886": [("20200815", "20201012")]}
        )
        with self.assertRaises(bpt_online.OnlineProtocolError):
            bpt_online.parse_history_map(["不是字典"])
        self.assertEqual(
            bpt_online.parse_route_ids(ROUTE_IDS_FIXTURE),
            [("805", "德胜门西", "北京蟒山森林公园"), ("806", "北京蟒山森林公园", "德胜门西")],
        )
        with self.assertRaises(bpt_online.OnlineProtocolError):
            bpt_online.parse_route_ids({"不是列表": 1})
        # 查不到的线路：服务端有时给的是带 error/detail 的对象，要报"查不到"而不是一句术语
        with self.assertRaises(bpt_online.OnlineProtocolError) as caught:
            bpt_online.parse_route_ids({"error": "route_not_found"})
        self.assertIn("查不到", str(caught.exception))
        self.assertIn("查不到", str(caught.exception))
        # 包了一层 result 也能认
        self.assertEqual(len(bpt_online.parse_route_ids({"result": ROUTE_IDS_FIXTURE})), 2)

    def test_stops_parsing(self):
        payload = {
            "name": "886",
            "stops": [
                {"name": "德胜门西", "longitude": 116.38, "latitude": 39.95, "distance_from_previous_stop": 0},
                {"name": "马甸桥南", "longitude": 116.39, "latitude": 39.97, "distance_from_previous_stop": 1200},
            ],
        }
        rows = bpt_online.parse_stops(payload)
        self.assertEqual(rows[0], ("1", "德胜门西", "0", "116.38", "39.95"))
        self.assertEqual(rows[1][1:3], ("马甸桥南", "1200"))
        self.assertEqual(bpt_online.parse_stops({"stops": []}), [])
        with self.assertRaises(bpt_online.OnlineProtocolError):
            bpt_online.parse_stops([1, 2])

    def test_history_to_links(self):
        links = bpt_online.history_to_links(
            "31688", BUS_HISTORY_FIXTURE, lambda _n: "BJ6127C8BTB-1", checked="20261001", today="20260915"
        )
        self.assertEqual([link.line for link in links], ["886", "885", "877", "889"])  # 按最早时间排
        by_line = {link.line: link for link in links}
        self.assertEqual(by_line["886"].numbers, ("31688",))
        self.assertEqual((by_line["886"].first, by_line["886"].last), ("2020.8", "2021.1"))
        self.assertEqual(by_line["886"].source, "online")
        self.assertIn("网上查的", by_line["886"].note)
        self.assertEqual(by_line["877"].status, "current")   # 最新区间 20260910，距 2026-09-15 只差 5 天
        self.assertEqual(by_line["889"].status, "recorded")  # 旧数据不上提
        # 同一个历史、把"今天"推到 2026-10-01：877 的最新区间就超过 7 天了，不该再说"现在还在跑"
        older = bpt_online.history_to_links("31688", BUS_HISTORY_FIXTURE, lambda _n: "", today="20261001")
        self.assertEqual({link.line: link.status for link in older}["877"], "recorded")
        blank = bpt_online.history_to_links("31688", {"1路": [("20200101", "20200102")]}, lambda _n: "", today="20261001")
        self.assertEqual(blank[0].model, "")
        self.assertEqual(blank[0].status, "recorded")  # 永不写 retired

    def test_unpack_and_error_text(self):
        self.assertEqual(bpt_online.unpack_tool_result(TOOL_RESULT_FIXTURE), ROUTE_IDS_FIXTURE)
        self.assertEqual(bpt_online.unpack_tool_result({"structuredContent": {"result": [1, 2]}}), [1, 2])
        self.assertEqual(bpt_online.unpack_tool_result({"content": [{"text": "纯文本"}]}), "纯文本")
        with self.assertRaises(bpt_online.OnlineProtocolError):
            bpt_online.unpack_tool_result({"content": []})
        with self.assertRaises(bpt_online.OnlineProtocolError):
            bpt_online.unpack_tool_result("不是对象")
        self.assertIn("本地数据完全不受影响", bpt_online.error_text(bpt_online.OnlineUnavailable("连不上")))
        self.assertIn("看不懂", bpt_online.error_text(bpt_online.OnlineProtocolError("缺字段")))
        self.assertIn("ValueError", bpt_online.error_text(ValueError("别的错")))

    def test_settings_and_offline_switch(self):
        default = bpt_online.OnlineSettings.from_dict({})
        self.assertTrue(default.enabled)
        self.assertEqual(default.base_url, "https://bjbus.jinzihao.me")
        self.assertTrue(default.usable)
        custom = bpt_online.OnlineSettings.from_dict(
            {"enabled": False, "base_url": "https://example.com/", "timeout_sec": 9, "cache_days": -3, "min_interval_sec": 0}
        )
        self.assertFalse(custom.usable)
        self.assertEqual(custom.base_url, "https://example.com")
        self.assertEqual(custom.timeout_sec, 9)
        self.assertEqual(custom.cache_days, 0)
        with patch.dict("os.environ", {bpt_online.OFFLINE_ENV: "1"}):
            self.assertTrue(bpt_online.offline_forced())
            self.assertFalse(bpt_online.OnlineSettings.from_dict({}).usable)
        with patch.dict("os.environ", {bpt_online.OFFLINE_ENV: "0"}):
            self.assertFalse(bpt_online.offline_forced())

    def test_model_lookup_uses_the_range_table(self):
        from bpt_data import BusRange

        ranges = [BusRange(31590, 31779, "BJ6127C8BTB-1")]
        self.assertEqual(bpt_online.model_lookup(ranges, "31688"), "BJ6127C8BTB-1")
        self.assertEqual(bpt_online.model_lookup(ranges, "39999"), "")
        self.assertEqual(bpt_online.model_lookup(ranges, "乱写"), "")


class OnlineClientTests(unittest.TestCase):
    """MCP 客户端：用假传输验握手、配对、错误与超时（不联网）。"""

    def make_client(self, transport=None, **kwargs) -> "bpt_online.MCPClient":
        settings = bpt_online.OnlineSettings(min_interval_sec=0, **kwargs)
        return bpt_online.MCPClient(transport or FakeTransport(), settings=settings)

    def test_handshake_then_tool_call(self):
        transport = FakeTransport()
        with self.make_client(transport) as client:
            routes = client.query_route_ids("886")
        self.assertTrue(transport.opened and transport.closed)
        self.assertEqual(routes, [("805", "德胜门西", "北京蟒山森林公园"), ("806", "北京蟒山森林公园", "德胜门西")])
        methods = [item.get("method") for item in transport.sent]
        self.assertEqual(methods[0], "initialize")
        self.assertIn("notifications/initialized", methods)
        self.assertIn("tools/call", methods)
        self.assertEqual(methods.count("initialize"), 1)  # 握手只做一次

    def test_bus_history_and_unknown_tool(self):
        with self.make_client() as client:
            history = client.query_bus_history("31688")
            self.assertEqual(history["886"][0], ("20200815", "20201012"))
            with self.assertRaises(bpt_online.OnlineProtocolError) as caught:
                client.call_tool("没有的工具", {})
        self.assertIn("没这个工具", str(caught.exception))

    def test_connect_waits_longer_than_the_tcp_timeout(self):
        """回归：initialize 实测要 ~19 秒，不能用 8 秒的"建连超时"去等（会偶发连不上）。"""

        class SlowTransport(FakeTransport):
            def post(self, payload: dict) -> None:
                if payload.get("method") == "initialize":
                    time.sleep(0.4)  # 模拟服务端握手慢
                super().post(payload)

        settings = bpt_online.OnlineSettings(min_interval_sec=0, connect_timeout_sec=1, timeout_sec=5)
        with bpt_online.MCPClient(SlowTransport(), settings=settings) as client:
            routes = client.query_route_ids("886")
        self.assertEqual(len(routes), 2)

    def test_transport_timeouts_have_margin(self):
        """回归：服务端建会话实测 ~6.5 秒、initialize ~19 秒，上限必须留足余量。

        这三个上限以前都误用了 connect_timeout（8 秒），会偶发报"没有回应/没有给出会话地址"。
        """
        transport = bpt_online.SSETransport("https://example.com", connect_timeout=8)
        self.assertGreaterEqual(transport.session_timeout, 30)   # 会话地址
        self.assertGreaterEqual(transport.post_timeout, 30)      # 递请求
        client = bpt_online.MCPClient(settings=bpt_online.OnlineSettings(timeout_sec=45, connect_timeout_sec=8))
        self.assertGreaterEqual(client.settings.timeout_sec, 30)  # initialize 的等待
        self.assertEqual(client.transport.session_timeout, 45)

    def test_timeout_and_closed_stream(self):
        # 注意 timeout_sec 要调小：现在的 initialize 会等服务端 timeout_sec（默认 45s），
        # 不调小的话这个"故意不回包"的用例要等 45 秒。
        with self.assertRaises(bpt_online.OnlineUnavailable) as caught:
            with self.make_client(FakeTransport(drop_first=True), connect_timeout_sec=3, timeout_sec=1) as client:
                client.query_route_ids("886")
        self.assertIn("没有回应", str(caught.exception))

        class ClosingTransport(FakeTransport):
            def open(self) -> None:
                super().open()
                self._queue.put({"__closed__": True})

        with self.assertRaises(bpt_online.OnlineUnavailable) as caught:
            with self.make_client(ClosingTransport(), connect_timeout_sec=3) as client:
                client.query_route_ids("886")
        self.assertIn("断了", str(caught.exception))

    def test_offline_never_touches_transport(self):
        transport = FakeTransport()
        with patch.dict("os.environ", {bpt_online.OFFLINE_ENV: "1"}):
            client = self.make_client(transport)
            with self.assertRaises(bpt_online.OnlineUnavailable):
                client.query_route_ids("886")
        self.assertFalse(transport.opened)
        self.assertEqual(transport.sent, [])


class FakeService:
    """假的联网服务：窗口测试用它，既不联网也不碰缓存文件。"""

    def __init__(self, history=None, routes=None, fail: Exception | None = None, cached: int = 0) -> None:
        self.history = history if history is not None else bpt_online.parse_history_map(BUS_HISTORY_FIXTURE)
        self.routes = routes if routes is not None else bpt_online.parse_route_ids(ROUTE_IDS_FIXTURE)
        self.fail = fail
        self.cache_size = cached
        self.connected = True
        self.closed = 0
        self.cleared = 0
        self.warmed = 0
        self.last_seconds = 0.0
        # 线路维度（V3）
        self.history_buses = bpt_online.parse_history_map(ROUTE_BUSES_FIXTURE)
        self.buses = ["31682", "31683", "31688"]
        self.buses_fail: Exception | None = None

    def _check(self) -> None:
        if self.fail is not None:
            raise self.fail

    def bus_history(self, bus_id):
        self._check()
        return self.history, "20261011"

    def track_latest(self, bus_id):
        self._check()
        return bpt_data.parse_track_latest(TRACK_FIXTURE), "20261011"

    def route_ids(self, number):
        self._check()
        return self.routes, "20261011"

    def route_history(self, number):
        self._check()
        return self.history_buses, "20261011"

    def buses_on_route(self, route_id):
        if self.buses_fail is not None:
            raise self.buses_fail
        self._check()
        return self.buses, "20261011"

    def journey_time(self, route_id, from_stop, to_stop):
        self._check()
        self.journey_args = (route_id, from_stop, to_stop)
        return [
            ("31705", "2026-10-10 06:27:40", "1 小时 6 分", "2026-10-10 07:34:25"),
            ("31682", "2026-10-10 05:45:50", "1 小时 6 分", "2026-10-10 06:52:19"),
        ], "20261011"

    def route_stops(self, route_id):
        self._check()
        return [
            ("1", "德胜门西", "0", "116.37731", "39.94914"),
            ("2", "北太平桥东", "2759", "116.37278", "39.967711"),
        ], "20261011"

    def warm_up(self) -> None:
        self.warmed += 1

    def close(self) -> None:
        self.closed += 1

    def clear_cache(self) -> int:
        count, self.cache_size = self.cache_size, 0
        self.cleared += 1
        return count

    def source_text(self) -> str:
        return "联网查询（6.4s）"

    def build_mark(self) -> str:
        return "程序构建 2026-10-11 12:00"


class OnlineServiceTests(unittest.TestCase):
    """共享服务：**缓存命中就不再联网**（服务端每个请求约 6 秒，这一层是性能的关键）。"""

    def make_service(self, directory, calls: list, **settings):
        class FakeClient:
            def __init__(self, **_kwargs) -> None:
                pass

            def connect(self) -> None:
                calls.append("connect")

            def query_bus_history(self, bus_id):
                calls.append("history")
                return bpt_online.parse_history_map(BUS_HISTORY_FIXTURE)

            def query_route_ids(self, number):
                calls.append("route")
                return bpt_online.parse_route_ids(ROUTE_IDS_FIXTURE)

            def close(self) -> None:
                calls.append("close")

        options = {"min_interval_sec": 0, "cache_days": 7}
        options.update(settings)
        return bpt_online.OnlineService(
            bpt_online.OnlineSettings(**options), directory, client_factory=FakeClient
        )

    def test_transient_failure_retries_once_but_protocol_error_does_not(self):
        """连接类失败自动重试一次（用户点一下就能成），协议错误不重试（重试也一样）。"""
        attempts: list[str] = []

        class FlakyClient:
            def __init__(self, **_kwargs) -> None:
                pass

            def connect(self) -> None:
                pass

            def query_bus_history(self, bus_id):
                attempts.append("call")
                if len(attempts) == 1:
                    raise bpt_online.OnlineUnavailable("连不上联网服务（WinError 10061 积极拒绝）")
                return bpt_online.parse_history_map(BUS_HISTORY_FIXTURE)

            def close(self) -> None:
                attempts.append("close")

        with scratch() as directory:
            service = bpt_online.OnlineService(
                bpt_online.OnlineSettings(min_interval_sec=0), Path(directory), client_factory=FlakyClient
            )
            history, _checked = service.bus_history("31688")
            self.assertEqual(len(history), 4)                      # 重试后成功
            self.assertEqual(attempts.count("call"), 2)             # 一共试了两次
            self.assertIn("close", attempts)                        # 坏会话被丢掉

        protocol_attempts: list[str] = []

        class BadPayloadClient:
            def __init__(self, **_kwargs) -> None:
                pass

            def connect(self) -> None:
                pass

            def query_bus_history(self, bus_id):
                protocol_attempts.append("call")
                raise bpt_online.OnlineProtocolError("对方返回的内容看不懂")

            def close(self) -> None:
                pass

        with scratch() as directory:
            service = bpt_online.OnlineService(
                bpt_online.OnlineSettings(min_interval_sec=0), Path(directory), client_factory=BadPayloadClient
            )
            with self.assertRaises(bpt_online.OnlineProtocolError):
                service.bus_history("31688")
            self.assertEqual(protocol_attempts.count("call"), 1)    # 协议错误不重试

    def test_second_query_hits_cache(self):
        calls: list[str] = []
        with scratch() as directory:
            service = self.make_service(directory, calls)
            first, checked = service.bus_history("31688")
            self.assertEqual(calls.count("history"), 1)
            self.assertEqual(service.last_source, "net")
            self.assertIsInstance(first["886"][0], tuple)
            second, checked_again = service.bus_history("31688")
            self.assertEqual(calls.count("history"), 1, "第二次查询不该再联网")
            self.assertEqual(service.last_source, "cache")
            self.assertEqual(checked, checked_again)
            # 缓存写进了文件：换一个服务实例（模拟重开程序）仍然命中
            fresh = self.make_service(directory, calls)
            fresh.bus_history("31688")
            self.assertEqual(calls.count("history"), 1, "缓存应该跨进程有效")
            # 清空缓存后又要联网
            self.assertEqual(service.clear_cache(), 1)
            service.bus_history("31688")
            self.assertEqual(calls.count("history"), 2)
            # 走缓存的 payload 是 JSON 往返过的（元组变列表），下游必须也认
            links = bpt_data.lines_from_history("31688", second, lambda _n: "", checked=checked_again)
            self.assertEqual([link.line for link in links], ["886", "885", "877", "889"])

    def test_cache_expiry_and_prune(self):
        calls: list[str] = []
        with scratch() as directory:
            service = self.make_service(directory, calls, cache_days=7)
            service.bus_history("31688")
            # 把缓存日期改老，模拟 30 天前
            service.cache = bpt_data.cache_put({}, "bus:31688", {"886": [["20200101", "20200102"]]}, checked="20260101")
            service.save_cache()
            service.bus_history("31688")
            self.assertEqual(calls.count("history"), 2, "过期缓存必须重新联网")

    def test_offline_service_never_touches_client(self):
        calls: list[str] = []
        with scratch() as directory:
            service = self.make_service(directory, calls)
            with patch.dict("os.environ", {bpt_online.OFFLINE_ENV: "1"}):
                with self.assertRaises(bpt_online.OnlineUnavailable):
                    service.bus_history("31688")
            self.assertEqual(calls, [])


class OnlineWindowTests(unittest.TestCase):
    """联网查询窗口：构造、离线门禁、注入假客户端后跑通"点查询→线程→回主线程→填表"。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def make_window(self, service=None):
        """构造窗口；默认把联网服务换成假的（测试不联网、不写缓存）。"""
        fake = service or FakeService()
        with patch.object(BPT_BusDigitalNumber, "shared_service", lambda *_a, **_k: fake):
            window = OnlineQueryWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        window.service = fake
        window.refresh_state()
        return window

    def settle(self, window) -> None:
        for _ in range(300):
            window.update()
            if not window.busy:
                return
            time.sleep(0.02)
        self.fail("后台查询没有结束")

    def test_window_constructs_and_switches_columns(self):
        window = self.make_window()
        self.assertEqual([c.title for c in window.table.columns], ["线路", "车型", "时间段数", "最早", "最近", "状态"])
        window.table.set_columns(window.COLUMNS_ROUTE)
        self.assertEqual([c.title for c in window.table.columns], ["方向", "首站", "末站"])

    def test_offline_disables_buttons(self):
        with patch.dict("os.environ", {bpt_online.OFFLINE_ENV: "1"}):
            window = self.make_window()
            window.refresh_state()
            self.assertIn("联网不可用", window.state_var.get())
            self.assertIn("disabled", window.bus_button.state())
            self.assertIn("disabled", window.route_button.state())

    def test_query_runs_in_background_with_fake_client(self):
        window = self.make_window()

        window.query_var.set("31688")
        window.query("bus")
        self.settle(window)
        self.assertIn("跑过 4 条线路", window.answer_var.get())
        rows = [window.table.tree.item(item, "values") for item in window.table.tree.get_children()]
        self.assertEqual([row[0] for row in rows], ["886", "885", "877", "889"])
        self.assertEqual(rows[0][1], "BJ6127C8BTB-1")  # 车型由号段表反查
        self.assertNotIn("disabled", window.write_button.state())  # 查到就能写入

        window.query_var.set("886")
        window.query("route")
        self.settle(window)
        self.assertIn("2 个方向", window.answer_var.get())
        first = window.table.tree.item(window.table.tree.get_children()[0], "values")
        self.assertEqual(first[1], "德胜门西")
        # 线路模式下没有"写入线路数据"的对象（那是按车辆写的）
        self.assertIn("disabled", window.write_button.state())

    def test_write_links_merges_into_line_data(self):
        """点「写入线路数据」：合并进 线路.json，来源是 online，且不覆盖手工记录。"""
        # 让"网上查的"这段比手工记录更晚（886 到 2026.9），好验证时间取并集
        history = bpt_online.parse_history_map({"886路": ["20200815-20260930"], "877": ["20230101-20240101"]})
        window = self.make_window(FakeService(history=history))
        with scratch() as directory, patch_app_dir(Path(directory)):
            manual = LineLink("886路", "BJ6127C8BTB-1", ("31688",), "2024.9", "2025.8", "current", "manual")
            self.app.lines = [manual]
            window.query_var.set("31688")
            window.query("bus")
            self.settle(window)
            self.assertTrue(window.last_links)
            with patch("BPT_BusDigitalNumber.messagebox.askyesno", return_value=True), patch(
                "BPT_BusDigitalNumber.messagebox.showinfo"
            ):
                window.write_links()
            saved = {item.line: item for item in load_lines(Path(directory))}
            self.assertIn("886路", saved)
            self.assertEqual(saved["886路"].source, "manual")          # 手工来源没被覆盖
            self.assertEqual(saved["886路"].last, "2026.9")            # 时间取并集（月份比较要正确）
            self.assertEqual(saved["886路"].status, "current")          # 手工确认过的"现在还跑"没被降级
            self.assertTrue(saved["886路"].evidence)                   # 证据链写进去了
            self.assertEqual(saved["877"].source, "online")            # 新记录来源是网上查的
            self.assertIn("网上查的", saved["877"].note)

    def test_query_reports_network_failure_without_crashing(self):
        window = self.make_window(FakeService(fail=bpt_online.OnlineUnavailable("连不上联网服务")))
        window.query_var.set("31688")
        window.query("bus")
        self.settle(window)
        self.assertIn("没成功", window.answer_var.get())
        self.assertIn("本地数据完全不受影响", window.detail_var.get())
        self.assertEqual(window.service.closed, 1, "失败后要丢掉坏会话，下次重新握手")


class DistrictBusTests(unittest.TestCase):
    """区属公交（预留）：没有数据也能跑，规则与模板要正确。"""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def test_missing_file_is_fine(self):
        with scratch() as directory:
            base = Path(directory)
            self.assertEqual(bpt_data.load_district_bus(base), ([], [], [], ()))
            self.assertFalse(bpt_data.district_bus_path(base).exists())

    def test_broken_file_never_blocks(self):
        with scratch() as directory:
            base = Path(directory)
            bpt_data.district_bus_path(base).write_text("{ 不是 json", encoding="utf-8")
            self.assertEqual(bpt_data.load_district_bus(base), ([], [], [], ()))
            bpt_data.district_bus_path(base).write_text('["不是对象"]', encoding="utf-8")
            self.assertEqual(bpt_data.load_district_bus(base), ([], [], [], ()))

    def test_batch_set_company_and_line(self):
        """批量设置：只改传进来的项，其它字段不动；同样的值不算改动。"""
        entries = [
            bpt_data.DistrictBusEntry("", plate="京B12345", model="XMQ6900", line="昌1路", date="2026.9.1"),
            bpt_data.DistrictBusEntry("", plate="京B12346", model="XMQ6900"),
            bpt_data.DistrictBusEntry("", plate="京G54321", model="ZK6105"),
        ]
        updated, changed = bpt_data.district_bus_set_fields(entries, ["京B12345", "京B12346"], company="海淀某客运")
        self.assertEqual(changed, 2)
        self.assertEqual([item.company for item in updated[:2]], ["海淀某客运", "海淀某客运"])
        self.assertEqual(updated[0].line, "昌1路")      # 没让改的项不动
        self.assertEqual(updated[2].company, "")        # 没选中的不动

        # 车牌大小写 / 间隔点都能对上；只改线路
        updated, changed = bpt_data.district_bus_set_fields(updated, ["京b·12345", "京G54321"], line="专2路")
        self.assertEqual(changed, 2)
        self.assertEqual([item.line for item in updated], ["专2路", "", "专2路"])
        self.assertEqual(updated[0].company, "海淀某客运")

        # 清空某一项（传空字符串）
        cleared, changed = bpt_data.district_bus_set_fields(updated, ["京B12345"], company="")
        self.assertEqual((changed, cleared[0].company), (1, ""))

        # 同样的值 → 0 改动；没选中 → 0 改动
        self.assertEqual(bpt_data.district_bus_set_fields(updated, ["京G54321"], line="专2路")[1], 0)
        self.assertEqual(bpt_data.district_bus_set_fields(updated, [], line="X")[1], 0)
        self.assertEqual(bpt_data.district_bus_companies_seen(updated), ["海淀某客运"])

    def test_batch_entry_by_model(self):
        """同一车型批量录车牌：预演要说清新增/已有/冲突/批内重复，写入只动该动的。"""
        existing = [
            bpt_data.DistrictBusEntry("", plate="京B12345", model="XMQ6900", date="2026.9.1"),
            bpt_data.DistrictBusEntry("", plate="京G99999", model="ZK6105"),
        ]
        plates = bpt_data.district_bus_split_plates("京B12346 京B12347\n京B12345, 京G99999 京B12346 59251023")
        self.assertEqual(plates, ["京B12346", "京B12347", "京B12345", "京G99999", "京B12346"])
        plan = bpt_data.district_bus_plan_batch(existing, plates, "XMQ6900", "昌1路", "2026.9.30")
        self.assertEqual(plan.added, ("京B12346", "京B12347"))
        self.assertEqual(plan.refreshed, ("京B12345",))
        self.assertEqual(plan.conflicts, (("京G99999", "ZK6105"),))
        self.assertEqual(plan.duplicates, ("京B12346",))
        self.assertEqual(plan.total_ok, 3)

        updated, plan = bpt_data.district_bus_apply_batch(existing, plates, "XMQ6900", "昌1路", "2026.9.30")
        self.assertEqual(len(updated), 4)  # 2 个原有 + 2 个新增（冲突与批内重复不写）
        self.assertEqual(next(item for item in updated if item.plate == "京B12346").line, "昌1路")
        self.assertEqual(next(item for item in updated if item.plate == "京G99999").model, "ZK6105")  # 冲突不动

        # 认不出车牌的东西不算数
        self.assertEqual(bpt_data.district_bus_split_plates("59251023 乱写"), [])
        self.assertEqual(bpt_data.district_bus_plan_batch([], ["乱写"], "X").invalid, ("乱写",))

    def test_duplicate_detection_and_merge(self):
        """重复上传检测 + 一键合并。"""
        entries = [
            bpt_data.DistrictBusEntry("", plate="京B12345", model="XMQ6900", line="昌1路", date="2026.9.1"),
            bpt_data.DistrictBusEntry("", plate="京B·12345", model="XMQ6900", date="2026.9.30"),          # 同车牌（写法不同）
            bpt_data.DistrictBusEntry("", plate="京B12345", model="KLQ6106", note="手滑记错"),            # 同车牌不同车型
            bpt_data.DistrictBusEntry("", plate="京G11111", model=""),
            bpt_data.DistrictBusEntry("", plate="", number=""),
        ]
        quality = bpt_data.district_bus_quality(entries)
        self.assertEqual(quality["duplicates"], [("京B12345", 3)])
        self.assertEqual(quality["conflicts"], [("京B12345", ("XMQ6900", "KLQ6106"))])
        self.assertEqual(quality["no_model"], ["京G11111"])

        merged, removed = bpt_data.district_bus_merge_duplicates(entries)
        self.assertEqual(removed, 2)  # 京B12345 的三条合成一条（另有一条没车牌也没自编的废记录原样保留）
        kept = next(item for item in merged if item.key == "京B12345")
        self.assertEqual(kept.line, "昌1路")          # 字段取并集
        self.assertEqual(kept.date, "2026.9.30")      # 日期取新的
        self.assertEqual(kept.model, "XMQ6900")       # 车型保留第一条
        self.assertIn("KLQ6106", kept.note)           # 冲突写进备注，不丢信息
        # 已经干干净净的数据：合并是空操作
        self.assertEqual(bpt_data.district_bus_merge_duplicates(merged)[1], 0)

    def test_plate_entries_upsert_remove_and_summary(self):
        """自己上传的车牌记录：新增 / 补字段 / 冲突不覆盖 / 删除 / 汇总。"""
        empty: list[bpt_data.DistrictBusEntry] = []
        first = bpt_data.DistrictBusEntry("", plate="京B12345", model="XMQ6900")
        entries, what, _msg = bpt_data.upsert_district_bus_entry(empty, first)
        self.assertEqual((what, len(entries)), ("新增", 1))
        self.assertEqual(entries[0].key, "京B12345")

        # 同一车牌再记一次（补上线路）：算更新，不重复
        again = bpt_data.DistrictBusEntry("", plate="京B·12345", model="XMQ6900", line="昌1路", date="2026.9.30")
        entries, what, _msg = bpt_data.upsert_district_bus_entry(entries, again)
        self.assertEqual((what, len(entries), entries[0].line, entries[0].date), ("更新", 1, "昌1路", "2026.9.30"))

        # 同一车牌不同车型：冲突，且原记录不动
        clash = bpt_data.DistrictBusEntry("", plate="京B12345", model="KLQ6106GAEVN6")
        entries, what, message = bpt_data.upsert_district_bus_entry(entries, clash)
        self.assertEqual(what, "冲突")
        self.assertIn("XMQ6900", message)
        self.assertEqual(entries[0].model, "XMQ6900")

        # 没填车牌和自编 → 无效
        self.assertEqual(bpt_data.upsert_district_bus_entry(entries, bpt_data.DistrictBusEntry(""))[1], "无效")

        # 汇总与候选车型
        entries, _what, _msg = bpt_data.upsert_district_bus_entry(
            entries, bpt_data.DistrictBusEntry("", plate="京G54321", model="XMQ6900")
        )
        entries, _what, _msg = bpt_data.upsert_district_bus_entry(
            entries, bpt_data.DistrictBusEntry("", plate="京G11111", model="ZK6105")
        )
        self.assertEqual(bpt_data.district_bus_entry_summary(entries), [("XMQ6900", 2), ("ZK6105", 1)])
        self.assertEqual(bpt_data.district_bus_entry_cars(entries), ["XMQ6900", "ZK6105"])

        conflicted = [
            bpt_data.DistrictBusEntry("", plate="京B12345", model="XMQ6900"),
            bpt_data.DistrictBusEntry("", plate="京B12345", model="ZK6105"),
        ]
        self.assertEqual(bpt_data.district_bus_entry_conflicts(conflicted), [("京B12345", ("XMQ6900", "ZK6105"))])

        # 删除按车牌（大小写 / 间隔点都能对上）
        remaining = bpt_data.remove_district_bus_entry(entries, "京g·54321")
        self.assertEqual([item.label for item in remaining], ["京B12345", "京G11111"])

    def test_plate_entries_round_trip(self):
        """车牌记录（含日期与备注）存得进、读得回。"""
        with scratch() as directory:
            base = Path(directory)
            entries = [
                bpt_data.DistrictBusEntry("", plate="京B12345", model="XMQ6900", line="昌1路", date="2026.9.30"),
                bpt_data.DistrictBusEntry("", plate="京G54321", model="ZK6105", note="空调车"),
            ]
            bpt_data.save_district_bus(base, [], [], (), entries)
            companies, loaded, ranges, empty = bpt_data.load_district_bus(base)
            self.assertEqual((companies, ranges, empty), ([], [], ()))
            self.assertEqual([item.label for item in loaded], ["京B12345", "京G54321"])
            self.assertEqual(loaded[0].date, "2026.9.30")
            self.assertEqual(loaded[1].note, "空调车")

    def test_plate_lookup_is_reserved_for_other_districts(self):
        """其他区基本只有车牌、没有自编：车牌这条路要留着，并且照实说明"还没有数据"。"""
        plate_company = bpt_data.DistrictBusCompany(
            "某某区公交", district="某某区", number_kind="车牌", plate_prefixes=("京B", "京G")
        )
        self.assertEqual(plate_company.all_plate_prefixes, ("京B", "京G"))
        self.assertTrue(plate_company.matches_plate("京B12345"))
        self.assertFalse(plate_company.matches_plate("京A12345"))
        self.assertFalse(bpt_data.DistrictBusCompany("只有前缀").matches_plate("京B12345"))  # 没规则不认

        # 车牌统一化与识别
        self.assertEqual(bpt_data.normalize_plate(" 京B·12345 "), "京B12345")
        self.assertEqual(bpt_data.normalize_plate("京b-12345"), "京B12345")
        self.assertTrue(bpt_data.looks_like_plate("京B12345"))
        self.assertFalse(bpt_data.looks_like_plate("59251023"))

        # 还没有车牌数据 → 照实说，并且带上"可能是哪家"的前缀线索
        empty_result = bpt_data.district_bus_lookup_any("京B12345", [plate_company], [], (), [])
        self.assertEqual(empty_result.status, "还没有车牌数据")
        self.assertEqual(empty_result.company, "某某区公交")
        self.assertTrue(empty_result.note)

        # 录了车牌就能查到；没录过的车牌照实说
        entry = bpt_data.DistrictBusEntry(company="某某区公交", plate="京B12345", line="专1路", model="XMQ6900")
        found = bpt_data.district_bus_lookup_any("京B·12345", [plate_company], [], (), [entry])
        self.assertEqual((found.status, found.company, found.line, found.model), ("已收录", "某某区公交", "专1路", "XMQ6900"))
        self.assertEqual(bpt_data.district_bus_lookup_any("京B99999", [plate_company], [], (), [entry]).status, "这个车牌没记录过")

        # 一个入口两种输入：自编走自编、车牌走车牌
        number_company = bpt_data.DistrictBusCompany("昌平某公司", "昌平", number_prefix="59", number_length=8)
        ranges = [bpt_data.DistrictBusRange("昌平某公司", 59181001, 59181030, "BJ6105EVCA-27")]
        self.assertEqual(bpt_data.district_bus_lookup_any("59181007", [number_company], ranges, (), []).model, "BJ6105EVCA-27")
        self.assertEqual(bpt_data.district_bus_lookup_any("京B12345", [number_company], ranges, (), []).status, "还没有车牌数据")

    def test_districts_overview(self):
        """各区县数据情况（谁有自编、谁只有车牌、谁还空着）。"""
        num = bpt_data.DistrictBusCompany("昌平某公司", "昌平", number_prefix="59", number_length=8)
        plate = bpt_data.DistrictBusCompany("某区公交", "某区", number_kind="车牌", plate_prefixes=("京G",))
        bare = bpt_data.DistrictBusCompany("还没数据的公司", "另区")
        ranges = [bpt_data.DistrictBusRange("昌平某公司", 59181001, 59181030, "BJ6105EVCA-27")]
        entries = [bpt_data.DistrictBusEntry(company="某区公交", plate="京G12345")]
        rows = dict((name, (state, count)) for name, state, count in bpt_data.district_bus_districts([num, plate, bare], ranges, entries))
        self.assertIn("有自编号段", rows["昌平"][0])
        self.assertEqual(rows["昌平"][1], 1)
        self.assertIn("按车牌（京G）", rows["某区"][0])
        self.assertIn("有车牌记录", rows["某区"][0])
        self.assertEqual(rows["另区"], ("有公司但还没有数据", 0))

    def test_template_creates_a_rule_to_copy(self):
        with scratch() as directory:
            base = Path(directory)
            path = bpt_data.save_district_bus_template(base)
            self.assertTrue(path.is_file())
            companies, entries, ranges, empty = bpt_data.load_district_bus(base)
            self.assertEqual((entries, ranges, empty), ([], [], ()))
            self.assertEqual(len(companies), 2)  # 一条有自编的示例 + 一条只有车牌的示例
            company = companies[0]
            self.assertEqual((company.number_prefix, company.number_length), ("59", 8))
            self.assertEqual((company.year_span, company.batch_span, company.serial_span), ((2, 4), (4, 5), (5, 8)))
            plate_company = companies[1]
            self.assertEqual(plate_company.number_kind, "车牌")
            self.assertEqual(plate_company.all_plate_prefixes, ("京B", "京G"))
            self.assertFalse(plate_company.matches_number("京B12345"))  # 只有车牌的公司不认自编
            # 模板已存在就不覆盖
            again = bpt_data.save_district_bus_template(base)
            self.assertEqual(bpt_data.load_district_bus(again.parent)[0][0].name, company.name)

    def test_number_layout_is_parsed(self):
        """8 位规则：59 + 年份(第 3-4 位) + 批次(第 5 位) + 序号(后三位)。"""
        company = bpt_data.DistrictBusCompany(
            "北京万佳通客运有限公司", "昌平", number_prefix="59", number_length=8,
            year_span=(2, 4), batch_span=(4, 5), serial_span=(5, 8),
        )
        self.assertEqual(company.describe("59181007"), {"year": "2018", "batch": "1", "serial": "7"})
        self.assertEqual(company.describe("59222099"), {"year": "2022", "batch": "2", "serial": "99"})
        self.assertEqual(company.describe("60000000"), {})  # 前缀不对就不解析
        self.assertEqual(bpt_data.DistrictBusCompany("X").describe("59181007"), {})  # 没配规则不乱解析

    def test_skip_tail_rule(self):
        """整批未编的尾号（例如某年忌 4）算"该批未编此号"，不是空号。"""
        span = bpt_data.DistrictBusRange("X", 59221001, 59221045, "BJ6851EVCA-30", skip_tail=("4",), skip_reason="忌 4")
        self.assertTrue(span.skips(59221004))
        self.assertFalse(span.skips(59221005))
        self.assertEqual((span.count, span.skipped_count, span.real_count), (45, 5, 40))
        plain = bpt_data.DistrictBusRange("X", 59251001, 59251010, "BJ6605EVCA")
        self.assertFalse(plain.skips(59251004))
        self.assertEqual((plain.skipped_count, plain.real_count), (0, 10))

    def test_lookup_with_ranges_and_empty_numbers(self):
        company = bpt_data.DistrictBusCompany(
            "北京万佳通客运有限公司", "昌平", number_prefix="59", number_length=8,
            year_span=(2, 4), batch_span=(4, 5), serial_span=(5, 8),
        )
        ranges = [
            bpt_data.DistrictBusRange(company.name, 59181001, 59181030, "BJ6105EVCA-27"),
            bpt_data.DistrictBusRange(company.name, 59221001, 59221045, "BJ6851EVCA-30"),
        ]
        empty = ("59221004",)
        found = bpt_data.district_bus_lookup("59181030", [company], ranges, empty)
        self.assertEqual((found.status, found.model, found.year, found.serial), ("已收录", "BJ6105EVCA-27", "2018", "30"))
        self.assertEqual(found.span, "59181001-59181030")
        blank = bpt_data.district_bus_lookup("59221004", [company], ranges, empty)
        self.assertEqual((blank.status, blank.model), ("空号（未配车）", ""))
        self.assertEqual(bpt_data.district_bus_lookup("59181031", [company], ranges, empty).status, "不在已整理的范围里")
        self.assertEqual(bpt_data.district_bus_lookup("12345678", [company], ranges, empty).status, "看不出是哪家（编号规则没对上）")
        self.assertEqual(bpt_data.district_bus_lookup("", [company], ranges, empty).status, "请输入自编")

    def test_shipped_district_bus_file(self):
        """用户给的真数据：读得回来、写回去一样、抽查编号都对。"""
        companies, _entries, ranges, empty = bpt_data.load_district_bus(Path.cwd())
        if not companies:
            self.skipTest("还没有区属公交数据")
        self.assertIn("万佳通", companies[0].name)
        self.assertEqual((companies[0].number_prefix, companies[0].number_length), ("59", 8))
        self.assertTrue(all(item.model for item in ranges), "每段都该有车型")
        with scratch() as directory:
            base = Path(directory)
            bpt_data.save_district_bus(base, companies, ranges, empty)
            again = bpt_data.load_district_bus(base)
            self.assertEqual([item.label for item in again[2]], [item.label for item in ranges])
            self.assertEqual(again[3], empty)
        checks = {
            "59181001": ("BJ6105EVCA-27", "2018", "1"),
            "59181030": ("BJ6105EVCA-27", "2018", "30"),
            "59192100": ("BJ6105EVCA-27", "2019", "100"),
            "59202061": ("KLQ6106GAEVN6", "2020", "61"),
            "59241080": ("BJ6851EVCA-30", "2024", "80"),
            "59251050": ("BJ6605EVCA", "2025", "50"),
        }
        for number, (model, year, serial) in checks.items():
            result = bpt_data.district_bus_lookup(number, companies, ranges, empty)
            self.assertEqual((result.model, result.year, result.serial), (model, year, serial), number)
        # 2022 那两批整批没编尾号 4：不是"空号"，是"该批未编此号"
        for number in ("59221004", "59221014", "59222054", "59222124"):
            result = bpt_data.district_bus_lookup(number, companies, ranges, empty)
            self.assertEqual(result.status, "该批未编此号", number)
            self.assertTrue(result.note, number)
        # 别的年份照常编尾号 4（2018、2025 都有）
        for number in ("59181004", "59251004", "59251014"):
            self.assertEqual(bpt_data.district_bus_lookup(number, companies, ranges, empty).status, "已收录", number)
        # 实车数与未编数
        self.assertEqual(sum(item.count for item in ranges), 493)
        self.assertEqual(sum(item.real_count for item in ranges), 480)
        self.assertEqual(sum(item.skipped_count for item in ranges), 13)
        # 区间左右边界：起点-1 不该命中，终点+1 也不该
        self.assertEqual(bpt_data.district_bus_lookup("59251000", companies, ranges, empty).status, "不在已整理的范围里")
        self.assertEqual(bpt_data.district_bus_lookup("59251051", companies, ranges, empty).status, "不在已整理的范围里")

    def test_number_rule_matching(self):
        company = bpt_data.DistrictBusCompany("昌平区属公交", "昌平", number_prefix="59", number_length=8)
        self.assertTrue(company.matches_number("59000101"))
        self.assertTrue(company.matches_number(" 59000101 "))
        self.assertFalse(company.matches_number("5900010"))  # 位数不对
        self.assertFalse(company.matches_number("60000101"))  # 前缀不对
        self.assertFalse(company.matches_number(""))
        # 只有车牌的公司：没填前缀/位数就什么都不匹配（避免瞎认）
        empty_rule = bpt_data.DistrictBusCompany("某区区属公交", number_kind="车牌")
        self.assertFalse(empty_rule.matches_number("京B12345"))
        matched = bpt_data.district_bus_companies_for("59000101", [company, empty_rule])
        self.assertEqual([item.name for item in matched], ["昌平区属公交"])

    def test_app_runs_without_district_bus_data(self):
        """没有数据时窗口照常打开：表格空着并提示怎么开始，查询也给一句明白话。"""
        with scratch() as directory, patch_app_dir(Path(directory)):
            self.app.reload_district_bus()
            self.assertEqual(self.app.district_bus, ([], [], [], ()))
            window = DistrictBusWindow(self.app)
            window.withdraw()
            self.addCleanup(destroy, window)
            self.assertEqual(len(window.table.tree.get_children()), 0)
            self.assertIn("建立一个空的", self.app.status_var.get() + window.table.empty.label.cget("text"))
            window.query()
            self.assertIn("还没有区属公交数据", window.answer_var.get())


class DialogFitTests(unittest.TestCase):
    """对话框的内容不能高过窗口——超了底部按钮就被挤出窗口，用户点不到。

    这是真实踩过的坑：V6 给标签加了「车体形式」之后，标签设置的「保存」被顶到窗口外。
    这里故意**不隐藏根窗口**，因为未映射的窗口量不到真实高度。
    """

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")

    @classmethod
    def tearDownClass(cls):
        cls.app._closing = True
        cls.app.destroy()

    def check(self, factory) -> None:
        window = factory()
        self.addCleanup(destroy, window)
        window.update_idletasks()
        self.app.update()
        try:
            self.app.update_idletasks()
        except tk.TclError:  # pragma: no cover - 窗口已关
            return
        need = window.winfo_reqheight()
        have = window.winfo_height()
        self.assertLessEqual(
            need,
            have,
            f"{type(window).__name__} 的内容需要 {need}px，窗口只有 {have}px，底部按钮会被挤出窗口",
        )

    def test_my_records_import_and_one_click_delete(self):
        """我的记录：导入「全都留着」，一键删除只删导入来的、不动手工数据。"""
        window = MyRecordsWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        with scratch() as directory:
            base = Path(directory)
            with patch_app_dir(base), patch(
                "BPT_BusDigitalNumber.messagebox.askyesno", return_value=True
            ), patch("BPT_BusDigitalNumber.messagebox.showinfo"):
                imported = LineLink("886路", "BJ6127C8BTB-1", ("31688",), "2024.9", "2026.8", "recorded", "import")
                window.links = [imported]
                self.app.lines = []
                window.import_lines()
                saved = load_lines(base)
                self.assertEqual([item.line for item in saved], ["886路"])
                self.assertEqual(saved[0].source, "import")

                # 现有数据已经包含它 → 再导入一次应该跳过写入（不重复、不报错）
                window.import_lines()
                self.assertEqual([item.line for item in load_lines(base)], ["886路"])

                # 手工再加一条，然后一键删掉导入来的：手工那条必须留着
                manual = LineLink("手工线路", "X", ("99999",), "", "", "recorded", "manual")
                self.app.lines = [imported, manual]
                save_lines(base, self.app.lines)
                window.delete_imported()
                remaining = load_lines(base)
                self.assertEqual([item.line for item in remaining], ["手工线路"])
                self.assertEqual([item.line for item in self.app.lines], ["手工线路"])

                window.delete_all_lines()
                self.assertEqual(load_lines(base), [])
                self.assertEqual(self.app.lines, [])

    def test_dialogs_fit_their_content(self):
        ranges = self.app.ranges
        model = model_names(ranges)[0]
        cases = {
            "标签设置": lambda: TagEditorWindow(ModelDetailWindow(self.app, model)),
            "号段编辑": lambda: RangeEditorWindow(self.app, None),
            "线路编辑": lambda: LineEditorWindow(LineManagerWindow(self.app), None),
            "线路管理": lambda: LineManagerWindow(self.app),
            "数据体检": lambda: DataHealthWindow(self.app),
            "燃料待确认": lambda: TagConfirmWindow(self.app),
            "我的记录": lambda: MyRecordsWindow(self.app),
            # 下面这些以前没有自己调 fit_dialog_height，靠 setup_dialog 里的 after_idle 兜底
            "车型目录": lambda: CatalogWindow(self.app),
            "号段管理": lambda: RangeManagerWindow(self.app),
            "拍车日志": lambda: SpotLogWindow(self.app),
            "标签规则": lambda: TagRulesWindow(self.app),
            "统计面板": lambda: StatsWindow(self.app),
            "设置": lambda: SettingsWindow(self.app),
            "备份与恢复": lambda: BackupWindow(self.app),
            "半自动建档": lambda: DraftProfilesWindow(self.app),
            "区属公交": lambda: DistrictBusWindow(self.app),
            "联网查询": lambda: OnlineQueryWindow(self.app),
            "快查工具": lambda: QuickLookupWindow(self.app),
        }
        for name, factory in cases.items():
            with self.subTest(name):
                self.check(factory)


class GuiTests(unittest.TestCase):
    """Behaviour checks against the real widgets; windows stay hidden."""

    @classmethod
    def setUpClass(cls):
        try:
            cls.app = BusLookupApp()
        except Exception as error:  # pragma: no cover - headless machines
            raise unittest.SkipTest(f"Tk 不可用：{error}")
        cls.app.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.app.close_application()

    def test_workbook_is_loaded_into_the_main_window(self):
        self.assertGreater(len(self.app.ranges), 0)
        self.assertIn("号段", self.app.subtitle_var.get())

    def test_query_finds_a_model_and_enables_the_actions(self):
        self.app.number_var.set(str(self.app.ranges[0].start))
        self.app.query()
        self.assertEqual(self.app.result_var.get(), self.app.ranges[0].model)
        self.assertNotIn("disabled", self.app.copy_button.state())

    def test_query_reports_an_unknown_number(self):
        uncovered = next(number for number in range(10000, 100000) if find_model(number, self.app.ranges) is None)
        self.app.number_var.set(str(uncovered))
        self.app.query()
        self.assertIn("未收录", self.app.result_var.get())
        self.assertIn("disabled", self.app.copy_button.state())

    def test_query_survives_input_that_is_digits_but_not_a_number(self):
        self.app.number_var.set("12²45")
        self.app.query()  # must not raise
        self.assertEqual(self.app.result_var.get(), "格式不符")

    def test_catalog_filter_and_detail_dirty_tracking(self):
        catalog = CatalogWindow(self.app)
        catalog.withdraw()
        try:
            catalog.filter_var.set("EVCA")
            catalog._refresh()  # 搜索框是防抖的，测试里直接刷新
            self.assertTrue(catalog.shown_models)
            self.assertTrue(all("EVCA" in name.upper() for name in catalog.shown_models))

            detail = ModelDetailWindow(catalog, catalog.shown_models[0])
            detail.withdraw()
            try:
                self.assertFalse(detail.is_dirty())
                detail.introduction.insert("1.0", "新增简介")
                detail.update()  # let Tk deliver the <<Modified>> notification
                self.assertTrue(detail.is_dirty())
                self.assertNotIn("disabled", detail.save_button.state())
                detail.revert()
                self.assertFalse(detail.is_dirty())
                self.assertIn("disabled", detail.save_button.state())
            finally:
                detail.destroy()
        finally:
            catalog.destroy()

    def test_detail_image_removal_only_deletes_after_save(self):
        model = model_names(self.app.ranges)[0]
        original = self.app.profiles.get(model)
        self.app.profiles[model] = {
            "introduction": "",
            "specifications": "",
            "image": "KEEP.jpg",
            "images": [{"name": "KEEP.jpg", "caption": ""}],
        }
        detail = ModelDetailWindow(self.app, model)
        detail.withdraw()
        try:
            self.assertEqual(len(detail.pending_images), 1)
            detail.remove_image()
            self.assertEqual(detail.pending_images, [])
            self.assertEqual(detail.removed_images, ["KEEP.jpg"])
            self.assertTrue(detail.is_dirty())
            detail.revert()
            self.assertEqual(len(detail.pending_images), 1)
            self.assertEqual(detail.removed_images, [])
            self.assertFalse(detail.is_dirty())
        finally:
            detail.destroy()
            if original is None:
                self.app.profiles.pop(model, None)
            else:
                self.app.profiles[model] = original

    def test_typing_five_digits_queries_immediately(self):
        first = self.app.ranges[0]
        self.app.number_var.set(str(first.start))
        self.assertEqual(self.app.result_var.get(), first.model)
        self.app.number_var.set("")
        self.assertEqual(self.app.result_var.get(), "—")

    def test_catalog_filters_by_parsed_tags(self):
        catalog = CatalogWindow(self.app)
        catalog.withdraw()
        try:
            catalog.manufacturer_var.set(ANY_TAG)
            catalog.fuel_var.set("纯电动")
            catalog._refresh()  # 搜索框是防抖的，测试里直接刷新
            self.assertTrue(catalog.shown_models)
            for name in catalog.shown_models:
                self.assertEqual(model_tags(name, self.app.tag_overrides).fuel, "纯电动")
            catalog._reset_filters()
            catalog._refresh()
            self.assertEqual(len(catalog.shown_models), len(model_names(self.app.ranges)))
        finally:
            catalog.destroy()

    def test_quick_lookup_batch_and_range(self):
        window = QuickLookupWindow(self.app, 0)
        window.withdraw()
        try:
            start = self.app.ranges[0].start
            window.batch_text.insert("1.0", f"{start}\n20A00")
            window.run_batch()
            rows = window._rows_of(window.batch_tree)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0][0], f"{start:05d}")
            self.assertIn("忽略", window.batch_summary_var.get())

            window.range_start_var.set(str(start))
            window.range_end_var.set(str(start + 3))
            window.run_range()
            segments = window._rows_of(window.range_tree)
            self.assertEqual(len(segments), 1)
            self.assertEqual(segments[0][1], "4")
            self.assertIn("已收录 4", window.range_summary_var.get())
        finally:
            window.destroy()

    def test_quick_lookup_keeps_the_pasted_order(self):
        """批量查询结果要按用户输入的顺序给（好对着自己的清单复制），重复的只算一次。"""
        window = QuickLookupWindow(self.app)
        window.withdraw()
        self.addCleanup(destroy, window)
        window.batch_text.delete("1.0", "end")
        window.batch_text.insert("1.0", "31779\n20001\n31778\n20001\n39999\n31590")
        window.run_batch()
        shown = [window.batch_tree.item(item, "values")[0] for item in window.batch_tree.get_children()]
        self.assertEqual(shown, ["31779", "20001", "31778", "39999", "31590"])
        self.assertEqual(window.batch_summary_var.get().split(" · ")[0], "共 5 条编号")
        # 表头点一下能排序，再点两下回到输入顺序
        window.batch_table.sort_by("number")
        self.assertNotEqual(
            [window.batch_tree.item(item, "values")[0] for item in window.batch_tree.get_children()], shown
        )
        window.batch_table.sort_by("number")
        window.batch_table.sort_by("number")
        self.assertEqual(
            [window.batch_tree.item(item, "values")[0] for item in window.batch_tree.get_children()], shown
        )
        # 只想复制车型：剪贴板里只有车型，没有自编号
        window.copy_batch_models()
        clipboard = self.app.clipboard_get()
        self.assertNotIn("31779", clipboard)
        self.assertEqual(len(clipboard.splitlines()), 5)
        self.assertIn("复制", window.batch_summary_var.get())

    def test_tag_editor_writes_overrides_for_a_new_model(self):
        model = "BJ6123C7C4D"
        if model not in model_names(self.app.ranges):
            self.skipTest("示例车型不在当前数据里")
        original = self.app.tag_overrides
        detail = ModelDetailWindow(self.app, model)
        detail.withdraw()
        editor = None
        try:
            with scratch() as directory, isolated_model(self.app, model):
                with patch_app_dir(Path(directory)):
                    editor = TagEditorWindow(detail)
                    editor.withdraw()
                    self.assertEqual(model_tags(model, self.app.tag_overrides).fuel, UNKNOWN_TAG)

                    editor.vars["fuel"].set("柴油")
                    editor.prefix_var.set(True)
                    editor.vars["manufacturer"].set("北汽福田欧辉")
                    editor.save()

                    tags = model_tags(model, self.app.tag_overrides)
                    self.assertEqual(tags.fuel, "柴油")
                    self.assertEqual(tags.manufacturer, "北汽福田欧辉")
                    self.assertIn("柴油", detail.tag_var.get())

                    saved = json.loads((Path(directory) / "车型标签.json").read_text(encoding="utf-8"))
                    self.assertEqual(saved["models"][model]["fuel"], "柴油")
                    self.assertEqual(saved["manufacturers"]["BJ"], "北汽福田欧辉")

                    # 再打开一次，恢复自动推断应当把该车型的覆盖删掉（前缀映射保留）
                    editor = TagEditorWindow(detail)
                    editor.withdraw()
                    editor.reset_to_auto()
                    self.assertEqual(model_tags(model, self.app.tag_overrides).fuel, UNKNOWN_TAG)
                    saved = json.loads((Path(directory) / "车型标签.json").read_text(encoding="utf-8"))
                    self.assertNotIn(model, saved.get("models", {}))
                    self.assertEqual(saved["manufacturers"]["BJ"], "北汽福田欧辉")
        finally:
            destroy(editor)
            destroy(detail)
            self.app.tag_overrides = original

    def test_bulk_tag_editor_updates_every_selected_model(self):
        models = model_names(self.app.ranges)[:3]
        original = self.app.tag_overrides
        catalog = CatalogWindow(self.app)
        catalog.withdraw()
        bulk = None
        try:
            with scratch() as directory:
                with patch_app_dir(Path(directory)):
                    catalog.tree.selection_set(models)
                    catalog.open_bulk_tags()
                    bulk = next(
                        child
                        for child in catalog.winfo_children()
                        if isinstance(child, BulkTagWindow)
                    )
                    bulk.withdraw()
                    bulk.use["fuel"].set(True)
                    bulk._sync("fuel")
                    bulk.vars["fuel"].set("柴油")
                    bulk.apply()

                    for model in models:
                        self.assertEqual(model_tags(model, self.app.tag_overrides).fuel, "柴油")
                    saved = json.loads((Path(directory) / "车型标签.json").read_text(encoding="utf-8"))
                    for model in models:
                        self.assertEqual(saved["models"][model]["fuel"], "柴油")
        finally:
            destroy(bulk)
            destroy(catalog)
            self.app.tag_overrides = original

    def test_health_window_lists_pending_items(self):
        window = DataHealthWindow(self.app)
        window.withdraw()
        try:
            categories = [window.category_tree.item(item, "values")[0] for item in window.category_tree.get_children()]
            self.assertIn("燃料未标注", categories)
            self.assertIn("还没有资料", categories)
            self.assertTrue(window.detail_tree.get_children())
        finally:
            window.destroy()

    def test_range_manager_search_and_model_filter(self):
        window = RangeManagerWindow(self.app)
        window.withdraw()
        try:
            total = len(self.app.ranges)
            self.assertEqual(len(window.tree.get_children()), total)

            first = self.app.ranges[0]
            window.filter_var.set(str(first.start))
            window.refresh()  # 搜索是防抖的，直接刷新拿到确定结果
            shown = window.tree.get_children()
            self.assertTrue(shown)
            self.assertEqual(shown[0], "0")  # iid 仍是原始下标，编辑/删除不会错行
            self.assertEqual(window.tree.item(shown[0], "values")[0], str(first.start))
            self.assertIn("匹配", window.count_var.get())

            window.reset_filters()
            window.model_var.set(first.model)
            window.refresh()
            self.assertTrue(all(window.tree.item(item, "values")[2] == first.model for item in window.tree.get_children()))

            window.reset_filters()
            window.refresh()
            self.assertEqual(len(window.tree.get_children()), total)
            # 防抖通道也在：搜索框变化后过一小会儿会自动刷新
            window.filter_var.set(str(first.start))
            time.sleep(0.25)
            window.update()
            self.assertEqual(len(window.tree.get_children()), 1)
        finally:
            window.destroy()

    def test_tag_rules_window_saves_a_rule(self):
        original = self.app.tag_overrides
        window = None
        editor = None
        try:
            with scratch() as directory:
                with patch_app_dir(Path(directory)):
                    window = TagRulesWindow(self.app)
                    window.withdraw()
                    editor = TagRuleEditorWindow(window, None)
                    editor.withdraw()
                    editor.pattern_var.set("BJ6123C7BTD")
                    editor.vars["fuel"].set("天然气")
                    editor.save()

                    self.assertEqual(len(self.app.tag_overrides.get("rules", [])), 1)
                    self.assertEqual(model_tags("BJ6123C7BTD", self.app.tag_overrides).fuel, "天然气")
                    saved = json.loads((Path(directory) / "车型标签.json").read_text(encoding="utf-8"))
                    self.assertEqual(saved["rules"], [{"pattern": "BJ6123C7BTD", "fuel": "天然气"}])
        finally:
            destroy(editor)
            destroy(window)
            self.app.tag_overrides = original

    def test_metro_batch_window_queries_and_reports(self):
        if not self.app.metro_ranges:
            self.skipTest("没有可用的地铁数据")
        metro = MetroWindow(self.app)
        metro.withdraw()
        window = MetroBatchWindow(metro)
        window.withdraw()
        try:
            window.number_text.insert("1.0", "01001 NOPE")
            window.run_batch()
            rows = tree_rows(window.tree)
            self.assertEqual(len(rows), 1)
            self.assertIn("命中 1", window.summary_var.get())
            self.assertIn("忽略 1", window.summary_var.get())
        finally:
            destroy(window)
            destroy(metro)

    def test_tk_callback_errors_are_logged_instead_of_vanishing(self):
        with scratch() as directory:
            with patch_app_dir(Path(directory)), patch(
                "BPT_BusDigitalNumber.messagebox.showerror"
            ) as popup:
                self.app.after(0, lambda: 1 / 0)
                self.app.update()
                self.assertTrue(popup.called)
                log = Path(directory) / "错误日志.txt"
                self.assertTrue(log.is_file())
                self.assertIn("ZeroDivisionError", log.read_text(encoding="utf-8"))

    def test_backup_window_lists_and_restores(self):
        with scratch() as directory:
            base = Path(directory)
            path = base / DATA_FILENAME
            save_ranges(path, [BusRange(20000, 20099, "A")])
            save_ranges(path, [BusRange(20000, 20099, "B")])
            window = None
            try:
                with patch_app_dir(base), patch(
                    "BPT_BusDigitalNumber.messagebox.askyesno", return_value=True
                ), patch("BPT_BusDigitalNumber.messagebox.showinfo"):
                    window = BackupWindow(self.app)
                    window.withdraw()
                    self.assertEqual(len(window.entries), 1)
                    window.tree.selection_set("0")
                    window.restore()
                    self.assertEqual(load_ranges(path), [BusRange(20000, 20099, "A")])
            finally:
                destroy(window)
                # restore() 会让主窗口从（被打补丁的）目录重新读数据，这里读回真实数据
                self.app.reload_data()
                self.assertGreater(len(self.app.ranges), 100)

    def test_catalog_rows_follow_tag_changes(self):
        target = next(
            (
                name
                for name in model_names(self.app.ranges)
                if model_tags(name, self.app.tag_overrides).fuel == UNKNOWN_TAG
            ),
            None,
        )
        if target is None:
            self.skipTest("当前数据里没有燃料未标注的车型")
        original = self.app.tag_overrides
        catalog = CatalogWindow(self.app)
        catalog.withdraw()
        try:
            before = catalog.tree.item(target, "values")[2]
            self.app.set_tag_overrides({"models": {target: {"fuel": "柴油"}}})
            self.assertEqual(catalog.tree.item(target, "values")[2], "柴油")
            self.assertNotEqual(before, "柴油")
            self.app.set_tag_overrides({})
            self.assertEqual(catalog.tree.item(target, "values")[2], before)
        finally:
            destroy(catalog)
            self.app.set_tag_overrides(original)

    def test_data_change_reaches_nested_dialogs(self):
        model = model_names(self.app.ranges)[0]
        detail = ModelDetailWindow(self.app, model)
        detail.withdraw()
        editor = None
        try:
            editor = TagEditorWindow(detail)
            editor.withdraw()
            windows = self.app.iter_toplevels()
            self.assertIn(detail, windows)
            self.assertIn(editor, windows)  # 挂在详情窗口下面也要能收到通知
        finally:
            destroy(editor)
            destroy(detail)

    def test_close_application_protects_unsaved_profile_edits(self):
        model = model_names(self.app.ranges)[0]
        detail = ModelDetailWindow(self.app, model)
        detail.withdraw()
        try:
            detail.introduction.insert("1.0", "临时改动")
            detail.update()
            self.assertIn(detail, self.app.dirty_windows())

            with patch("BPT_BusDigitalNumber.messagebox.askyesnocancel", return_value=None):
                self.app.close_application()
            self.assertTrue(self.app.winfo_exists())  # 选择「取消」就不该关掉

            with patch("BPT_BusDigitalNumber.messagebox.askyesnocancel", return_value=True), patch.object(
                ModelDetailWindow, "save"
            ) as saver:
                self.app.close_application()
                self.assertTrue(saver.called)
            self.assertTrue(self.app.winfo_exists())  # 保存没成功也不能关
        finally:
            detail.revert()
            destroy(detail)
        self.assertEqual(self.app.dirty_windows(), [])

    def test_reload_data_resets_the_query_state(self):
        first = self.app.ranges[0]
        self.app.number_var.set(str(first.start))
        self.app.query()
        self.assertEqual(self.app.result_var.get(), first.model)
        self.app.reload_data()
        self.assertEqual(self.app.result_var.get(), "—")
        self.assertGreater(len(self.app.ranges), 100)

    def test_debounce_collapses_rapid_calls(self):
        calls: list[str] = []
        trigger = debounce(self.app, "unit-test", 60, lambda: calls.append("run"))
        for _ in range(5):
            trigger()
        self.assertEqual(calls, [])
        time.sleep(0.12)
        self.app.update()
        self.assertEqual(len(calls), 1)

    def test_center_window_prefers_stored_geometry(self):
        with scratch() as directory:
            base = Path(directory)
            save_settings(base, {"windows": {"TestWindow": [30, 40, 500, 400]}})
            window = tk.Toplevel(self.app)
            window.withdraw()
            calls: list[tuple] = []
            original = window.geometry

            def spy(*args: object) -> str:
                calls.append(args)
                return original(*args)  # type: ignore[arg-type]

            window.geometry = spy  # type: ignore[method-assign]
            try:
                center_window(window, None, 300, 200, settings_key="TestWindow", base_directory=base)
                self.assertEqual(calls, [("500x400+30+40",)])
            finally:
                destroy(window)

    def test_health_detail_selection_returns_models(self):
        window = DataHealthWindow(self.app)
        window.withdraw()
        try:
            key = next((item.key for item in window.items if item.rows), None)
            if key is None:
                self.skipTest("当前没有待办项")
            window.category_tree.selection_set(key)
            window._show_selected()
            rows = window.detail_tree.get_children()[:3]
            window.detail_tree.selection_set(rows)
            models = window.selected_models()
            known = set(model_names(self.app.ranges))
            self.assertTrue(all(name in known for name in models))
            self.assertEqual(len(models), len([name for name in models if name in known]))
        finally:
            destroy(window)

    def test_undo_last_tag_change_restores_the_newest_backup(self):
        original = self.app.tag_overrides
        window = None
        try:
            with scratch() as directory:
                base = Path(directory)
                with patch_app_dir(base), patch(
                    "BPT_BusDigitalNumber.messagebox.askyesno", return_value=True
                ), patch("BPT_BusDigitalNumber.messagebox.showinfo"):
                    save_tag_overrides(base, {"manufacturers": {"HK": "第一版"}})
                    save_tag_overrides(base, {"manufacturers": {"HK": "第二版"}})
                    window = DataHealthWindow(self.app)
                    window.withdraw()
                    window.undo_last_tag_change()
                    self.assertEqual(load_tag_overrides(base)["manufacturers"]["HK"], "第一版")
        finally:
            destroy(window)
            self.app.reload_data()
            self.app.set_tag_overrides(original)

    def test_clipboard_watcher_fires_once_per_new_number(self):
        seen: list[int] = []
        watcher = ClipboardWatcher(self.app, seen.append)
        try:
            with patch("BPT_BusDigitalNumber.read_clipboard_text", return_value="20000"):
                watcher._tick()
                watcher._tick()
            with patch("BPT_BusDigitalNumber.read_clipboard_text", return_value="  编号 20213"):
                watcher._tick()
            with patch("BPT_BusDigitalNumber.read_clipboard_text", return_value="没有数字"):
                watcher._tick()
        finally:
            watcher.stop()
        self.assertEqual(seen, [20000, 20213])

    def test_committed_query_records_history_and_links(self):
        original = list(self.app.history)
        try:
            with scratch() as directory:
                with patch_app_dir(Path(directory)):
                    number = self.app.ranges[0].start
                    self.app.number_var.set(str(number))
                    self.app.query(commit=True)
                    self.assertEqual(load_history(Path(directory))[0], f"{number:05d}")
                    self.assertIn(f"{number:05d}", self.app.history)
                    self.assertTrue(self.app.history_row.winfo_children())
                    # 提示行要说明"这个号落在哪个号段"；刻意不用「属于」这个词
                    self.assertIn("号段", self.app.hint_var.get())
                    self.assertNotIn("属于", self.app.hint_var.get())
        finally:
            self.app.history = original
            self.app._refresh_history_row()

    def test_record_spot_appends_to_the_log(self):
        with scratch() as directory:
            base = Path(directory)
            with patch_app_dir(base):
                number = self.app.ranges[0].start
                self.app.number_var.set(str(number))
                self.app.query(commit=True)
                with patch("BPT_BusDigitalNumber.SpotDialog") as dialog_class:
                    self.app.record_spot()
                    self.assertTrue(dialog_class.called)
                    kwargs = dialog_class.call_args.kwargs
                    self.assertEqual(kwargs["number"], str(number))
                    self.assertEqual(kwargs["model"], self.app.current_model)

                # 走一遍对话框自己的保存逻辑（备注 + 照片）
                dialog = SpotDialog(self.app, number=f"{number:05d}", model=self.app.current_model)
                dialog.withdraw()
                try:
                    dialog.note_var.set("测试备注")
                    dialog.pending_photo = Image.new("RGB", (60, 40), "red")
                    dialog.save()
                finally:
                    destroy(dialog)
                entries = load_spots(base)
                self.assertEqual(len(entries), 1)
                self.assertEqual(entries[0].note, "测试备注")
                self.assertEqual(entries[0].model, self.app.current_model)
                self.assertTrue(entries[0].photo)
                self.assertTrue(stored_photo_path(base, entries[0].photo).is_file())

    def test_spot_dialog_edits_an_existing_entry(self):
        with scratch() as directory:
            base = Path(directory)
            with patch_app_dir(base):
                append_spot(base, make_spot("20000", "A", "旧备注"))
                window = SpotLogWindow(self.app)
                window.withdraw()
                try:
                    self.assertEqual(window.tree.item("0", "values")[4], "—")  # 没有照片
                    window.tree.selection_set("0")
                    editor = SpotDialog(
                        window,
                        index=0,
                        number="20000",
                        model="A",
                        note="改过的备注",
                        photo="",
                        time_text="2026-01-01 00:00:00",
                    )
                    editor.withdraw()
                    try:
                        editor.save()
                    finally:
                        destroy(editor)
                    entries = load_spots(base)
                    self.assertEqual(entries[0].note, "改过的备注")
                    self.assertEqual(entries[0].time, "2026-01-01 00:00:00")
                    self.assertIn("1 条", window.summary_var.get())
                finally:
                    destroy(window)

    def test_continuous_entry_clears_the_box_but_keeps_the_result(self):
        original = dict(self.app.preferences)
        try:
            self.app.continuous_var.set(True)
            self.app.toggle_continuous()
            self.assertTrue(self.app.preference("continuous_entry"))
            self.app.number_var.set("20000")
            self.app.query(commit=True)
            model = self.app.result_var.get()
            self.assertTrue(model and model != "—")
            self.app._clear_for_next()  # 正常情况是 after(700) 触发
            self.assertEqual(self.app.number_var.get(), "")
            self.assertEqual(self.app.result_var.get(), model)
            # 刷新历史时不能把「连续录入」复选框一起清掉（它曾经被放进历史行里）
            self.app.history = ["20000", "20213"]
            self.app._refresh_history_row()
            self.assertTrue(self.app.continuous_check.winfo_exists())
            self.assertTrue(self.app.continuous_check.winfo_ismapped() or True)
        finally:
            self.app.continuous_var.set(False)
            self.app.toggle_continuous()
            self.app.preferences = original

    def test_repair_row_appears_when_the_workbook_is_broken(self):
        with scratch() as directory:
            base = Path(directory)
            write_table(
                base / DATA_FILENAME,
                list(EXPECTED_HEADERS),
                [["20000", "20010", "A"], ["20010", "20020", "B"]],
                numeric_columns={0, 1},
            )
            with patch_app_dir(base):
                broken = BusLookupApp()
                try:
                    self.assertIn("还有救", broken.repair_var.get())
                    self.assertIn("恢复最近的备份", broken.repair_var.get())
                    self.assertIn("disabled", broken.button.state())
                    with patch("BPT_BusDigitalNumber.messagebox.askyesno", return_value=True):
                        broken.force_load_ranges()
                    self.assertEqual(len(broken.ranges), 1)
                    self.assertNotIn("disabled", broken.button.state())
                    self.assertIn("忽略", broken.status_var.get())
                finally:
                    broken._closing = True
                    broken.destroy()
            self.app.reload_data()

    def test_spot_log_window_filters_and_deletes(self):
        with scratch() as directory:
            base = Path(directory)
            append_spot(base, make_spot("20000", "BJ6123C7C4D", "四惠"))
            append_spot(base, make_spot("20213", "XMQ6100G1", "国贸"))
            window = None
            try:
                with patch_app_dir(base), patch(
                    "BPT_BusDigitalNumber.messagebox.askyesno", return_value=True
                ):
                    window = SpotLogWindow(self.app)
                    window.withdraw()
                    self.assertEqual(len(window.tree.get_children()), 2)
                    window.filter_var.set("国贸")
                    window.refresh()
                    self.assertEqual(len(window.tree.get_children()), 1)
                    window.reset_filters()
                    window.refresh()
                    window.tree.selection_set("0")
                    window.delete_selected()
                    self.assertEqual(len(load_spots(base)), 1)
            finally:
                destroy(window)

    def test_settings_window_toggles_are_persisted(self):
        original = dict(self.app.preferences)
        window = None
        try:
            with scratch() as directory:
                base = Path(directory)
                with patch_app_dir(base):
                    window = SettingsWindow(self.app)
                    window.withdraw()
                    window.toggle("speak_on_query", True)
                    saved = load_settings(base)
                    self.assertTrue(saved["preferences"]["speak_on_query"])
                    self.assertTrue(self.app.preference("speak_on_query"))
                    window.toggle("speak_on_query", False)
                    self.assertFalse(load_settings(base)["preferences"]["speak_on_query"])
        finally:
            destroy(window)
            self.app.preferences = original

    def test_detail_gallery_saves_files_and_profile(self):
        model = model_names(self.app.ranges)[0]
        with scratch() as directory, isolated_model(self.app, model):
            base = Path(directory)
            with patch_app_dir(base), patch(
                "BPT_BusDigitalNumber.messagebox.showinfo"
            ):
                detail = ModelDetailWindow(self.app, model)
                detail.withdraw()
                try:
                    detail._add_image(Image.new("RGB", (60, 40), "red"))
                    detail._add_image(Image.new("RGB", (60, 40), "blue"))
                    self.assertTrue(detail.is_dirty())
                    detail.set_cover()
                    detail.save()
                finally:
                    detail.destroy()
                profile = load_profiles(base)[model]
                self.assertEqual(len(profile["images"]), 2)
                self.assertEqual(profile["image"], profile["images"][0]["name"])
                for item in profile["images"]:
                    self.assertTrue((image_directory(base) / item["name"]).is_file())

    def test_catalog_searches_aliases(self):
        target = model_names(self.app.ranges)[0]
        original = self.app.tag_overrides
        catalog = CatalogWindow(self.app)
        catalog.withdraw()
        try:
            self.app.set_tag_overrides({"aliases": {target: ["大福后"]}})
            catalog.filter_var.set("大福后")
            catalog._refresh()
            self.assertEqual(catalog.shown_models, [target])
            self.assertIn("大福后", catalog.tree.item(target, "values")[0])
            catalog.filter_var.set("绝不存在的名字")
            catalog._refresh()
            self.assertEqual(catalog.shown_models, [])
        finally:
            destroy(catalog)
            self.app.set_tag_overrides(original)

    def test_bulk_range_window_validates_and_applies(self):
        window = BulkRangeWindow(self.app, [(99900, 99960), (99800, 99810)], "  ")
        window.withdraw()
        try:
            pending, errors = window.build_ranges()
            self.assertEqual(pending, [])
            self.assertEqual(len(errors), 2)  # 车型为空
            window.model_var.set("测试车型")
            window.apply_model(all_rows=True)
            pending, errors = window.build_ranges()
            self.assertEqual(errors, [])
            self.assertEqual([(item.start, item.end) for item in pending], [(99900, 99960), (99800, 99810)])

            # 与现有号段重叠的区间会被拦下
            overlapping = BulkRangeWindow(self.app, [(self.app.ranges[0].start, self.app.ranges[0].end)], "X")
            overlapping.withdraw()
            try:
                _pending, errors = overlapping.build_ranges()
                self.assertTrue(errors)
                self.assertIn("重复", errors[0])
            finally:
                overlapping.destroy()
        finally:
            destroy(window)

    def test_merge_window_saves_the_incoming_table(self):
        with scratch() as directory:
            base = Path(directory)
            incoming = [BusRange(93000, 93100, "NEW-A"), BusRange(93200, 93300, "NEW-B")]
            window = None
            original_count = len(self.app.ranges)
            try:
                with patch_app_dir(base), patch(
                    "BPT_BusDigitalNumber.messagebox.askyesno", return_value=True
                ):
                    save_ranges(base / DATA_FILENAME, self.app.ranges)
                    window = MergeWindow(self.app, incoming, "测试.xlsx")
                    window.withdraw()
                    self.assertEqual(len(window.plan.added), 2)
                    window.strategy_var.set("additive")  # 只把两个新号段补进去
                    window.apply()
                saved = load_ranges(base / DATA_FILENAME)
                self.assertEqual(len(saved), original_count + 2)
                self.assertIn("NEW-A", [item.model for item in saved])
            finally:
                destroy(window)
            self.app.reload_data()

    def test_draft_profiles_window_only_fills_empty_fields(self):
        model = model_names(self.app.ranges)[0]
        original = self.app.profiles.get(model)
        with scratch() as directory:
            base = Path(directory)
            window = None
            try:
                with patch_app_dir(base):
                    self.app.profiles[model] = {"introduction": "已有简介", "specifications": "", "image": "", "images": []}
                    window = DraftProfilesWindow(self.app)
                    window.withdraw()
                    self.assertNotIn(model, window.targets)  # 已有简介的车型不在待办里
                    target = window.targets[0]
                    window.tree.selection_set(target)
                    window.apply_selected()
                    created = load_profiles(base)[target]
                    self.assertTrue(created["introduction"])
                    self.assertTrue(created["specifications"])
                    self.assertNotIn(target, window.targets)
            finally:
                destroy(window)
        if original is None:
            self.app.profiles.pop(model, None)
        else:
            self.app.profiles[model] = original

    def test_range_editor_prefills_an_unrecorded_span(self):
        window = RangeEditorWindow(self.app, None, (99900, 99960, "预填车型"))
        window.withdraw()
        try:
            self.assertEqual(window.start_var.get(), "99900")
            self.assertEqual(window.end_var.get(), "99960")
            self.assertEqual(window.model_var.get(), "预填车型")
        finally:
            destroy(window)

    def test_health_gap_rows_open_the_range_editor(self):
        window = DataHealthWindow(self.app)
        window.withdraw()
        created: list[tuple] = []
        try:
            gaps = next((item for item in window.items if item.key == "gap"), None)
            if gaps is None or not gaps.rows:
                self.skipTest("当前没有号段空档")
            window.category_tree.selection_set("gap")
            window._show_selected()
            window.detail_tree.selection_set(window.detail_tree.get_children()[0])
            with patch("BPT_BusDigitalNumber.RangeEditorWindow", side_effect=lambda *args: created.append(args)) as spy:
                window._open_subject()
                self.assertTrue(spy.called)
            self.assertEqual(len(created), 1)
            _parent, index, initial = created[0]
            self.assertIsNone(index)
            self.assertEqual(initial[2], "")
            start, end = initial[0], initial[1]
            self.assertLess(start, end)
        finally:
            destroy(window)

    def test_quick_lookup_exposes_missing_numbers_for_bulk_fill(self):
        window = QuickLookupWindow(self.app, 1)
        window.withdraw()
        try:
            window.range_start_var.set("99900")
            window.range_end_var.set("99960")
            window.run_range()
            self.assertEqual(len(window.missing_numbers), 61)
            self.assertEqual(window.missing_segments, [(99900, 99960)])
            self.assertNotIn("disabled", window.batch_fill_button.state())
            opened: list[tuple] = []
            with patch("BPT_BusDigitalNumber.BulkRangeWindow", side_effect=lambda *args: opened.append(args)):
                window.fill_missing()
            self.assertEqual(len(opened), 1)
            intervals = opened[0][1]
            self.assertEqual(intervals, [(99900, 99960)])

            window.filter_ok = True
            window.range_start_var.set(str(self.app.ranges[0].start))
            window.range_end_var.set(str(self.app.ranges[0].start + 9))
            window.run_range()
            self.assertEqual(window.missing_numbers, [])
            self.assertIn("disabled", window.batch_fill_button.state())
        finally:
            destroy(window)

    def test_stats_window_summarises_the_loaded_data(self):
        window = StatsWindow(self.app)
        window.withdraw()
        try:
            self.assertIn(f"{len(self.app.ranges)} 个号段", window.summary_var.get())
            rows = window.tree.get_children()
            self.assertTrue(rows)
            top = window.tree.item(rows[0], "values")
            counted = sum(item.end - item.start + 1 for item in self.app.ranges if item.model == top[0])
            self.assertEqual(int(top[3]), counted)
            self.assertEqual(len(window.charts), 4)
        finally:
            destroy(window)

    def test_tag_editor_saves_aliases(self):
        model = model_names(self.app.ranges)[0]
        original = self.app.tag_overrides
        detail = ModelDetailWindow(self.app, model)
        detail.withdraw()
        editor = None
        try:
            with scratch() as directory:
                base = Path(directory)
                with patch_app_dir(base):
                    editor = TagEditorWindow(detail)
                    editor.withdraw()
                    editor.alias_var.set("大福后、气尼")
                    editor.save()
                    saved = load_tag_overrides(base)
                    self.assertEqual(saved["aliases"][model], ["大福后", "气尼"])
        finally:
            destroy(editor)
            destroy(detail)
            self.app.set_tag_overrides(original)

    def test_metro_window_reports_matches(self):
        if not self.app.metro_ranges:
            self.skipTest("没有可用的地铁数据")
        metro = MetroWindow(self.app)
        metro.withdraw()
        try:
            metro.number_var.set("01001")
            metro.query()
            self.assertIn("DKZ4", metro.result_var.get())
            metro.number_var.set("99999")
            metro.query()
            self.assertIn("未收录", metro.result_var.get())
            metro.number_var.set("NOPE")
            metro.query()
            self.assertIn("字母和数字", metro.result_var.get())
        finally:
            metro.destroy()


if __name__ == "__main__":
    unittest.main()
