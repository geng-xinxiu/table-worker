#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
table_worker.py —— 「一人数据部」表格工引擎 V1
================================================

一句话：把班委收到的"烂表"变成四件套交付物 —— 干净表 / 异常清单 / 运行日志 / 汇总报告。

用法：
    python table_worker.py <input.xlsx|input.csv> --out <输出目录> [--rules rules.json]

输出（四件套）：
    clean.xlsx    清洗后的干净表（统一格式、去重、规范化，附源行号与修正说明）
    issues.xlsx   异常清单（可筛选，按严重度着色）+ 类型统计表
    issues.md     同一份异常清单的 Markdown 版本（便于贴群/发客户）
    run.log       机器写的运行日志（时间戳、耗时到毫秒、每步命中条数）—— 参赛证据
    report.md     一页汇总报告（可直接发客户）

设计原则：
    1. 不编造：所有数字都来自真实处理结果，报告里的每个数字都能在 issues/run.log 里对上。
    2. 不越权：只做"能规则化判定"的修正；需要业务判断的一律标 error 交人工确认。
    3. 可复现：同一输入重复运行，结果一致、耗时相近（run.log 里记录每次的毫秒数）。

依赖：openpyxl（读写 xlsx）+ Python 标准库。不需要 pandas。
"""

import argparse
import csv
import json
import re
import sys
import time
import unicodedata
from datetime import datetime, date, timedelta
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:                                        # pragma: no cover
    print("缺少 openpyxl。请先安装：pip install openpyxl")
    raise SystemExit(1)

VERSION = "1.0.0"
PROCESS_START = time.perf_counter()
WALL_CLOCK = datetime.now()

# ---------------------------------------------------------------- 默认规则

DEFAULT_CONFIG = {
    "input": {"sheet": None, "header_scan_rows": 12},
    "columns": {
        "aliases": {
            "姓名": ["姓名", "名字", "学生姓名", "姓 名", "填报人", "联系人"],
            "学号": ["学号", "学 号", "学籍号", "学生学号", "studentid", "student_id"],
            "手机号": ["手机号", "手机", "电话", "联系电话", "手机号码", "联系方式", "手机号（11位）"],
            "班级": ["班级", "行政班", "所在班级", "班号"],
            "学院": ["学院", "院系", "所属学院", "学院名称"],
            "专业": ["专业", "所学专业", "专业名称"],
            "绩点": ["绩点", "gpa", "平均绩点", "学分绩点"],
            "报名时间": ["报名时间", "报名日期", "提交时间", "填报时间", "日期"],
            "备注": ["备注", "说明", "留言", "其他"],
        },
        "required": ["姓名", "手机号", "学号"],
        "phone_pattern": "^1[3-9]\\d{9}$",
        "student_id_pattern": "^[A-Za-z]?\\d{7,11}$",
        "dates": ["报名时间"],
        "numeric_ranges": {"绩点": {"min": 0, "max": 5}},
        "whitelist": {
            "班级": {"values": ["2023级1班", "2023级2班", "2024级1班", "2024级2班", "2025级1班"],
                     "message": "班级不在可用班级名单内"},
            "学院": {"values": ["计算机学院", "机械工程学院", "经济与工商管理学院", "外国语学院", "建筑城规学院"],
                     "message": "学院名称与学校标准名称不一致"},
        },
    },
    "rules": {
        "R01_duplicate_row": True,
        "R02_duplicate_student_id": True,
        "R03_duplicate_phone": True,
        "R04_phone_format": True,
        "R05_student_id_format": True,
        "R06_required_missing": True,
        "R07_name_normalize": True,
        "R08_date_normalize": True,
        "R09_column_name_normalize": True,
        "R10_numeric_range": True,
        "R11_whitelist": True,
        "R12_structure_merge": True,
    },
    "output": {"include_source_row": True, "include_fix_note": True, "max_issue_rows": 5000},
    "report": {"org": "一人数据部", "engine": "table_worker", "version": VERSION},
}

RULE_LABELS = {
    "R01_duplicate_row": "整行重复",
    "R02_duplicate_student_id": "学号重复",
    "R03_duplicate_phone": "手机号重复",
    "R04_phone_format": "手机号格式",
    "R05_student_id_format": "学号格式",
    "R06_required_missing": "必填项缺失",
    "R07_name_normalize": "姓名不规范",
    "R08_date_normalize": "日期格式",
    "R09_column_name_normalize": "列名不规范",
    "R10_numeric_range": "数值超出区间",
    "R11_whitelist": "白名单不匹配",
    "R12_structure_merge": "表头结构（合并单元格/多行表头）",
}

SEVERITY_LABEL = {"error": "错误", "warn": "警告", "info": "提示"}
SEVERITY_ORDER = {"error": 0, "warn": 1, "info": 2}


# ---------------------------------------------------------------- 小工具


def fmt_ts(moment=None):
    moment = moment or datetime.now()
    return moment.strftime("%Y-%m-%d %H:%M:%S.") + "{:03d}".format(moment.microsecond // 1000)


def elapsed_ms(since=None):
    return (time.perf_counter() - (since or PROCESS_START)) * 1000.0


def nfkc(text):
    """全角转半角、兼容字符归一（NFKC）。"""
    return unicodedata.normalize("NFKC", text)


def is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


def raw_text(value):
    """给"原始值"列用：保留原样，只做最小处理。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S").replace(" 00:00:00", "")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_text(value):
    """给"值"用：NFKC + 去首尾空白。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, date)):
        return raw_text(value)
    return nfkc(str(value)).strip()


def strip_all_spaces(text):
    """去掉所有空白（半角空格/全角空格/制表符/零宽字符）。"""
    return re.sub(r"[\s\u3000\u200b\u200c\u200d\ufeff]+", "", text)


class RunLog:
    """运行日志：每条带毫秒级时间戳与相对耗时，最后落盘 run.log。"""

    def __init__(self):
        self.entries = []

    def add(self, level, message):
        self.entries.append("[{}] {:<5} (+{:>8.1f} ms) {}".format(
            fmt_ts(), level, elapsed_ms(), message))

    def info(self, message):
        self.add("INFO", message)

    def warn(self, message):
        self.add("WARN", message)

    def error(self, message):
        self.add("ERROR", message)

    def dump(self):
        return "\n".join(self.entries)


# ---------------------------------------------------------------- 读取输入


def read_rows_xlsx(path, sheet_name, log):
    """按"展开合并单元格"的方式读出二维表，保证合并表头不丢信息。"""
    wb = load_workbook(filename=str(path), data_only=True, read_only=False)
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise ValueError("工作表「{}」不存在，可选：{}".format(sheet_name, "、".join(wb.sheetnames)))
        ws = wb[sheet_name]
        log.info("按指定名称打开工作表「{}」".format(sheet_name))
    else:
        ws = wb.worksheets[0]
        log.info("打开工作表「{}」（共 {} 个工作表）".format(ws.title, len(wb.sheetnames)))

    rows = [[cell.value for cell in row] for row in ws.iter_rows()]

    # 合并单元格：把左上角的值铺满整个合并区域（否则 pandas/openpyxl 只给第一格值）
    merged_ranges = list(ws.merged_cells.ranges)
    propagated = 0
    for rng in merged_ranges:
        anchor = rows[rng.min_row - 1][rng.min_col - 1] if rng.min_row - 1 < len(rows) else None
        if anchor is None:
            continue
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if r - 1 < len(rows) and c - 1 < len(rows[r - 1]):
                    rows[r - 1][c - 1] = anchor
                    propagated += 1
    wb.close()
    return rows, len(merged_ranges), propagated


def read_rows_csv(path, log):
    """CSV 容错读取：utf-8-sig → utf-8 → gb18030 依次尝试。"""
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            with open(str(path), "r", encoding=encoding, newline="") as handle:
                sample = handle.read(4096)
                handle.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
                except csv.Error:
                    dialect = csv.excel
                rows = [list(row) for row in csv.reader(handle, dialect)]
            log.info("CSV 读取成功（编码 {}，{} 行）".format(encoding, len(rows)))
            return rows, 0, 0
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 编码无法识别（已尝试 utf-8-sig / utf-8 / gb18030），请另存为 UTF-8 后再试。")


def detect_header_row(rows, scan_rows, log):
    """自动找表头行：取前 N 行里"互不相同的文字单元格最多"的一行。"""
    best_index, best_score = 0, -1.0
    for index, row in enumerate(rows[:max(1, scan_rows)]):
        cells = [clean_text(cell) for cell in row if not is_blank(cell)]
        if not cells:
            continue
        texty = [c for c in cells if not re.fullmatch(r"-?\d+(\.\d+)?", c)]
        distinct = len(set(cells))
        score = distinct * 2 + len(texty) - index * 0.5
        if len(texty) >= 2 and score > best_score:
            best_index, best_score = index, score
    log.info("表头识别：第 {} 行（在前 {} 行中评分最高）".format(best_index + 1, min(scan_rows, len(rows))))
    return best_index


def build_alias_map(config):
    """别名 -> 标准列名。"""
    alias_map = {}
    for canonical, aliases in config["columns"]["aliases"].items():
        for alias in [canonical] + list(aliases):
            alias_map[strip_all_spaces(nfkc(str(alias))).lower()] = canonical
    return alias_map


def build_table(rows, header_index, config, issues, log):
    """把二维表转成 列名 + 行字典，并做列名规范化（R09）。"""
    alias_map = build_alias_map(config)
    header = rows[header_index] if header_index < len(rows) else []
    raw_columns, columns, renamed = [], [], []
    seen = {}

    for position, cell in enumerate(header):
        original = clean_text(cell) or "列{}".format(position + 1)
        raw_columns.append(original)
        key = strip_all_spaces(original).lower()
        canonical = alias_map.get(key)
        if canonical is None:
            canonical = original
        else:
            renamed.append((original, canonical))
        # 重名列加序号，避免覆盖
        if canonical in seen:
            seen[canonical] += 1
            canonical = "{}_{}".format(canonical, seen[canonical])
        else:
            seen[canonical] = 1
        columns.append(canonical)

    if config["rules"].get("R09_column_name_normalize", True):
        for original, canonical in renamed:
            issues.append({
                "row": header_index + 1, "column": original,
                "kind": RULE_LABELS["R09_column_name_normalize"], "rule": "R09_column_name_normalize",
                "severity": "info", "original": original, "suggestion": canonical,
                "note": "列名已按同义词映射为标准列名，数据不受影响",
            })
        log.info("列名规范化：{} 列，命中别名 {} 列".format(len(columns), len(renamed)))

    # 数据行：表头之后、至少有一个非空单元格
    records = []
    for offset, row in enumerate(rows[header_index + 1:], start=header_index + 2):
        values = {}
        empty = True
        for position, column in enumerate(columns):
            value = row[position] if position < len(row) else None
            if not is_blank(value):
                empty = False
            values[column] = value
        if empty:
            continue
        records.append({"row": offset, "raw": values,
                        "data": {column: clean_text(values.get(column)) for column in columns}})

    return {"columns": columns, "raw_columns": raw_columns, "records": records,
            "header_index": header_index, "renamed": renamed}


# ---------------------------------------------------------------- 字段级规范化


def normalize_phone(text):
    """手机号清洗：去空格/横线/括号/+86/全角。返回 (清洗后, 是否变化)。"""
    original = text
    value = nfkc(text)
    value = re.sub(r"[\s\u3000\-—–_()（）/]", "", value)
    value = value.replace("+", "")
    if value.startswith("0086"):
        value = value[4:]
    elif value.startswith("86") and len(value) > 11:
        value = value[2:]
    return value, value != original


DATE_PATTERNS = [
    (re.compile(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$"), "ymd"),
    (re.compile(r"^(\d{4})年(\d{1,2})月(\d{1,2})日?$"), "ymd"),
    (re.compile(r"^(\d{4})(\d{2})(\d{2})$"), "ymd"),
    (re.compile(r"^(\d{4})[-/.](\d{1,2})$"), "ym"),
    (re.compile(r"^(\d{1,2})月(\d{1,2})日?$"), "md"),
    (re.compile(r"^(\d{1,2})[-/.](\d{1,2})$"), "md"),
]


def normalize_date(text, default_year=None):
    """把各种写法的日期统一成 YYYY-MM-DD。返回 (iso 或 None, 说明)。"""
    value = clean_text(text)
    if not value:
        return None, "空值"
    compact = re.sub(r"\s+", "", nfkc(value))
    for pattern, kind in DATE_PATTERNS:
        match = pattern.match(compact)
        if not match:
            continue
        try:
            if kind == "ymd":
                year, month, day = int(match.group(1)), int(match.group(2)), int(match.group(3))
            elif kind == "ym":
                year, month, day = int(match.group(1)), int(match.group(2)), 1
            else:
                year = int(default_year or WALL_CLOCK.year)
                month, day = int(match.group(1)), int(match.group(2))
            return date(year, month, day).isoformat(), "解析为 {}".format(date(year, month, day).isoformat())
        except ValueError:
            return None, "日期不合法（月/日超范围）"
    # Excel 日期序列号（1900 体系）
    if re.fullmatch(r"\d{5}(\.\d+)?", compact):
        serial = float(compact)
        if 20000 <= serial <= 80000:
            moment = datetime(1899, 12, 30) + timedelta(days=serial)
            return moment.date().isoformat(), "按 Excel 日期序列号解析"
    return None, "无法识别的日期格式"


def normalize_name(text):
    """姓名规范化：去所有空白与零宽字符。返回 (新值, 是否变化)。"""
    cleaned = strip_all_spaces(clean_text(text))
    return cleaned, cleaned != clean_text(text)


def parse_number(text):
    value = clean_text(text).replace(",", "").replace("％", "%").replace("%", "")
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


# ---------------------------------------------------------------- 主处理


def process_table(table, config, log, structure_notes):
    """按规则顺序处理，返回 (cleaned_records, issues, stats)。"""
    issues = []
    rules = config["rules"]
    columns = table["columns"]
    records = table["records"]
    disabled = [RULE_LABELS[key] for key, on in rules.items() if not on]
    if disabled:
        log.warn("以下规则已在 rules 中关闭：{}".format("、".join(disabled)))

    required = [c for c in config["columns"]["required"] if c in columns]
    phone_pattern = re.compile(config["columns"]["phone_pattern"])
    sid_pattern = re.compile(config["columns"]["student_id_pattern"])
    date_columns = [c for c in config["columns"]["dates"] if c in columns]
    numeric_ranges = config["columns"].get("numeric_ranges", {})
    whitelist = config["columns"].get("whitelist", {})

    fixed_total = 0
    per_rule_hits = {key: 0 for key in rules}

    def add_issue(record, column, kind, rule, severity, original, suggestion, note):
        issues.append({
            "row": record["row"], "column": column, "kind": kind, "rule": rule,
            "severity": severity, "original": original, "suggestion": suggestion, "note": note,
        })
        per_rule_hits[rule] = per_rule_hits.get(rule, 0) + 1

    # ---- 逐行字段级规则
    for record in records:
        data = record["data"]
        original = record["raw_data_snapshot"] = {c: raw_text(record["raw"].get(c)) for c in columns}
        fixes = record["fixes"] = []

        # R06 必填项缺失
        if rules.get("R06_required_missing", True):
            for column in required:
                if strip_all_spaces(data.get(column, "")) == "":
                    add_issue(record, column, RULE_LABELS["R06_required_missing"], "R06_required_missing",
                              "error", "", "", "该列是必填项，本行为空，需人工补录或退回重填")

        # R07 姓名规范化
        if rules.get("R07_name_normalize", True) and data.get("姓名"):
            new_name, changed = normalize_name(data["姓名"])
            if changed:
                add_issue(record, "姓名", RULE_LABELS["R07_name_normalize"], "R07_name_normalize",
                          "warn", original["姓名"], new_name, "姓名含空格/全角空格/零宽字符，已自动去除")
                data["姓名"] = new_name
                fixes.append("姓名去空格")
                fixed_total += 1

        # R04 手机号格式
        if rules.get("R04_phone_format", True) and data.get("手机号"):
            cleaned, changed = normalize_phone(data["手机号"])
            if phone_pattern.match(cleaned):
                if changed:
                    add_issue(record, "手机号", RULE_LABELS["R04_phone_format"], "R04_phone_format",
                              "warn", original["手机号"], cleaned, "去除空格/横线/国家码后格式正确，已自动修正")
                    data["手机号"] = cleaned
                    fixes.append("手机号去分隔符")
                    fixed_total += 1
            else:
                hint = ""
                digits = re.sub(r"\D", "", cleaned)
                if len(digits) == 10 and digits.startswith("1"):
                    hint = "疑似少一位（当前 {} 位）".format(len(digits))
                elif len(digits) == 12:
                    hint = "疑似多了 1 位（当前 {} 位）".format(len(digits))
                elif not digits:
                    hint = "没有数字"
                add_issue(record, "手机号", RULE_LABELS["R04_phone_format"], "R04_phone_format",
                          "error", original["手机号"], "", "不符合大陆手机号规则（1[3-9] 开头的 11 位数字）"
                          + ("，" + hint if hint else "") + "，需向本人核实")

        # R05 学号格式
        if rules.get("R05_student_id_format", True) and data.get("学号"):
            sid = strip_all_spaces(data["学号"])
            if sid != data["学号"]:
                add_issue(record, "学号", RULE_LABELS["R05_student_id_format"], "R05_student_id_format",
                          "warn", original["学号"], sid, "学号含空格，已自动去除")
                data["学号"] = sid
                fixes.append("学号去空格")
                fixed_total += 1
            if not sid_pattern.match(sid):
                add_issue(record, "学号", RULE_LABELS["R05_student_id_format"], "R05_student_id_format",
                          "error", original["学号"], "", "学号长度/格式异常（期望：可选 1 位字母 + 7~11 位数字），需人工核实")

        # R08 日期统一
        if rules.get("R08_date_normalize", True):
            for column in date_columns:
                if not data.get(column):
                    continue
                iso, note = normalize_date(data[column])
                if iso is None:
                    add_issue(record, column, RULE_LABELS["R08_date_normalize"], "R08_date_normalize",
                              "error", original[column], "", "{}，需人工确认".format(note))
                elif iso != clean_text(data[column]):
                    add_issue(record, column, RULE_LABELS["R08_date_normalize"], "R08_date_normalize",
                              "warn", original[column], iso, "已统一为 YYYY-MM-DD（{}）".format(note))
                    data[column] = iso
                    fixes.append("{}统一格式".format(column))
                    fixed_total += 1

        # R10 数值区间
        if rules.get("R10_numeric_range", True):
            for column, spec in numeric_ranges.items():
                if column not in columns or not data.get(column):
                    continue
                number = parse_number(data[column])
                if number is None:
                    add_issue(record, column, RULE_LABELS["R10_numeric_range"], "R10_numeric_range",
                              "error", original[column], "", "不是数字，无法参与统计，需人工确认")
                    continue
                low, high = spec.get("min"), spec.get("max")
                if (low is not None and number < low) or (high is not None and number > high):
                    add_issue(record, column, RULE_LABELS["R10_numeric_range"], "R10_numeric_range",
                              "error", original[column], "", "数值 {} 超出合理区间 [{}, {}]，疑似填错".format(
                                  clean_text(data[column]), low, high))

        # R11 白名单
        if rules.get("R11_whitelist", True):
            for column, spec in whitelist.items():
                if column not in columns or not data.get(column):
                    continue
                allowed = [strip_all_spaces(clean_text(v)) for v in spec.get("values", [])]
                if strip_all_spaces(data[column]) not in allowed:
                    add_issue(record, column, RULE_LABELS["R11_whitelist"], "R11_whitelist",
                              "warn", original[column], "",
                              spec.get("message", "不在白名单内") + "；可用值：" + "、".join(allowed[:6]))

    # ---- R01 整行重复（修正后再判重，避免"138 0000 0001"与"13800000001"漏判）
    ignore_for_dup = {"备注"}
    kept_records = []
    duplicate_index = {}
    if rules.get("R01_duplicate_row", True):
        for record in records:
            key = tuple(strip_all_spaces(record["data"].get(c, "")) for c in columns if c not in ignore_for_dup)
            if key in duplicate_index:
                first = duplicate_index[key]
                add_issue(record, "整行", RULE_LABELS["R01_duplicate_row"], "R01_duplicate_row",
                          "warn", "（与第 {} 行完全相同）".format(first), "删除本行",
                          "判定为重复提交，已从 clean.xlsx 剔除")
                record["dropped"] = True
            else:
                duplicate_index[key] = record["row"]
                kept_records.append(record)
    else:
        kept_records = list(records)

    # ---- R02 / R03 关键字段重复（保留首条，其余标记为需人工确认，不自动删）
    for rule_key, column, label in (("R02_duplicate_student_id", "学号", "学号重复"),
                                    ("R03_duplicate_phone", "手机号", "手机号重复")):
        if not rules.get(rule_key, True) or column not in columns:
            continue
        first_seen = {}
        for record in kept_records:
            value = strip_all_spaces(record["data"].get(column, ""))
            if not value:
                continue
            if value in first_seen:
                add_issue(record, column, RULE_LABELS[rule_key], rule_key,
                          "error", value, "", "与第 {} 行「{}」重复，疑似同一人重复提交，需人工确认保留哪一条".format(
                              first_seen[value], value))
            else:
                first_seen[value] = record["row"]

    # ---- 排序：严重度 → 行号 → 列名
    issues.sort(key=lambda item: (SEVERITY_ORDER.get(item["severity"], 9), item["row"], item["column"]))

    # ---- 统计
    by_kind = {}
    by_severity = {"error": 0, "warn": 0, "info": 0}
    for item in issues:
        by_kind[item["kind"]] = by_kind.get(item["kind"], 0) + 1
        by_severity[item["severity"]] = by_severity.get(item["severity"], 0) + 1

    error_rows = {item["row"] for item in issues if item["severity"] == "error"}
    stats = {
        "input_rows": len(records),
        "clean_rows": len(kept_records),
        "dropped_rows": len(records) - len(kept_records),
        "issue_total": len(issues),
        "fixed_total": fixed_total,
        "manual_rows": len(error_rows),
        "by_kind": by_kind,
        "by_severity": by_severity,
        "per_rule": per_rule_hits,
        "columns": len(columns),
        "structure_notes": structure_notes,
    }
    return kept_records, issues, stats


# ---------------------------------------------------------------- 输出：xlsx


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
SEVERITY_FILL = {
    "error": PatternFill("solid", fgColor="FFC7CE"),
    "warn": PatternFill("solid", fgColor="FFEB9C"),
    "info": PatternFill("solid", fgColor="DDEBF7"),
}


def style_sheet(ws, headers, widths=None):
    for position, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=position, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        letter = get_column_letter(position)
        width = (widths or {}).get(header)
        if width is None:
            width = max(10, min(42, len(str(header)) * 2 + 4))
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_clean_xlsx(path, columns, records, config):
    wb = Workbook()
    ws = wb.active
    ws.title = "清洗结果"
    headers = list(columns)
    if config["output"].get("include_source_row", True):
        headers.append("_源行号")
    if config["output"].get("include_fix_note", True):
        headers.append("_修正说明")

    for position, header in enumerate(headers, start=1):
        ws.cell(row=1, column=position, value=header)

    for index, record in enumerate(records, start=2):
        for position, column in enumerate(columns, start=1):
            ws.cell(row=index, column=position, value=record["data"].get(column, ""))
        position = len(columns) + 1
        if config["output"].get("include_source_row", True):
            ws.cell(row=index, column=position, value=record["row"])
            position += 1
        if config["output"].get("include_fix_note", True):
            ws.cell(row=index, column=position, value="；".join(record["fixes"]) or "")

    style_sheet(ws, headers)
    wb.save(str(path))
    return len(records), len(headers)


def write_issues_xlsx(path, issues, stats):
    wb = Workbook()
    ws = wb.active
    ws.title = "异常清单"
    headers = ["严重度", "行号", "列名", "问题类型", "原始值", "建议修正值", "说明", "规则编号"]
    for position, header in enumerate(headers, start=1):
        ws.cell(row=1, column=position, value=header)

    widths = {"严重度": 8, "行号": 7, "列名": 12, "问题类型": 16,
              "原始值": 26, "建议修正值": 26, "说明": 52, "规则编号": 24}
    for index, item in enumerate(issues, start=2):
        values = [SEVERITY_LABEL.get(item["severity"], item["severity"]), item["row"],
                  item["column"], item["kind"], item["original"], item["suggestion"],
                  item["note"], item["rule"]]
        for position, value in enumerate(values, start=1):
            cell = ws.cell(row=index, column=position, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=(position in (5, 6, 7)))
            if position == 1:
                cell.fill = SEVERITY_FILL.get(item["severity"], PatternFill())
    style_sheet(ws, headers, widths)

    # 第二张表：类型统计
    ws2 = wb.create_sheet("类型统计")
    headers2 = ["问题类型", "条数", "其中需人工确认（错误）"]
    for position, header in enumerate(headers2, start=1):
        ws2.cell(row=1, column=position, value=header)
    error_by_kind = {}
    for item in issues:
        if item["severity"] == "error":
            error_by_kind[item["kind"]] = error_by_kind.get(item["kind"], 0) + 1
    row_index = 2
    for kind, count in sorted(stats["by_kind"].items(), key=lambda kv: (-kv[1], kv[0])):
        ws2.cell(row=row_index, column=1, value=kind)
        ws2.cell(row=row_index, column=2, value=count)
        ws2.cell(row=row_index, column=3, value=error_by_kind.get(kind, 0))
        row_index += 1
    ws2.cell(row=row_index, column=1, value="合计")
    ws2.cell(row=row_index, column=2, value=stats["issue_total"])
    ws2.cell(row=row_index, column=3, value=stats["by_severity"]["error"])
    style_sheet(ws2, headers2, {"问题类型": 20, "条数": 10, "其中需人工确认（错误）": 24})
    wb.save(str(path))


def write_issues_md(path, issues, stats, input_name, max_rows):
    lines = []
    lines.append("# 异常清单：{}".format(input_name))
    lines.append("")
    lines.append("- 生成时间：{}".format(fmt_ts()))
    lines.append("- 问题总数：**{} 条**（错误 {} / 警告 {} / 提示 {}）".format(
        stats["issue_total"], stats["by_severity"]["error"],
        stats["by_severity"]["warn"], stats["by_severity"]["info"]))
    lines.append("- 需人工确认的行：**{} 行**（含错误的行）".format(stats["manual_rows"]))
    lines.append("")
    lines.append("> 严重度含义：错误＝必须人工处理，否则数据不可用；警告＝已自动修正，请抽查确认；提示＝结构性说明，无需处理。")
    lines.append("")
    lines.append("## 一、按类型统计")
    lines.append("")
    lines.append("| 问题类型 | 条数 |")
    lines.append("| --- | ---: |")
    for kind, count in sorted(stats["by_kind"].items(), key=lambda kv: (-kv[1], kv[0])):
        lines.append("| {} | {} |".format(kind, count))
    lines.append("| **合计** | **{}** |".format(stats["issue_total"]))
    lines.append("")
    lines.append("## 二、明细（前 {} 条，完整见 issues.xlsx）".format(min(max_rows, len(issues))))
    lines.append("")
    lines.append("| 严重度 | 行号 | 列名 | 问题类型 | 原始值 | 建议修正值 | 说明 |")
    lines.append("| --- | ---: | --- | --- | --- | --- | --- |")
    for item in issues[:max_rows]:
        lines.append("| {} | {} | {} | {} | {} | {} | {} |".format(
            SEVERITY_LABEL.get(item["severity"], item["severity"]), item["row"], item["column"],
            item["kind"], md_escape(item["original"]), md_escape(item["suggestion"]), md_escape(item["note"])))
    if len(issues) > max_rows:
        lines.append("")
        lines.append("（其余 {} 条见 issues.xlsx）".format(len(issues) - max_rows))
    lines.append("")
    Path(path).write_text("\n".join(lines), encoding="utf-8")


def md_escape(text):
    return str(text).replace("|", "\\|").replace("\n", " ")


# ---------------------------------------------------------------- 输出：日志与报告


def write_run_log(path, log, context):
    header = [
        "# table_worker 运行日志",
        "# 引擎：{} v{}".format(context["engine"], context["version"]),
        "# 输入：{}".format(context["input_path"]),
        "# 输出目录：{}".format(context["out_dir"]),
        "# 开始时间：{}".format(context["started_at"]),
        "# 结束时间：{}".format(context["finished_at"]),
        "# Python：{}".format(context["python"]),
        "# {}-".format("-" * 62),
    ]
    footer = [
        "# {}-".format("-" * 62),
        "[{}] INFO  (+{:>8.1f} ms) 处理完成：原始 {} 行 → 干净 {} 行，问题 {} 条，总耗时 {:.1f} ms".format(
            fmt_ts(), context["elapsed_ms"], context["stats"]["input_rows"],
            context["stats"]["clean_rows"], context["stats"]["issue_total"], context["elapsed_ms"]),
    ]
    Path(path).write_text("\n".join(header + log.entries + footer), encoding="utf-8")


def write_report(path, context):
    stats = context["stats"]
    config = context["config"]
    lines = []
    lines.append("# 表格处理报告：{}".format(Path(context["input_path"]).name))
    lines.append("")
    lines.append("- 处理方：{}（{} v{}）".format(
        config["report"].get("org", "一人数据部"), config["report"].get("engine", "table_worker"),
        config["report"].get("version", VERSION)))
    lines.append("- 处理时间：{}".format(context["finished_at"]))
    lines.append("- 输入文件：`{}`（{}）".format(
        Path(context["input_path"]).name, human_size(context["input_size"])))
    lines.append("- 本报告全部数字来自本次真实运行，可在 run.log 与 issues.xlsx 中逐条核对。")
    lines.append("- **样例数据说明**：姓名均为化名、手机号为虚拟号，不含任何真实个人信息。"
                 if context.get("sample") else "")
    lines.append("")
    lines.append("## 一、结果概览")
    lines.append("")
    lines.append("| 指标 | 数值 | 说明 |")
    lines.append("| --- | ---: | --- |")
    lines.append("| 原始数据行 | {} 行 | 表头之下、至少一个单元格非空的行 |".format(stats["input_rows"]))
    lines.append("| 干净数据行 | {} 行 | 写入 clean.xlsx 的行 |".format(stats["clean_rows"]))
    lines.append("| 剔除重复行 | {} 行 | 整行内容完全重复（自动剔除，保留首条） |".format(stats["dropped_rows"]))
    lines.append("| 自动修正 | {} 处 | 姓名空格、手机号分隔符、日期格式等可规则化修正 |".format(stats["fixed_total"]))
    lines.append("| 需人工确认 | {} 行 | 含错误级问题，不能自动修，必须人工核实 |".format(stats["manual_rows"]))
    lines.append("| 问题总数 | {} 条 | 见 issues.xlsx / issues.md |".format(stats["issue_total"]))
    lines.append("| 处理耗时 | {:.0f} ms | 机器实际耗时（含读写文件） |".format(context["elapsed_ms"]))
    lines.append("")
    lines.append("## 二、异常类型分布")
    lines.append("")
    if stats["by_kind"]:
        lines.append("| 问题类型 | 条数 |")
        lines.append("| --- | ---: |")
        for kind, count in sorted(stats["by_kind"].items(), key=lambda kv: (-kv[1], kv[0])):
            lines.append("| {} | {} |".format(kind, count))
    else:
        lines.append("本次未发现任何异常。")
    lines.append("")
    lines.append("## 三、规则命中情况（每条规定是否生效）")
    lines.append("")
    lines.append("| 规则 | 说明 | 状态 | 命中条数 |")
    lines.append("| --- | --- | --- | ---: |")
    for key, label in RULE_LABELS.items():
        enabled = config["rules"].get(key, False)
        hits = stats["per_rule"].get(key, 0)
        if key == "R12_structure_merge":
            hits = context.get("structure_hits", 0)
        lines.append("| {} | {} | {} | {} |".format(
            key, label, "启用" if enabled else "已关闭", hits))
    lines.append("")
    lines.append("## 四、结构处理说明（合并单元格 / 多行表头）")
    lines.append("")
    for note in stats["structure_notes"]:
        lines.append("- {}".format(note))
    lines.append("")
    lines.append("## 五、交付物清单")
    lines.append("")
    lines.append("| 文件 | 用途 |")
    lines.append("| --- | --- |")
    lines.append("| clean.xlsx | 清洗后的干净表，可直接用于报名统计/评奖材料 |")
    lines.append("| issues.xlsx | 异常清单（可筛选、按严重度着色）+ 类型统计 |")
    lines.append("| issues.md | 异常清单的文字版，便于发群/发给填表人 |")
    lines.append("| run.log | 机器运行日志（时间戳、耗时到毫秒、每步命中条数） |")
    lines.append("| report.md | 本报告 |")
    lines.append("")
    lines.append("## 六、下一步建议")
    lines.append("")
    lines.append("1. 先处理 **错误级**（{} 条，涉及 {} 行）：这些是数据不可用的原因，交给填表人核实。".format(
        stats["by_severity"]["error"], stats["manual_rows"]))
    lines.append("2. 抽查 **警告级**（{} 条）：系统已自动修正，确认修正方向正确即可。".format(
        stats["by_severity"]["warn"]))
    lines.append("3. 错误清零后，再跑一次本引擎，得到最终可用于统计的干净表。")
    lines.append("4. 需要调整规则（必填列、白名单、日期列）时，改 rules.json 后重跑，无需改代码。")
    lines.append("")
    lines.append("## 七、如何复现")
    lines.append("")
    lines.append("```")
    lines.append("python table_worker.py \"{}\" --out \"{}\"".format(
        Path(context["input_path"]).name, context["out_dir"]))
    lines.append("```")
    lines.append("")
    lines.append("同一输入重复运行，输出的行数与问题条数完全一致；耗时略有浮动（机器负载）。")
    lines.append("")
    Path(path).write_text("\n".join(lines), encoding="utf-8")


def human_size(num_bytes):
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return "{:.1f} {}".format(size, unit) if unit != "B" else "{:.0f} B".format(size)
        size /= 1024.0
    return "{:.1f} GB".format(size)


# ---------------------------------------------------------------- 入口


def load_config(path, log):
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    if not path:
        log.info("未指定 --rules，使用内置默认规则")
        return config
    rules_path = Path(path)
    if not rules_path.exists():
        raise FileNotFoundError("找不到规则文件：{}".format(rules_path))
    user = json.loads(rules_path.read_text(encoding="utf-8"))
    for section in ("input", "columns", "rules", "output", "report"):
        if isinstance(user.get(section), dict):
            config[section].update(user[section])
    log.info("已载入规则文件：{}（自定义节：{}）".format(
        rules_path.name, "、".join(k for k in user.keys() if not k.startswith("_"))))
    return config


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="表格工引擎 V1：把烂表清洗成 干净表 / 异常清单 / 运行日志 / 汇总报告 四件套",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python table_worker.py samples\\dirty_报名表.xlsx --out samples\\out1\n"
            "  python table_worker.py 客户表.xlsx --out 结果 --rules rules.json\n"
        ),
    )
    parser.add_argument("input", help="输入文件（.xlsx 或 .csv）")
    parser.add_argument("--out", required=True, help="输出目录（不存在会自动创建）")
    parser.add_argument("--rules", help="规则文件 rules.json（不填用内置默认规则）")
    parser.add_argument("--sheet", help="指定工作表名（默认第一个工作表）")
    parser.add_argument("--max-issues-in-md", type=int, default=80, help="issues.md 里最多列多少条明细")
    args = parser.parse_args(argv)

    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    log = RunLog()
    log.info("引擎启动 {} v{}，Python {}".format(
        DEFAULT_CONFIG["report"]["engine"], VERSION, sys.version.split()[0]))

    input_path = Path(args.input)
    if not input_path.exists():
        print("错误：找不到输入文件：{}".format(input_path))
        return 2
    out_dir = Path(args.out)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        print("错误：无法创建输出目录 {} —— {}".format(out_dir, error))
        return 2

    try:
        config = load_config(args.rules, log)
        if args.sheet:
            config["input"]["sheet"] = args.sheet
        log.info("输入文件：{}（{}）".format(input_path.name, human_size(input_path.stat().st_size)))

        # 1) 读取
        step_start = time.perf_counter()
        suffix = input_path.suffix.lower()
        if suffix in (".xlsx", ".xlsm"):
            rows, merged_count, propagated = read_rows_xlsx(input_path, config["input"].get("sheet"), log)
        elif suffix in (".csv", ".txt"):
            rows, merged_count, propagated = read_rows_csv(input_path, log)
        else:
            print("错误：只支持 .xlsx / .csv，当前是 {}".format(suffix))
            return 2
        log.info("读取完成：{} 行 × 最多 {} 列，耗时 {:.1f} ms".format(
            len(rows), max((len(r) for r in rows), default=0), elapsed_ms(step_start)))

        # 2) 找表头
        header_index = detect_header_row(rows, config["input"].get("header_scan_rows", 12), log)
        structure_notes = []
        if merged_count:
            structure_notes.append(
                "检测到 {} 处合并单元格，已把左上角的值展开到整个区域（共填充 {} 个格子），避免多行表头丢列名。".format(
                    merged_count, propagated))
        if header_index > 0:
            title_cells = [clean_text(c) for c in rows[0] if not is_blank(c)]
            structure_notes.append(
                "表头不在第 1 行：第 1~{} 行被判为标题/多行表头，已跳过。第 1 行内容：{}。".format(
                    header_index, "、".join(sorted(set(title_cells))[:3]) or "（空）"))
        else:
            structure_notes.append("表头位于第 1 行，结构正常。")
        log.info("结构处理：{}".format(" / ".join(structure_notes)))

        # 3) 建表 + 列名规范化
        issues = []
        table = build_table(rows, header_index, config, issues, log)

        # 4) 规则处理
        step_start = time.perf_counter()
        records, rule_issues, stats = process_table(table, config, log, structure_notes)
        issues.extend(rule_issues)
        issues.sort(key=lambda item: (SEVERITY_ORDER.get(item["severity"], 9), item["row"], item["column"]))
        log.info("规则执行完成：命中 {} 条，自动修正 {} 处，耗时 {:.1f} ms".format(
            stats["issue_total"], stats["fixed_total"], elapsed_ms(step_start)))
        for key, label in RULE_LABELS.items():
            if not config["rules"].get(key, False):
                continue
            hits = stats["per_rule"].get(key, 0) if key != "R12_structure_merge" else (merged_count or (1 if header_index > 0 else 0))
            log.info("规则 {} {} → 命中 {} 条".format(key, label, hits))

        # 5) 写四件套
        step_start = time.perf_counter()
        write_clean_xlsx(out_dir / "clean.xlsx", table["columns"], records, config)
        write_issues_xlsx(out_dir / "issues.xlsx", issues, stats)
        write_issues_md(out_dir / "issues.md", issues, stats, input_path.name, args.max_issues_in_md)
        log.info("写出 clean.xlsx（{} 行）/ issues.xlsx（{} 条） / issues.md，耗时 {:.1f} ms".format(
            stats["clean_rows"], stats["issue_total"], elapsed_ms(step_start)))

        context = {
            "engine": DEFAULT_CONFIG["report"]["engine"],
            "version": VERSION,
            "input_path": str(input_path.resolve()),
            "input_size": input_path.stat().st_size,
            "out_dir": str(out_dir.resolve()),
            "started_at": fmt_ts(WALL_CLOCK),
            "finished_at": fmt_ts(),
            "python": sys.version.split()[0],
            "stats": stats,
            "config": config,
            "structure_hits": merged_count or (1 if header_index > 0 else 0),
            "elapsed_ms": elapsed_ms(),
            "sample": "samples" in str(input_path).lower() or "dirty" in input_path.name.lower(),
        }
        write_report(out_dir / "report.md", context)
        write_run_log(out_dir / "run.log", log, context)

        print("=" * 72)
        print("表格工引擎 v{} 处理完成".format(VERSION))
        print("输入：{}".format(input_path))
        print("输出目录：{}".format(out_dir.resolve()))
        print("-" * 72)
        print("原始 {} 行 → 干净 {} 行；剔除重复 {} 行；自动修正 {} 处；需人工确认 {} 行".format(
            stats["input_rows"], stats["clean_rows"], stats["dropped_rows"],
            stats["fixed_total"], stats["manual_rows"]))
        print("问题总数 {} 条（错误 {} / 警告 {} / 提示 {}），总耗时 {:.1f} ms".format(
            stats["issue_total"], stats["by_severity"]["error"],
            stats["by_severity"]["warn"], stats["by_severity"]["info"], context["elapsed_ms"]))
        print("-" * 72)
        for name in ("clean.xlsx", "issues.xlsx", "issues.md", "run.log", "report.md"):
            print("  [产出] {}".format(out_dir / name))
        print("=" * 72)
        return 0
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as error:
        print("错误：{}".format(error))
        return 3
    except Exception as error:                              # 兜底：把意外也记进日志
        log.error("未预期异常：{}: {}".format(type(error).__name__, error))
        try:
            (out_dir / "run.log").write_text(log.dump(), encoding="utf-8")
        except OSError:
            pass
        print("错误：处理中断 —— {}: {}".format(type(error).__name__, error))
        return 4


if __name__ == "__main__":
    sys.exit(main())
