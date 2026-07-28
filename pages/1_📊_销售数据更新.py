from __future__ import annotations

import io
import json
import re
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# =========================
# 页面配置
# =========================
st.set_page_config(
    page_title="汇发销售自动更新系统 V3",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 只美化自定义区域，不覆盖 Streamlit 上传器、下拉框等内部组件。
st.markdown(
    """
    <style>
    .block-container {
        max-width: 1380px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }
    .hero,
    .hero-title,
    .hero-subtitle,
    .section-title,
    .rule-card,
    .upload-title,
    .review-name,
    .review-reason,
    .footer-note {
        font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif;
    }

    .hero {
        text-align: center;
        padding: 18px 16px 14px 16px;
        border-bottom: 1px solid #e7ebf0;
        margin-bottom: 18px;
    }
    .hero-title {
        font-size: 25px;
        font-weight: 750;
        color: #16243a;
        line-height: 1.5;
    }
    .hero-subtitle {
        font-size: 16px;
        color: #6b7280;
        margin-top: 8px;
    }
    .section-title {
        text-align: center;
        font-size: 19px;
        font-weight: 700;
        color: #1f2937;
        margin: 25px 0 15px 0;
    }
    .rule-card {
        border: 1px solid #dfe5ec;
        border-radius: 14px;
        padding: 18px 22px;
        background: #fbfcfe;
        margin-bottom: 22px;
        font-size: 16px;
        line-height: 1.9;
        color: #263247;
    }
    .upload-title {
        text-align: center;
        font-size: 16px;
        font-weight: 700;
        margin: 8px 0 7px 0;
        color: #28354a;
    }
    .review-name {
        text-align: center;
        font-size: 17px;
        font-weight: 750;
        color: #172033;
        margin-bottom: 4px;
    }
    .review-reason {
        text-align: center;
        font-size: 15px;
        color: #6b7280;
        margin-bottom: 12px;
    }
    .footer-note {
        text-align: center;
        color: #7b8493;
        font-size: 14px;
        margin-top: 25px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# 固定配置
# =========================
SUMMARY_SHEET = "汇总"
CHECK_SHEET = "更新检查"
MERGE_LOG_SHEET = "Merge_Log"
AUXILIARY_SHEETS = {"数据口径", "匹配检查", "使用说明", "客户主数据", "客户别名", "客户明细", MERGE_LOG_SHEET}
ALIAS_FILE = Path(__file__).with_name("customer_alias.json")

BUSINESS_PREFIXES = [
    "参加活动的", "参加活动", "活动客户", "活动的", "活动",
    "赠品客户", "赠送客户", "赠品", "赠送",
    "临时客户", "新客户", "客户",
]

ACTIVITY_WORDS = ["参加活动的", "参加活动", "活动客户", "活动的", "活动"]
SERVICE_STATION_PREFIX_RE = re.compile(r"^[Ff](?=[\u4e00-\u9fff])")

SPECIAL_TOWN_DEFAULTS = {
    "参加活动的韩冰": "农安镇",
}


# =========================
# 数据模型
# =========================
@dataclass
class LedgerValue:
    opening: float = 0.0
    receivable: float = 0.0
    received: float = 0.0
    ending: float = 0.0


@dataclass
class CustomerRecord:
    key: str
    name: str
    town: str
    row: int
    core: str
    aliases: Set[str]


@dataclass
class MatchDecision:
    source_key: str
    source_name: str
    status: str  # exact / alias / saved_new / ignored / town_suggest / ambiguous / new
    target_key: Optional[str] = None
    suggested_town: Optional[str] = None
    candidates: Optional[List[str]] = None
    reason: str = ""


# =========================
# 通用工具
# =========================
def clean_name(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("【", "[").replace("】", "]")
    text = re.sub(r"[\s\u3000]+", "", text)
    return text


def display_name(value: object) -> str:
    return "" if value is None else str(value).strip()


def number(value: object) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    try:
        return float(text) if text else 0.0
    except ValueError:
        return 0.0


def strip_service_station_prefix(name: str) -> str:
    """匹配时忽略最前面的 F/f 服务站标记，显示名称保持原样。"""
    return SERVICE_STATION_PREFIX_RE.sub("", clean_name(name), count=1)


def strip_business_words(name: str) -> str:
    """仅生成匹配核心，不修改源表中的原始客户名称。"""
    text = strip_service_station_prefix(name)

    # 活动标记可能出现在名称前、中、后，统一从匹配核心中删除。
    for word in sorted(ACTIVITY_WORDS, key=len, reverse=True):
        text = text.replace(word, "")

    # 其他业务标记只从首尾剥离。
    other_words = [word for word in BUSINESS_PREFIXES if word not in ACTIVITY_WORDS]
    changed = True
    while changed:
        changed = False
        for word in sorted(other_words, key=len, reverse=True):
            if text.startswith(word):
                text = text[len(word):]
                changed = True
                break
            if text.endswith(word):
                text = text[:-len(word)]
                changed = True
                break
    return text


def strip_town_prefix(name: str, towns: Iterable[str]) -> str:
    text = clean_name(name)
    for town in sorted(towns, key=lambda x: len(clean_name(x)), reverse=True):
        town_key = clean_name(town)
        if town_key and text.startswith(town_key):
            return text[len(town_key):]
    return text


def normalize_core(name: str, towns: Iterable[str]) -> str:
    text = strip_business_words(name)
    text = strip_town_prefix(text, towns)
    return text.strip("-—_·,，。:：;；/\\")

def split_person_tokens(name: str, towns: Iterable[str]) -> List[str]:
    """
    将多人客户名称拆成独立姓名。

    示例：
    哈拉海李伟，李明 -> ["李伟", "李明"]
    李连凤/赵长荣 -> ["李连凤", "赵长荣"]
    顾海滨（李洪俊） -> ["顾海滨", "李洪俊"]
    """
    core = normalize_core(name, towns)

    if not core:
        return []

    tokens: List[str] = []

    # 先取得括号外内容。
    outer = re.sub(r"\([^)]*\)", "", core).strip()

    # 支持中文逗号、英文逗号、顿号、斜杠、分号、&、和、及。
    for token in re.split(r"[/、,，;；&]|(?:和)|(?:及)", outer):
        token = clean_name(token).strip("-—_·:：")
        if len(token) >= 2:
            tokens.append(token)

    # 提取括号内姓名。
    for inside in re.findall(r"\(([^)]*)\)", core):
        for token in re.split(r"[/、,，;；&]|(?:和)|(?:及)", inside):
            token = clean_name(token).strip("-—_·:：")
            if len(token) >= 2:
                tokens.append(token)

    # 去重，同时保持原顺序。
    return list(dict.fromkeys(tokens))


def extract_alias_tokens(name: str, towns: Iterable[str]) -> Set[str]:
    """提取严格别名，不使用拼音或编辑距离。"""
    core = normalize_core(name, towns)
    aliases: Set[str] = set()

    if core:
        aliases.add(core)

    # 将多人姓名拆开，例如：
    # 哈拉海李伟，李明 -> 李伟、李明
    # 李连凤/赵长荣 -> 李连凤、赵长荣
    for token in split_person_tokens(name, towns):
        aliases.add(token)

    outer = re.sub(r"\([^)]*\)", "", core).strip()
    if outer:
        aliases.add(outer)

    return {
        clean_name(x)
        for x in aliases
        if len(clean_name(x)) >= 2
    }


def load_alias_store() -> Dict[str, Dict[str, str]]:
    if not ALIAS_FILE.exists():
        return {}
    try:
        data = json.loads(ALIAS_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_alias_store(data: Dict[str, Dict[str, str]]) -> None:
    ALIAS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def find_header_row(ws, required: Iterable[str], max_rows: int = 60) -> Tuple[int, Dict[str, int]]:
    required = list(required)
    for row_idx in range(1, min(ws.max_row, max_rows) + 1):
        values = [display_name(ws.cell(row_idx, c).value) for c in range(1, ws.max_column + 1)]
        mapping = {v: i + 1 for i, v in enumerate(values) if v}
        if all(name in mapping for name in required):
            return row_idx, mapping
    raise ValueError(f"在工作表《{ws.title}》中找不到表头：{', '.join(required)}")


def extract_report_end_date(ws) -> Optional[datetime]:
    pattern = re.compile(r"(20\d{2})[./-](\d{1,2})[./-](\d{1,2})")
    found: List[datetime] = []
    for row in ws.iter_rows(min_row=1, max_row=min(25, ws.max_row), values_only=True):
        for value in row:
            if value is None:
                continue
            for y, m, d in pattern.findall(str(value)):
                try:
                    found.append(datetime(int(y), int(m), int(d)))
                except ValueError:
                    pass
    return max(found) if found else None


# =========================
# 读取三张源数据
# =========================
def read_sales(file_obj) -> Tuple[Dict[str, float], Dict[str, str], Optional[datetime]]:
    wb = load_workbook(file_obj, data_only=True, read_only=True)
    ws = wb.active
    header_row, headers = find_header_row(ws, ["客户"])
    amount_col = headers.get("含税金额") or headers.get("金额")
    if not amount_col:
        raise ValueError("销货单统计表缺少“含税金额”或“金额”列。")

    totals: Dict[str, float] = {}
    originals: Dict[str, str] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        original = display_name(ws.cell(r, headers["客户"]).value)
        key = clean_name(original)
        if not key or key in {"合计", "总计"}:
            continue
        totals[key] = totals.get(key, 0.0) + number(ws.cell(r, amount_col).value)
        originals.setdefault(key, original)
    return totals, originals, extract_report_end_date(ws)


def read_orders(file_obj) -> Tuple[Dict[str, float], Dict[str, str], Optional[datetime]]:
    wb = load_workbook(file_obj, data_only=True, read_only=True)
    ws = wb.active
    header_row, headers = find_header_row(ws, ["客户", "未销货金额"])

    totals: Dict[str, float] = {}
    originals: Dict[str, str] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        original = display_name(ws.cell(r, headers["客户"]).value)
        key = clean_name(original)
        if not key or key in {"合计", "总计"}:
            continue
        totals[key] = totals.get(key, 0.0) + number(ws.cell(r, headers["未销货金额"]).value)
        originals.setdefault(key, original)
    return totals, originals, extract_report_end_date(ws)


def read_ledger(file_obj) -> Tuple[Dict[str, LedgerValue], Dict[str, str], Optional[datetime]]:
    wb = load_workbook(file_obj, data_only=True, read_only=True)
    ws = wb.active
    header_row, headers = find_header_row(ws, ["结算客户", "期初余额", "应收", "已收", "期末余额"])

    totals: Dict[str, LedgerValue] = {}
    originals: Dict[str, str] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        original = display_name(ws.cell(r, headers["结算客户"]).value)
        key = clean_name(original)
        if not key or key in {"合计", "总计"}:
            continue
        current = totals.setdefault(key, LedgerValue())
        current.opening += number(ws.cell(r, headers["期初余额"]).value)
        current.receivable += number(ws.cell(r, headers["应收"]).value)
        current.received += number(ws.cell(r, headers["已收"]).value)
        current.ending += number(ws.cell(r, headers["期末余额"]).value)
        originals.setdefault(key, original)
    return totals, originals, extract_report_end_date(ws)


# =========================
# 基础客户与精准匹配
# =========================
def find_total_row(ws) -> int:
    for r in range(ws.max_row, 1, -1):
        if display_name(ws.cell(r, 1).value) in {"合计", "总计"}:
            return r
    return ws.max_row + 1


def read_base_customers(base_bytes: bytes):
    wb = load_workbook(io.BytesIO(base_bytes), data_only=False)
    towns = [
        name for name in wb.sheetnames
        if name not in {SUMMARY_SHEET, CHECK_SHEET} | AUXILIARY_SHEETS
    ]
    customers: Dict[str, CustomerRecord] = {}
    alias_index: Dict[str, Set[str]] = defaultdict(set)

    for town in towns:
        ws = wb[town]
        total_row = find_total_row(ws)
        for r in range(3, total_row):
            original = display_name(ws.cell(r, 1).value)
            key = clean_name(original)
            if not key:
                continue
            core = normalize_core(original, towns)
            aliases = extract_alias_tokens(original, towns)
            rec = CustomerRecord(
                key=key,
                name=original,
                town=town,
                row=r,
                core=core,
                aliases=aliases,
            )
            customers[key] = rec
            for alias in aliases:
                alias_index[alias].add(key)
    return wb, towns, customers, alias_index


def infer_explicit_town(source_name: str, towns: List[str]) -> Optional[str]:
    if source_name in SPECIAL_TOWN_DEFAULTS and SPECIAL_TOWN_DEFAULTS[source_name] in towns:
        return SPECIAL_TOWN_DEFAULTS[source_name]
    clean = clean_name(source_name)
    for town in sorted(towns, key=len, reverse=True):
        if clean.startswith(clean_name(town)):
            return town
    return None


def decide_match(
    source_key: str,
    source_name: str,
    towns: List[str],
    customers: Dict[str, CustomerRecord],
    alias_index: Dict[str, Set[str]],
    saved_aliases: Dict[str, Dict[str, str]],
) -> MatchDecision:
    # 1. 完整名称完全一致：直接合并。
    if source_key in customers:
        return MatchDecision(
            source_key, source_name, "exact",
            target_key=source_key,
            reason="完整名称完全一致",
        )

    # 2. 人工记忆规则。
    saved = saved_aliases.get(source_key)
    if saved:
        action = saved.get("action")
        if action == "merge" and saved.get("target_key") in customers:
            return MatchDecision(
                source_key, source_name, "alias",
                target_key=saved["target_key"],
                reason="使用已保存的人工合并规则",
            )
        if action == "new" and saved.get("town") in towns:
            return MatchDecision(
                source_key, source_name, "saved_new",
                suggested_town=saved["town"],
                reason="使用已保存的独立客户归属规则",
            )
        if action == "ignore":
            return MatchDecision(
                source_key, source_name, "ignored",
                reason="使用已保存的删除/忽略规则",
            )

    source_core = normalize_core(source_name, towns)
    explicit_town = infer_explicit_town(source_name, towns)
    person_tokens = split_person_tokens(source_name, towns)

    # 第一姓名作为主要姓名。
    # 哈拉海李伟，李明 -> primary_core = 李伟
    primary_core = person_tokens[0] if person_tokens else source_core

    candidates: Set[str] = set()

    # 3. 优先按主要姓名匹配，不做拼音或编辑距离。
    if primary_core:
        candidates |= alias_index.get(primary_core, set())

    # 如果源名称中带明确乡镇，优先只保留相同乡镇的客户。
    if explicit_town and candidates:
        same_town_candidates = {
            key
            for key in candidates
            if customers[key].town == explicit_town
        }

        if same_town_candidates:
            candidates = same_town_candidates

    # 主要姓名只有一个可靠客户时，直接自动合并。
    if len(candidates) == 1:
        target_key = next(iter(candidates))
        town = customers[target_key].town

        return MatchDecision(
            source_key,
            source_name,
            "rule_auto",
            target_key=target_key,
            suggested_town=town,
            candidates=[target_key],
            reason=(
                f"规则自动识别：主要姓名“{primary_core}”"
                f"唯一对应{town}的“{customers[target_key].name}”"
            ),
        )

    if len(candidates) > 1:
        ordered = sorted(candidates, key=lambda k: (customers[k].town, customers[k].name))
        return MatchDecision(
            source_key, source_name, "ambiguous",
            suggested_town=explicit_town,
            candidates=ordered,
            reason=f"姓名核心“{source_core}”对应多个客户，必须人工选择",
        )

    if explicit_town:
        return MatchDecision(
            source_key, source_name, "new",
            suggested_town=explicit_town,
            reason="名称中包含明确乡镇前缀",
        )

    fallback = "其余" if "其余" in towns else towns[0]
    return MatchDecision(
        source_key, source_name, "new",
        suggested_town=fallback,
        reason="未找到唯一可靠关联",
    )


# =========================
# 写入 Excel
# =========================
def copy_row_style(ws, source_row: int, target_row: int, max_col: int = 10) -> None:
    """复制整行样式。仅处理导出的Excel，不修改Streamlit界面。"""
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)


def ensure_town_layout(ws) -> None:
    """
    乡镇明细固定为10列：
    A客户、B25年销货、C26年销货单、D26年订单、E26年合计、
    F期初、G已收、H总应收、I客户余款、J总销售。
    """
    headers = [
        "客户", "25年销货", "26年销货单", "26年订单", "26年合计",
        "期初", "已收", "总应收", "客户余款", "总销售",
    ]

    # J列表头沿用I列样式。
    if ws.max_column < 10 or not ws.cell(2, 10).has_style:
        source = ws.cell(2, 9 if ws.max_column >= 9 else max(1, ws.max_column))
        target = ws.cell(2, 10)
        if source.has_style:
            target._style = copy(source._style)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format

    for col, title in enumerate(headers, start=1):
        ws.cell(2, col).value = title

    # 标题横跨10列。
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row == 1 and merged.max_row == 1 and merged.min_col == 1:
            ws.unmerge_cells(str(merged))
    ws.merge_cells("A1:J1")

    ws.freeze_panes = "A3"
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 30)
    for col in "BCDEFGHIJ":
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, 15)


def add_customer_row(ws, customer_name: str) -> int:
    """始终在合计行上方新增客户，避免合计行被打乱。"""
    ensure_town_layout(ws)
    total_row = find_total_row(ws)
    ws.insert_rows(total_row, 1)
    copy_row_style(ws, max(3, total_row - 1), total_row, max_col=10)
    ws.cell(total_row, 1).value = customer_name
    ws.cell(total_row, 2).value = 0
    for col in range(3, 11):
        ws.cell(total_row, col).value = 0
    return total_row


def _clear_tables(ws) -> None:
    """删除工作表中的旧 Excel Table 定义，避免 table*.xml 损坏。"""
    for table_name in list(ws.tables.keys()):
        del ws.tables[table_name]
    ws.auto_filter.ref = None


def _apply_stable_filter(ws, total_row: int, max_col: int = 10) -> None:
    """只筛选表头和客户数据，不包含最后的合计行。"""
    _clear_tables(ws)
    last_col = get_column_letter(max_col)
    last_data_row = total_row - 1
    if last_data_row >= 3:
        ws.auto_filter.ref = f"A2:{last_col}{last_data_row}"
    else:
        ws.auto_filter.ref = f"A2:{last_col}2"


def rebuild_town_total(ws, table_name: str = "") -> int:
    """
    重建乡镇合计，并用Excel Table的汇总行锁定在底部。
    客户数据可以自由升序、降序和筛选，但合计行永远不参与排序。
    """
    ensure_town_layout(ws)
    total_row = find_total_row(ws)
    last_data_row = total_row - 1
    ws.cell(total_row, 1).value = "合计"

    # B-G正常求和。
    for col in range(2, 8):
        letter = get_column_letter(col)
        ws.cell(total_row, col).value = (
            f"=SUM({letter}3:{letter}{last_data_row})"
            if last_data_row >= 3 else 0
        )

    if last_data_row >= 3:
        # 乡镇总应收：负数预付款不能抵消其他客户欠款。
        ws.cell(total_row, 8).value = (
            f"=SUM(H3:H{last_data_row})+SUM(I3:I{last_data_row})"
        )
        ws.cell(total_row, 9).value = f"=SUM(I3:I{last_data_row})"
        ws.cell(total_row, 10).value = f"=SUM(J3:J{last_data_row})"
    else:
        for col in range(8, 11):
            ws.cell(total_row, col).value = 0

    # J列沿用I列样式。
    if ws.cell(total_row, 9).has_style:
        ws.cell(total_row, 10)._style = copy(ws.cell(total_row, 9)._style)
        ws.cell(total_row, 10).alignment = copy(ws.cell(total_row, 9).alignment)
        ws.cell(total_row, 10).number_format = ws.cell(total_row, 9).number_format

    for row in ws.iter_rows(min_row=2, max_row=total_row, min_col=1, max_col=10):
        for cell in row:
            font = copy(cell.font)
            font.sz = 14
            cell.font = font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

    # 稳定方案：筛选范围不包含合计行，不创建 Excel Table。
    _apply_stable_filter(ws, total_row, max_col=10)
    return total_row


def town_sum_formula(town: str, value_col: str) -> str:
    """按客户数据动态求和，不依赖合计行的具体行号。"""
    quoted = town.replace("'", "''")
    return (
        f'=SUMIFS(\'{quoted}\'!{value_col}:{value_col},'
        f'\'{quoted}\'!A:A,"<>合计",\'{quoted}\'!A:A,"<>")'
    )


def find_summary_total_row(ws) -> int:
    for r in range(ws.max_row, 2, -1):
        if display_name(ws.cell(r, 2).value) in {"总计", "合计"}:
            return r
    return ws.max_row + 1


def normalize_town_key(value: object) -> str:
    """
    地区名称比对专用：忽略空格和常见隐藏字符。
    例如汇总中的“S松原”和工作表“S 松原”视为同一个地区。
    """
    text = clean_name(value)
    text = text.replace("$", "S")
    return text


def rebuild_summary_sheet(ws, towns: List[str], end_date: datetime) -> None:
    """
    重新生成汇总地区明细，保证每个工作表只出现一次。
    同时保留：客户余款、总销售、已收款和不被预付款抵消的总应收。
    """
    headers = [
        "序号", "地区", "上年销货单（不变）", "销货单（按乡镇更新）",
        "订单金额（按乡镇更新）", "2026年4.1-当前 合计",
        "客户余款（实时更新）", "总销售（合计+预付款）",
        "已收款（实时更新）", "总应收（不算他人的预付款）",
    ]

    # 保存原汇总中地区的显示写法，如“S松原”。
    old_total_row = find_summary_total_row(ws)
    display_labels: Dict[str, str] = {}
    for r in range(3, old_total_row):
        label = display_name(ws.cell(r, 2).value)
        if label:
            display_labels.setdefault(normalize_town_key(label), label)

    # 保存数据行和总计行样式模板。
    data_style_row = 3 if old_total_row > 3 else max(2, old_total_row)
    total_style_row = old_total_row

    # 标题与表头扩展到J列。
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row == 1 and merged.max_row == 1 and merged.min_col == 1:
            ws.unmerge_cells(str(merged))
    ws.merge_cells("A1:J1")
    ws.cell(1, 1).value = (
        "农安县汇发烟花有限公司各乡镇销售统计表"
        f"26.4.1-{end_date.strftime('%y.%m.%d')}"
    )

    for col, header in enumerate(headers, start=1):
        # 新增列沿用原最后一列表头样式。
        if col > 7 and ws.cell(2, min(7, ws.max_column)).has_style:
            ws.cell(2, col)._style = copy(ws.cell(2, min(7, ws.max_column))._style)
            ws.cell(2, col).alignment = copy(ws.cell(2, min(7, ws.max_column)).alignment)
            ws.cell(2, col).number_format = ws.cell(2, min(7, ws.max_column)).number_format
        ws.cell(2, col).value = header

    # 彻底重建地区行，解决“S松原/S 松原”被重复插入的问题。
    old_detail_count = max(0, old_total_row - 3)
    if old_detail_count:
        ws.delete_rows(3, old_detail_count)

    # 删除后总计行位于第3行，先在其上方插入准确数量的地区行。
    ws.insert_rows(3, len(towns))
    new_total_row = 3 + len(towns)

    for idx, town in enumerate(towns, start=1):
        r = idx + 2
        copy_row_style(ws, data_style_row, r, max_col=10)

        # 优先沿用原汇总显示名称，否则使用工作表名去空格后的名称。
        town_key = normalize_town_key(town)
        label = display_labels.get(town_key, clean_name(town))

        ws.cell(r, 1).value = idx
        ws.cell(r, 2).value = label
        ws.cell(r, 3).value = town_sum_formula(town, "B")
        ws.cell(r, 4).value = town_sum_formula(town, "C")
        ws.cell(r, 5).value = town_sum_formula(town, "D")
        ws.cell(r, 6).value = f"=D{r}+E{r}"
        ws.cell(r, 7).value = town_sum_formula(town, "I")
        ws.cell(r, 8).value = f"=F{r}+G{r}"
        ws.cell(r, 9).value = town_sum_formula(town, "G")
        # 正数应收 = 客户总应收净额 + 客户余款。
        ws.cell(r, 10).value = (
            f"={town_sum_formula(town, 'H')[1:]}+{town_sum_formula(town, 'I')[1:]}"
        )

    # 总计行。
    copy_row_style(ws, total_style_row, new_total_row, max_col=10)
    ws.cell(new_total_row, 1).value = ""
    ws.cell(new_total_row, 2).value = "总计"
    for col in range(3, 11):
        letter = get_column_letter(col)
        ws.cell(new_total_row, col).value = (
            f"=SUM({letter}3:{letter}{new_total_row - 1})"
            if new_total_row > 3 else 0
        )

    # 汇总页只筛选地区明细，不包含最后的总计行。
    _apply_stable_filter(ws, new_total_row, max_col=10)
    ws.freeze_panes = "A3"

    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 9)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 14)
    for col in "CDEFGHIJ":
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, 20)

    for row in ws.iter_rows(min_row=1, max_row=new_total_row, min_col=1, max_col=10):
        for cell in row:
            font = copy(cell.font)
            font.sz = 18
            cell.font = font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )


def build_output(
    base_bytes: bytes,
    sales: Dict[str, float],
    orders: Dict[str, float],
    ledger: Dict[str, LedgerValue],
    originals: Dict[str, str],
    assignments: Dict[str, Tuple[str, str]],
    end_date: datetime,
    decisions: Optional[Dict[str, MatchDecision]] = None,
) -> Tuple[bytes, Dict[str, int]]:
    wb, towns, customers, _ = read_base_customers(base_bytes)

    # 在任何插入、删除行之前，先清除底表中可能残留的 Table 定义。
    for sheet in wb.worksheets:
        _clear_tables(sheet)

    source_to_target: Dict[str, str] = {}
    all_source_keys = set(sales) | set(orders) | set(ledger)

    for source_key in all_source_keys:
        if source_key in customers:
            source_to_target[source_key] = source_key
            continue

        action, value = assignments[source_key]
        if action == "ignore":
            continue

        if action == "merge":
            if value not in customers:
                raise ValueError(f"客户合并目标不存在：{value}")
            source_to_target[source_key] = value
            continue

        # 保留独立名头并新增到所选乡镇。
        town = value
        if town not in towns:
            raise ValueError(f"乡镇不存在：{town}")
        original = originals.get(source_key, source_key)
        target_key = clean_name(original)
        if target_key not in customers:
            row = add_customer_row(wb[town], original)
            customers[target_key] = CustomerRecord(
                target_key,
                original,
                town,
                row,
                normalize_core(original, towns),
                set(),
            )
        source_to_target[source_key] = target_key

    # 清空当期字段，保留25年固定值。
    for rec in customers.values():
        ws = wb[rec.town]
        ensure_town_layout(ws)
        row = rec.row
        ws.cell(row, 3).value = 0
        ws.cell(row, 4).value = 0
        ws.cell(row, 5).value = f"=C{row}+D{row}"
        ws.cell(row, 6).value = 0
        ws.cell(row, 7).value = 0
        ws.cell(row, 8).value = f"=E{row}+F{row}-G{row}"
        ws.cell(row, 9).value = f'=IF(H{row}<0,ABS(H{row}),0)'
        ws.cell(row, 10).value = f"=E{row}+I{row}"

    agg_sales: Dict[str, float] = defaultdict(float)
    agg_orders: Dict[str, float] = defaultdict(float)
    agg_ledger: Dict[str, LedgerValue] = {}

    for source_key, amount in sales.items():
        target = source_to_target.get(source_key)
        if target:
            agg_sales[target] += amount

    for source_key, amount in orders.items():
        target = source_to_target.get(source_key)
        if target:
            agg_orders[target] += amount

    for source_key, value in ledger.items():
        target = source_to_target.get(source_key)
        if not target:
            continue
        cur = agg_ledger.setdefault(target, LedgerValue())
        cur.opening += value.opening
        cur.receivable += value.receivable
        cur.received += value.received
        cur.ending += value.ending

    for target_key, rec in customers.items():
        ws = wb[rec.town]
        row = rec.row
        lv = agg_ledger.get(target_key, LedgerValue())
        ws.cell(row, 3).value = round(agg_sales.get(target_key, 0.0), 2)
        ws.cell(row, 4).value = round(agg_orders.get(target_key, 0.0), 2)
        ws.cell(row, 5).value = f"=C{row}+D{row}"
        ws.cell(row, 6).value = round(lv.opening, 2)
        ws.cell(row, 7).value = round(lv.received, 2)
        ws.cell(row, 8).value = f"=E{row}+F{row}-G{row}"
        ws.cell(row, 9).value = f'=IF(H{row}<0,ABS(H{row}),0)'
        ws.cell(row, 10).value = f"=E{row}+I{row}"

    for town_index, town in enumerate(towns, start=1):
        town_ws = wb[town]
        rebuild_town_total(town_ws, f"TownTable_{town_index:03d}")
        town_ws.cell(1, 1).value = (
            f"销货单统计表（更新至{end_date.strftime('%Y.%m.%d')}）"
        )

    if SUMMARY_SHEET in wb.sheetnames:
        rebuild_summary_sheet(wb[SUMMARY_SHEET], towns, end_date)

    # Merge_Log：记录每个源名称如何归并及其金额构成。
    if MERGE_LOG_SHEET in wb.sheetnames:
        del wb[MERGE_LOG_SHEET]
    merge_log = wb.create_sheet(MERGE_LOG_SHEET, 1)
    merge_log.append([
        "序号", "原始名称", "最终客户", "所属乡镇", "匹配方式", "匹配说明",
        "销货单", "订单", "期初", "已收", "26年合计", "总应收影响",
    ])

    log_index = 0
    for source_key in sorted(all_source_keys, key=lambda key: originals.get(key, key)):
        target_key = source_to_target.get(source_key)
        if not target_key:
            continue

        rec = customers[target_key]
        ledger_value = ledger.get(source_key, LedgerValue())
        sale_amount = round(sales.get(source_key, 0.0), 2)
        order_amount = round(orders.get(source_key, 0.0), 2)
        opening_amount = round(ledger_value.opening, 2)
        received_amount = round(ledger_value.received, 2)
        current_total = round(sale_amount + order_amount, 2)
        receivable_effect = round(current_total + opening_amount - received_amount, 2)

        decision = (decisions or {}).get(source_key)
        if source_key == target_key:
            match_method = "完整名称一致"
            match_reason = "完整名称完全一致"
        elif decision and decision.status == "rule_auto":
            match_method = "规则自动合并"
            match_reason = decision.reason
        elif decision and decision.status == "alias":
            match_method = "已保存规则合并"
            match_reason = decision.reason
        else:
            match_method = "人工确认合并"
            match_reason = decision.reason if decision else ""

        log_index += 1
        merge_log.append([
            log_index, originals.get(source_key, source_key), rec.name, rec.town,
            match_method, match_reason, sale_amount, order_amount, opening_amount,
            received_amount, current_total, receivable_effect,
        ])

    for cell in merge_log[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True, size=14)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in merge_log.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(size=12)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col in range(7, 13):
            row[col - 1].number_format = "#,##0.00"

    merge_log.column_dimensions["A"].width = 8
    merge_log.column_dimensions["B"].width = 28
    merge_log.column_dimensions["C"].width = 28
    merge_log.column_dimensions["D"].width = 14
    merge_log.column_dimensions["E"].width = 18
    merge_log.column_dimensions["F"].width = 52
    for col in "GHIJKL":
        merge_log.column_dimensions[col].width = 16
    merge_log.freeze_panes = "A2"
    if merge_log.max_row >= 2:
        merge_log.auto_filter.ref = f"A1:L{merge_log.max_row}"

    # 更新检查页。
    if CHECK_SHEET in wb.sheetnames:
        del wb[CHECK_SHEET]
    check = wb.create_sheet(CHECK_SHEET, 1)
    check.append(["更新检查", "结果"])
    check.append(["截止日期", end_date.strftime("%Y-%m-%d")])
    check.append(["基础客户数（更新后）", len(customers)])
    check.append(["销货单客户数", len(sales)])
    check.append(["订单客户数", len(orders)])
    check.append(["应收总账客户数", len(ledger)])
    check.append(["人工确认名称数", len(assignments)])
    check.append([])
    check.append(["源表客户", "处理结果"])

    for source_key in sorted(assignments):
        action, value = assignments[source_key]
        if action == "merge":
            rec = customers[value]
            result = f"合并至：{rec.name}（{rec.town}）"
        elif action == "ignore":
            result = "删除/忽略：不进入本次结果"
        else:
            result = f"保留独立名头，归属：{value}"
        check.append([originals.get(source_key, source_key), result])

    for cell in check[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True, size=16)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in check.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(size=14)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    check.column_dimensions["A"].width = 36
    check.column_dimensions["B"].width = 54
    check.freeze_panes = "A2"

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    # 保存前再次清除所有 Table 定义，保证输出包内不存在 /xl/tables/table*.xml。
    for sheet in wb.worksheets:
        _clear_tables(sheet)

    output = io.BytesIO()
    wb.save(output)

    # 用 openpyxl 重新读取一次，确认文件结构可正常打开。
    output.seek(0)
    check_wb = load_workbook(output, data_only=False, read_only=False)
    check_wb.close()
    output.seek(0)

    stats = {
        "客户总数": len(customers),
        "乡镇数": len(towns),
        "销货客户": len(sales),
        "订单客户": len(orders),
        "应收客户": len(ledger),
        "人工确认": len(assignments),
        "删除忽略": sum(1 for action, _ in assignments.values() if action == "ignore"),
    }
    return output.getvalue(), stats


def reset_workflow_state() -> None:
    for key in [
        "parsed", "base_bytes", "result_bytes", "result_stats",
        "end_date", "decisions", "last_read_signature",
    ]:
        st.session_state.pop(key, None)


# =========================
# 页面
# =========================
st.markdown(
    """
    <div class="hero">
      <div class="hero-title">📊 汇发销售自动更新系统 V3</div>
      <div class="hero-subtitle">精准匹配 · 人工确认 · 本机处理 · 自动生成乡镇及总汇总</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("查看精准匹配规则", expanded=False):
    st.markdown(
        """
        <div class="rule-card">
        1. 完整名称完全一致：自动归入原客户。<br>
        2. 已人工确认并保存的规则：以后自动执行。<br>
        3. 姓名核心全库唯一：只推荐乡镇，默认仍保留独立客户名头。<br>
        4. 括号内姓名可识别，例如“鲍家顾海滨（李洪俊）”。<br>
        5. 存在多个同名候选：必须人工选择，绝不自动错合并。<br>
        6. 可删除/忽略客户，并永久记住。<br>
        7. 不使用拼音或相似度自动合并，避免韩冰、韩兵、韩斌被混在一起。
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="section-title">第一步：上传四张 Excel</div>', unsafe_allow_html=True)

left, right = st.columns(2, gap="large")
with left:
    st.markdown('<div class="upload-title">① 康总基础数据</div>', unsafe_allow_html=True)
    base_file = st.file_uploader(
        "上传康总基础数据",
        type=["xlsx"],
        key="base_file",
        label_visibility="collapsed",
    )
    st.markdown('<div class="upload-title">② 销货单统计表</div>', unsafe_allow_html=True)
    sales_file = st.file_uploader(
        "上传销货单统计表",
        type=["xlsx"],
        key="sales_file",
        label_visibility="collapsed",
    )

with right:
    st.markdown('<div class="upload-title">③ 销售订单执行表</div>', unsafe_allow_html=True)
    order_file = st.file_uploader(
        "上传销售订单执行表",
        type=["xlsx"],
        key="order_file",
        label_visibility="collapsed",
    )
    st.markdown('<div class="upload-title">④ 应收总账</div>', unsafe_allow_html=True)
    ledger_file = st.file_uploader(
        "上传应收总账",
        type=["xlsx"],
        key="ledger_file",
        label_visibility="collapsed",
    )

if st.button("读取并精准检查", type="primary", use_container_width=True):
    reset_workflow_state()
    if not all([base_file, sales_file, order_file, ledger_file]):
        st.error("请先上传完整的四张 Excel。")
    else:
        try:
            with st.spinner("正在读取四张表并检查客户名称……"):
                base_bytes = base_file.getvalue()
                _, towns, customers, alias_index = read_base_customers(base_bytes)
                sales, sales_names, sales_date = read_sales(io.BytesIO(sales_file.getvalue()))
                orders, order_names, order_date = read_orders(io.BytesIO(order_file.getvalue()))
                ledger, ledger_names, ledger_date = read_ledger(io.BytesIO(ledger_file.getvalue()))

                originals = {**sales_names, **order_names, **ledger_names}
                saved_aliases = load_alias_store()
                all_source = sorted(set(sales) | set(orders) | set(ledger))
                decisions = {
                    key: decide_match(
                        key,
                        originals.get(key, key),
                        towns,
                        customers,
                        alias_index,
                        saved_aliases,
                    )
                    for key in all_source
                }

                dates = [d for d in [sales_date, order_date, ledger_date] if d]
                end_date = max(dates) if dates else datetime.today()

                st.session_state.parsed = {
                    "towns": towns,
                    "customers": customers,
                    "sales": sales,
                    "orders": orders,
                    "ledger": ledger,
                    "originals": originals,
                }
                st.session_state.decisions = decisions
                st.session_state.base_bytes = base_bytes
                st.session_state.end_date = end_date

            review_count = sum(
                1 for d in decisions.values()
                if d.status not in {
                    "exact",
                    "alias",
                    "rule_auto",
                    "saved_new",
                    "ignored",
                }
            )
            st.success(
                f"读取成功。截止日期：{end_date.strftime('%Y-%m-%d')}；"
                f"需要人工确认：{review_count} 个名称。"
            )
        except Exception as exc:
            st.exception(exc)


if "parsed" in st.session_state:
    parsed = st.session_state.parsed
    decisions: Dict[str, MatchDecision] = st.session_state.decisions
    review = [
        d for d in decisions.values()
       if d.status not in {
                           "exact",
                           "alias",
                           "rule_auto",
                           "saved_new",
                           "ignored",
        }
    ]

    st.markdown('<div class="section-title">第二步：人工确认非完全匹配名称</div>', unsafe_allow_html=True)

    info_cols = st.columns(5)
    info_cols[0].metric("基础客户", len(parsed["customers"]))
    info_cols[1].metric("销货客户", len(parsed["sales"]))
    info_cols[2].metric("订单客户", len(parsed["orders"]))
    info_cols[3].metric("应收客户", len(parsed["ledger"]))
    info_cols[4].metric("待确认", len(review))

    if not review:
        st.success("所有名称均已完全匹配或已有保存规则，可以直接生成 Excel。")

    assignments: Dict[str, Tuple[str, str]] = {}
    alias_updates: Dict[str, Dict[str, str]] = {}

    with st.form("matching_form"):
        for idx, decision in enumerate(review):
            original = decision.source_name
            candidate_keys = decision.candidates or []
            candidate_labels = [
                f"合并到：{parsed['customers'][k].name}｜{parsed['customers'][k].town}"
                for k in candidate_keys
            ]
            default_town = (
                decision.suggested_town
                if decision.suggested_town in parsed["towns"]
                else parsed["towns"][0]
            )
            options = [
                "保留独立客户名头（推荐）",
                "删除/忽略该客户（不进入结果）",
            ] + candidate_labels

            with st.container(border=True):
                st.markdown(
                    f'<div class="review-name">{idx + 1}. {original}</div>',
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f'<div class="review-reason">{decision.reason}</div>',
                    unsafe_allow_html=True,
                )

                action_col, town_col = st.columns([2, 1], gap="medium")
                with action_col:
                    choice = st.selectbox(
                        "处理方式",
                        options,
                        key=f"choice_{idx}",
                    )
                with town_col:
                    town = st.selectbox(
                        "所属乡镇",
                        parsed["towns"],
                        index=parsed["towns"].index(default_town),
                        key=f"town_{idx}",
                        disabled=choice != options[0],
                    )

                remember = st.checkbox(
                    "记住本次处理，下次自动执行",
                    value=True,
                    key=f"remember_{idx}",
                )

                if choice == options[0]:
                    assignments[decision.source_key] = ("new", town)
                    if remember:
                        alias_updates[decision.source_key] = {
                            "action": "new",
                            "town": town,
                            "display_name": original,
                        }
                elif choice == options[1]:
                    assignments[decision.source_key] = ("ignore", "")
                    if remember:
                        alias_updates[decision.source_key] = {
                            "action": "ignore",
                            "display_name": original,
                        }
                else:
                    selected = candidate_keys[options.index(choice) - 2]
                    assignments[decision.source_key] = ("merge", selected)
                    if remember:
                        alias_updates[decision.source_key] = {
                            "action": "merge",
                            "target_key": selected,
                            "display_name": original,
                        }

        submitted = st.form_submit_button(
            "生成更新后的 Excel",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        try:
            for key, decision in decisions.items():
                if decision.status in {"exact", "alias", "rule_auto"}:
                    assignments[key] = ("merge", decision.target_key)
                elif decision.status == "ignored":
                    assignments[key] = ("ignore", "")
                elif decision.status == "saved_new":
                    assignments[key] = ("new", decision.suggested_town)

            with st.spinner("正在更新乡镇明细、应收、余款和总汇总……"):
                result, stats = build_output(
                    st.session_state.base_bytes,
                    parsed["sales"],
                    parsed["orders"],
                    parsed["ledger"],
                    parsed["originals"],
                    assignments,
                    st.session_state.end_date,
                    decisions,
                )

                if alias_updates:
                    store = load_alias_store()
                    store.update(alias_updates)
                    save_alias_store(store)

            st.session_state.result_bytes = result
            st.session_state.result_stats = stats
            st.success("更新完成。已勾选的处理规则已保存到本机。")
        except Exception as exc:
            st.exception(exc)


if "result_bytes" in st.session_state:
    st.markdown('<div class="section-title">第三步：下载结果</div>', unsafe_allow_html=True)
    stats = st.session_state.result_stats
    cols = st.columns(len(stats))
    for col, (label, value) in zip(cols, stats.items()):
        col.metric(label, value)

    date_text = st.session_state.end_date.strftime("%Y-%m-%d")
    st.download_button(
        "⬇️ 下载更新后的康总数据",
        data=st.session_state.result_bytes,
        file_name=f"26年给康总数据_更新至{date_text}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

st.markdown(
    '<div class="footer-note">所有 Excel 均在本机内存中处理，不上传外部服务器。</div>',
    unsafe_allow_html=True,
)
