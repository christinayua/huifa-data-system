
import io
import re
from copy import copy
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="应收账款跟进表更新",
    page_icon="💰",
    layout="wide",
)

OUTPUT_HEADERS = [
    "序号",
    "客户名称",
    "总应收",
    "负责人",
    "联系日期",
    "可付款时间",
    "付款状态",
    "下次跟进日期",
]

FOLLOWUP_FIELDS = [
    "负责人",
    "联系日期",
    "可付款时间",
    "付款状态",
    "下次跟进日期",
]

MASTER_FIELDS = [
    "客户",
    "总应收",
    "客户余款",
]


def normalize_name(value):
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"[\s\u3000]+", "", text)
    return text


def relaxed_name(value):
    """用于兼容 F/S 前缀、活动账户等常见写法。"""
    text = normalize_name(value)
    text = re.sub(r"^[FSfs]+", "", text)
    text = text.replace("参加活动的", "")
    text = text.replace("参加活动", "")
    text = text.replace("活动", "")
    return text


def safe_number(value):
    if value in (None, ""):
        return 0.0
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return 0.0


def find_header_row(ws, required_headers, max_rows=10):
    for row_idx in range(1, min(ws.max_row, max_rows) + 1):
        values = [
            str(ws.cell(row_idx, col_idx).value).strip()
            if ws.cell(row_idx, col_idx).value is not None else ""
            for col_idx in range(1, ws.max_column + 1)
        ]
        if all(header in values for header in required_headers):
            return row_idx, values
    return None, None


def read_master_records(source_bytes):
    """读取《给康总数据》中全部客户。"""
    wb = load_workbook(io.BytesIO(source_bytes), data_only=True)

    records = {}
    duplicate_names = []
    skipped_sheets = []

    for ws in wb.worksheets:
        if ws.title == "汇总":
            continue

        header_row, headers = find_header_row(ws, ["客户", "总应收", "客户余款"])
        if header_row is None:
            skipped_sheets.append(ws.title)
            continue

        index = {
            header: headers.index(header) + 1
            for header in MASTER_FIELDS
        }

        for row_idx in range(header_row + 1, ws.max_row + 1):
            customer = ws.cell(row_idx, index["客户"]).value
            if not customer:
                continue

            customer_text = str(customer).strip()
            if customer_text in {"合计", "总计"}:
                continue

            record = {
                "客户名称": customer_text,
                "总应收": safe_number(ws.cell(row_idx, index["总应收"]).value),
                "客户余款": safe_number(ws.cell(row_idx, index["客户余款"]).value),
                "来源乡镇": ws.title,
            }

            key = normalize_name(customer_text)
            if key in records:
                duplicate_names.append(customer_text)
            records[key] = record

    return records, duplicate_names, skipped_sheets


def find_receivable_sheet(wb):
    for ws in wb.worksheets:
        header_row, _ = find_header_row(ws, ["客户名称", "总应收"])
        if header_row:
            return ws
    return wb.active


def read_old_followup(template_bytes):
    """读取旧应收跟进表中的人工跟进信息。"""
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = find_receivable_sheet(wb)

    # 原文件带有 Excel Table。更新客户数量后，旧 Table 的 ref / totalsRow
    # 容易与新数据范围不一致，Excel 打开时会提示修复。
    # 本 App 保留单元格样式与筛选数据，但移除底层 Table 定义，确保文件稳定。
    for table_name in list(ws.tables.keys()):
        del ws.tables[table_name]

    header_row, headers = find_header_row(ws, ["客户名称", "总应收"])

    if header_row is None:
        raise ValueError("旧表中未找到“客户名称”和“总应收”表头。")

    col_map = {
        header: headers.index(header) + 1
        for header in headers
        if header
    }

    missing = [header for header in OUTPUT_HEADERS if header not in col_map]
    if missing:
        raise ValueError("旧应收表缺少字段：" + "、".join(missing))

    history = {}
    old_names = []

    for row_idx in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row_idx, col_map["客户名称"]).value
        if not name or str(name).strip() in {"合计", "总计"}:
            continue
        if str(name).strip().startswith("共") and "户" in str(name):
            continue

        name_text = str(name).strip()
        item = {
            field: ws.cell(row_idx, col_map[field]).value
            for field in FOLLOWUP_FIELDS
        }

        history[normalize_name(name_text)] = item
        history[relaxed_name(name_text)] = item
        old_names.append(name_text)

    return wb, ws, header_row, col_map, history, old_names


def create_blank_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "应收账款跟进表"

    headers = OUTPUT_HEADERS
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(2, col_idx, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    widths = [10, 34, 16, 14, 16, 18, 16, 18]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = None
    return wb, ws, 2


def copy_row_style(ws, source_row, target_row, max_col=8):
    if source_row < 1 or source_row > ws.max_row:
        return

    for col_idx in range(1, max_col + 1):
        source = ws.cell(source_row, col_idx)
        target = ws.cell(target_row, col_idx)

        if source.has_style:
            target._style = copy(source._style)
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)

    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def clear_data_area(ws, start_row, end_row, max_col=8):
    for row_idx in range(start_row, end_row + 1):
        for col_idx in range(1, max_col + 1):
            ws.cell(row_idx, col_idx).value = None


def style_total_row(ws, row_idx):
    thin = Side(style="thin", color="D6B656")
    for col_idx in range(1, 9):
        cell = ws.cell(row_idx, col_idx)
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
        cell.font = Font(bold=True, color="7F6000")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def build_receivable_output(source_bytes, template_bytes=None):
    master_records, duplicates, skipped = read_master_records(source_bytes)

    # 规则：只保留总应收 > 0，且客户余款 <= 0
    eligible = [
        record
        for record in master_records.values()
        if record["总应收"] > 0 and record["客户余款"] <= 0
    ]
    eligible.sort(key=lambda record: record["总应收"], reverse=True)

    if template_bytes:
        wb, ws, header_row, col_map, history, old_names = read_old_followup(template_bytes)
    else:
        wb, ws, header_row = create_blank_template()
        history = {}
        old_names = []
        col_map = {header: idx + 1 for idx, header in enumerate(OUTPUT_HEADERS)}

    old_exact_keys = {normalize_name(name) for name in old_names}
    new_customers = [
        record["客户名称"]
        for record in eligible
        if normalize_name(record["客户名称"]) not in old_exact_keys
    ]

    active_names = {normalize_name(record["客户名称"]) for record in eligible}
    removed_customers = [
        name for name in old_names
        if normalize_name(name) not in active_names
    ]

    data_start_row = header_row + 1
    total_row = data_start_row + len(eligible)

    old_max_row = ws.max_row
    clear_data_area(
        ws,
        data_start_row,
        max(old_max_row, total_row + 3),
    )

    # 复制旧表正文行格式
    if template_bytes and old_max_row >= data_start_row:
        for row_idx in range(data_start_row, total_row + 1):
            copy_row_style(ws, data_start_row, row_idx)

    # 确保表头名称和顺序正确
    for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
        ws.cell(header_row, col_idx).value = header

    new_fill = PatternFill("solid", fgColor="E2F0D9")
    preview_rows = []

    for seq, record in enumerate(eligible, start=1):
        row_idx = data_start_row + seq - 1
        customer = record["客户名称"]

        saved = (
            history.get(normalize_name(customer))
            or history.get(relaxed_name(customer))
            or {}
        )

        values = [
            seq,
            customer,
            record["总应收"],
            saved.get("负责人"),
            saved.get("联系日期"),
            saved.get("可付款时间"),
            saved.get("付款状态"),
            saved.get("下次跟进日期"),
        ]

        for col_idx, value in enumerate(values, start=1):
            ws.cell(row_idx, col_idx).value = value

        ws.cell(row_idx, 3).number_format = '#,##0.00'
        ws.cell(row_idx, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row_idx, 3).alignment = Alignment(horizontal="center", vertical="center")

        # 新客户整行浅绿色标记
        if customer in new_customers:
            for col_idx in range(1, 9):
                ws.cell(row_idx, col_idx).fill = new_fill

        preview_rows.append({
            "序号": seq,
            "客户名称": customer,
            "总应收": record["总应收"],
            "负责人": saved.get("负责人"),
            "联系日期": saved.get("联系日期"),
            "可付款时间": saved.get("可付款时间"),
            "付款状态": saved.get("付款状态"),
            "下次跟进日期": saved.get("下次跟进日期"),
        })

    # 合计行保持用户原表形式：A列“共XX户”，B列“合计”
    ws.cell(total_row, 1).value = f"共{len(eligible)}户"
    ws.cell(total_row, 2).value = "合计"
    ws.cell(total_row, 3).value = f"=SUM(C{data_start_row}:C{total_row - 1})"
    ws.cell(total_row, 3).number_format = '#,##0.00'
    style_total_row(ws, total_row)

    # 日期列格式
    for row_idx in range(data_start_row, total_row):
        ws.cell(row_idx, 5).number_format = "yyyy-mm-dd"
        ws.cell(row_idx, 8).number_format = "yyyy-mm-dd"

    ws.freeze_panes = None
    ws.auto_filter.ref = None

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return {
        "bytes": output.getvalue(),
        "preview": pd.DataFrame(preview_rows),
        "new_customers": new_customers,
        "removed_customers": removed_customers,
        "duplicates": duplicates,
        "skipped_sheets": skipped,
        "customer_count": len(eligible),
        "total_receivable": sum(record["总应收"] for record in eligible),
    }


st.title("💰 应收账款跟进表自动更新")
st.caption(
    "这是独立的应收账款 App，不处理服务站数据。"
    "上传最新《给康总数据》和上一版应收账款跟进表，即可自动生成新表。"
)

with st.sidebar:
    st.subheader("更新规则")
    st.markdown(
        """
- 只保留 **总应收 > 0**
- 只保留 **客户余款 ≤ 0**
- 按总应收从高到低排序
- 自动新增符合条件的新客户
- 自动删除不再符合条件的客户
- 保留人工跟进信息：
  - 负责人
  - 联系日期
  - 可付款时间
  - 付款状态
  - 下次跟进日期
- 新客户整行浅绿色标记
- 自动更新“共XX户”和应收合计
- 不冻结窗口
- 输出文件不保留易损坏的 Excel Table 结构，避免打开时提示修复
        """
    )

source_file = st.file_uploader(
    "① 上传最新《给康总数据》",
    type=["xlsx"],
    help="例如：26年给康总数据_更新至2026-07-27.xlsx",
)

template_file = st.file_uploader(
    "② 上传上一版《应收账款跟进表》",
    type=["xlsx"],
    help="用于保留负责人、联系日期、付款状态等人工填写内容。",
)

if st.button("🚀 开始更新应收账款跟进表", type="primary", use_container_width=True):
    if source_file is None:
        st.error("请先上传最新《给康总数据》。")
    elif template_file is None:
        st.error("请上传上一版《应收账款跟进表》，否则无法保留人工跟进信息。")
    else:
        try:
            with st.spinner("正在读取最新数据、匹配客户并保留跟进信息……"):
                result = build_receivable_output(
                    source_file.getvalue(),
                    template_file.getvalue(),
                )

            c1, c2, c3 = st.columns(3)
            c1.metric("应收客户数", f"{result['customer_count']} 户")
            c2.metric("总应收", f"¥{result['total_receivable']:,.2f}")
            c3.metric("新增客户", f"{len(result['new_customers'])} 户")

            if result["new_customers"]:
                st.success("本次新增：" + "、".join(result["new_customers"]))

            if result["removed_customers"]:
                with st.expander(f"本次移除 {len(result['removed_customers'])} 户"):
                    st.write("、".join(result["removed_customers"]))

            if result["duplicates"]:
                st.warning(
                    "康总数据中发现重复客户名，已采用最后一次出现的数据："
                    + "、".join(sorted(set(result["duplicates"])))
                )

            st.subheader("更新结果预览")
            st.dataframe(
                result["preview"],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "总应收": st.column_config.NumberColumn("总应收", format="%.2f"),
                    "联系日期": st.column_config.DateColumn("联系日期", format="YYYY-MM-DD"),
                    "下次跟进日期": st.column_config.DateColumn("下次跟进日期", format="YYYY-MM-DD"),
                },
            )

            filename = f"应收账款跟进表_更新至{datetime.now():%Y-%m-%d}.xlsx"
            st.download_button(
                "⬇️ 下载更新后的应收账款跟进表",
                data=result["bytes"],
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as exc:
            st.exception(exc)
