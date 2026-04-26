import re
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


_NUM_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
_MIN_ANCHOR_WORD_LEN = 2
_TWO_DECIMAL_PLACES = Decimal("0.01")
_TWO_DECIMAL_PLACES_EXPONENT = -2


def norm_text(value: str) -> str:
    normalized = value.casefold().replace("ё", "е")
    normalized = normalized.replace("\xa0", " ")
    normalized = re.sub(r"[^\w]+", " ", normalized, flags=re.UNICODE)
    return " ".join(normalized.split())


def match_anchor(cell_value: str, anchor: str) -> bool:
    if not isinstance(cell_value, str):
        return False

    cell_norm = norm_text(cell_value)
    anchor_words = [
        word for word in norm_text(anchor).split() if len(word) >= _MIN_ANCHOR_WORD_LEN
    ]
    if not anchor_words:
        return False
    cell_words = cell_norm.split()

    for anchor_word in anchor_words:
        if not any(word.startswith(anchor_word) for word in cell_words):
            return False

    return True


def find_header_columns(
    ws: Worksheet, header_anchors: dict[str, str]
) -> tuple[int, dict[str, int]] | None:
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        found_columns: dict[str, int] = {}
        for cell in row:
            cell_value = cell.value
            if not isinstance(cell_value, str):
                continue

            for key, anchor in header_anchors.items():
                if key in found_columns:
                    continue
                if match_anchor(cell_value, anchor):
                    found_columns[key] = cell.column

        if len(found_columns) == len(header_anchors):
            return row_idx, found_columns
    return None


def find_metric_column_idx(
    ws: Worksheet,
    *,
    item_anchor: str,
    metric_anchor: str | tuple[str, ...],
    header_scan_depth: int = 2,
) -> int | None:
    header_data = find_header_columns(ws, {"item_col_idx": item_anchor})
    if header_data is None:
        return None

    header_row_idx, columns = header_data
    item_col_idx = columns.get("item_col_idx")
    if item_col_idx is None:
        return None

    scan_end_row_idx = min(ws.max_row, header_row_idx + header_scan_depth)
    anchors = metric_anchor if isinstance(metric_anchor, tuple) else (metric_anchor,)
    for col_idx in range(item_col_idx + 1, ws.max_column + 1):
        for scan_row_idx in range(header_row_idx, scan_end_row_idx + 1):
            header_value = ws.cell(row=scan_row_idx, column=col_idx).value
            if not isinstance(header_value, str):
                continue
            if any(match_anchor(header_value, anchor) for anchor in anchors):
                return col_idx
    return None


def is_stop_row_by_anchor_match(value: str, stop_anchors: tuple[str, ...]) -> bool:
    return any(match_anchor(value, anchor) for anchor in stop_anchors)


def advance_empty_rows_series(
    empty_rows_in_series: int, *, is_empty: bool, max_empty_rows_in_series: int
) -> tuple[int, bool]:
    if is_empty:
        next_count = empty_rows_in_series + 1
        return next_count, next_count >= max_empty_rows_in_series
    return 0, False


def to_decimal(value: Any) -> Decimal | None:
    result: Decimal | None = None

    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, int | float):
        result = Decimal(str(value))
    elif isinstance(value, str):
        stripped = value.strip()
        if stripped:
            if stripped.startswith("#"):
                return None
            stripped = stripped.replace("\xa0", " ").replace(" ", "")
            match = _NUM_RE.search(stripped)
            if match:
                number = match.group(0).replace(",", ".")
                try:
                    result = Decimal(number)
                except InvalidOperation:
                    result = None

    if result is None:
        return None
    if result.as_tuple().exponent < _TWO_DECIMAL_PLACES_EXPONENT:
        return result.quantize(_TWO_DECIMAL_PLACES, rounding=ROUND_HALF_UP)
    return result


def find_anchor_row_idx(ws: Worksheet, anchor: str) -> int | None:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and match_anchor(cell.value, anchor):
                return cell.row
    return None


def extract_amount_from_row(ws: Worksheet, row_idx: int, anchor: str) -> Decimal | None:
    rows = list(ws.iter_rows(min_row=row_idx, max_row=row_idx))
    row_cells = list(rows[0]) if rows else []

    for cell in reversed(row_cells):
        cell_value = cell.value
        if isinstance(cell_value, str) and match_anchor(cell_value, anchor):
            continue
        decimal_value = to_decimal(cell_value)
        if decimal_value is not None:
            return decimal_value
    return None


def extract_amount_by_anchor(
    ws: Worksheet,
    anchor: str,
    *,
    preferred_col_idx: int | None = None,
) -> Decimal | None:
    row_idx = find_anchor_row_idx(ws, anchor)
    if row_idx is None:
        return None
    if preferred_col_idx is not None:
        preferred_value = ws.cell(row=row_idx, column=preferred_col_idx).value
        if not (
            isinstance(preferred_value, str) and match_anchor(preferred_value, anchor)
        ):
            preferred_decimal = to_decimal(preferred_value)
            if preferred_decimal is not None:
                return preferred_decimal
    return extract_amount_from_row(ws, row_idx, anchor)


def resolve_merged_columns_range(
    ws: Worksheet,
    *,
    row_idx: int,
    col_idx: int,
) -> tuple[int, int]:
    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_row <= row_idx <= merged_range.max_row
            and merged_range.min_col <= col_idx <= merged_range.max_col
        ):
            return merged_range.min_col, merged_range.max_col
    return col_idx, col_idx


def find_anchor_cell_in_row_window(
    ws: Worksheet,
    *,
    anchor: str,
    start_row_idx: int,
    scan_depth: int,
) -> tuple[int, int] | None:
    scan_end_row_idx = min(ws.max_row, start_row_idx + scan_depth)
    for row_idx in range(start_row_idx, scan_end_row_idx + 1):
        for col_idx in range(1, ws.max_column + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if not isinstance(value, str):
                continue
            if match_anchor(value, anchor):
                return row_idx, col_idx
    return None


def find_anchor_columns_range_in_row_window(
    ws: Worksheet,
    *,
    anchor: str,
    start_row_idx: int,
    scan_depth: int,
) -> tuple[int, int] | None:
    anchor_cell = find_anchor_cell_in_row_window(
        ws,
        anchor=anchor,
        start_row_idx=start_row_idx,
        scan_depth=scan_depth,
    )
    if anchor_cell is None:
        return None

    row_idx, col_idx = anchor_cell
    return resolve_merged_columns_range(ws, row_idx=row_idx, col_idx=col_idx)


def is_service_subheader(
    value: str,
    service_anchors: tuple[str, ...],
) -> bool:
    if not isinstance(value, str):
        return True

    normalized_value = norm_text(value)
    if not normalized_value:
        return True

    return any(match_anchor(normalized_value, anchor) for anchor in service_anchors)


def _find_sheet_by_normalized_name(
    wb: Workbook,
    *,
    expected_norm: str,
    excluded_sheet_names: set[str],
) -> Worksheet | None:
    for name in wb.sheetnames:
        if name in excluded_sheet_names:
            continue
        if norm_text(name) == expected_norm:
            return wb[name]
    return None


def _find_sheet_by_anchor_name_match(
    wb: Workbook,
    *,
    expected_name: str,
    excluded_sheet_names: set[str],
) -> Worksheet | None:
    for name in wb.sheetnames:
        if name in excluded_sheet_names:
            continue
        if match_anchor(name, expected_name) and match_anchor(expected_name, name):
            return wb[name]
    return None


def _find_sheet_by_anchor_fallback(
    wb: Workbook,
    *,
    anchor_fallback: str,
    excluded_sheet_names: set[str],
) -> Worksheet | None:
    for name in wb.sheetnames:
        if name in excluded_sheet_names:
            continue
        ws = wb[name]
        row_idx = find_anchor_row_idx(ws, anchor_fallback)
        if row_idx is not None:
            return ws
    return None


def select_sheet(
    wb: Workbook,
    expected_name: str,
    anchor_fallback: str | None,
    *,
    excluded_sheet_names: set[str] | None = None,
) -> Worksheet | None:
    excluded_sheet_names = excluded_sheet_names or set()
    expected_norm = norm_text(expected_name)

    ws_by_normalized_name = _find_sheet_by_normalized_name(
        wb,
        expected_norm=expected_norm,
        excluded_sheet_names=excluded_sheet_names,
    )
    if ws_by_normalized_name is not None:
        return ws_by_normalized_name

    ws_by_anchor_name = _find_sheet_by_anchor_name_match(
        wb,
        expected_name=expected_name,
        excluded_sheet_names=excluded_sheet_names,
    )
    if ws_by_anchor_name is not None:
        return ws_by_anchor_name

    if anchor_fallback:
        return _find_sheet_by_anchor_fallback(
            wb,
            anchor_fallback=anchor_fallback,
            excluded_sheet_names=excluded_sheet_names,
        )

    return None
