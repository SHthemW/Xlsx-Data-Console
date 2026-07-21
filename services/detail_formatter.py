from typing import Iterable, Sequence, Tuple

from tabulate import tabulate
from wcwidth import wcswidth, wrap

from entities.color import Colors
from utilities.terminal import get_terminal_width


WIDE_DETAILS_PER_ROW = 2
COLUMNS_PER_DETAIL = 3
TABLE_COLUMN_GAP = 2
COLON_WIDTH = 1
VALUE_MIN_WIDTH = 10
FIELD_NAME_MIN_WIDTH = 12
NARROW_LAYOUT_MIN_WIDTH = 20
WIDE_LAYOUT_MIN_WIDTH = (
    WIDE_DETAILS_PER_ROW * (FIELD_NAME_MIN_WIDTH + COLON_WIDTH + VALUE_MIN_WIDTH)
    + (WIDE_DETAILS_PER_ROW * COLUMNS_PER_DETAIL - 1) * TABLE_COLUMN_GAP
)


def _style_field(
    field_name: object,
    value: object,
    highlights: Tuple[str, ...],
) -> Tuple[str, str, str]:
    field_name_text = str(field_name or "")
    value_text = str(value)
    color = (
        Colors.RED
        if any(target in value_text for target in highlights)
        else Colors.LIGHTYELLOW_EX
    )
    return (
        f"{Colors.GREEN}{field_name_text}{Colors.RESET}",
        f"{Colors.GREEN}:{Colors.RESET}",
        f"{color}{value_text}{Colors.RESET}",
    )


def _get_text_width(value: object) -> int:
    lines = str(value).expandtabs().splitlines() or [""]
    return max(1, max(wcswidth(line) for line in lines))


def _grow_widths(widths: list, targets: Sequence[int], available_width: int) -> int:
    while available_width > 0:
        expandable_columns = [
            index
            for index, width in enumerate(widths)
            if width < targets[index]
        ]
        if not expandable_columns:
            break

        width_share = max(1, available_width // len(expandable_columns))
        for index in expandable_columns:
            growth = min(
                width_share,
                targets[index] - widths[index],
                available_width,
            )
            widths[index] += growth
            available_width -= growth
            if available_width == 0:
                break

    return available_width


def _get_column_widths(
    fields: Sequence[Tuple[object, object]],
    terminal_width: int,
    details_per_row: int,
) -> Sequence[int]:
    column_count = details_per_row * COLUMNS_PER_DETAIL
    gap_width = (column_count - 1) * TABLE_COLUMN_GAP
    content_width = terminal_width - gap_width - details_per_row * COLON_WIDTH
    field_name_targets = [
        max(_get_text_width(field_name) for field_name, _ in fields[index::details_per_row])
        for index in range(details_per_row)
    ]
    value_targets = [
        max(_get_text_width(value) for _, value in fields[index::details_per_row])
        for index in range(details_per_row)
    ]
    field_name_widths = [
        min(FIELD_NAME_MIN_WIDTH, target)
        for target in field_name_targets
    ]
    value_widths = [
        min(VALUE_MIN_WIDTH, target)
        for target in value_targets
    ]

    if sum(field_name_widths) + sum(value_widths) <= content_width:
        available_width = content_width - sum(field_name_widths) - sum(value_widths)
        available_width = _grow_widths(
            field_name_widths,
            field_name_targets,
            available_width,
        )
        _grow_widths(value_widths, value_targets, available_width)
    else:
        field_width, remainder = divmod(content_width, details_per_row)
        field_name_widths = []
        value_widths = []

        for index in range(details_per_row):
            available_width = field_width + (1 if index < remainder else 0)
            value_width = max(1, available_width // 3)
            field_name_widths.append(max(1, available_width - value_width))
            value_widths.append(value_width)

    widths = []
    for index in range(details_per_row):
        widths.extend((field_name_widths[index], COLON_WIDTH, value_widths[index]))

    return widths


def format_detail_table(
    fields: Sequence[Tuple[object, object]],
    highlighted_values: Iterable[object],
    terminal_width: int = None,
) -> str:
    """按当前终端宽度格式化支持中文宽字符和 ANSI 颜色的详情表格。"""
    if not fields:
        return ""

    if terminal_width is None:
        terminal_width = get_terminal_width()
    terminal_width = max(1, terminal_width)
    highlights = tuple(map(str, highlighted_values))
    styled_fields = [
        _style_field(field_name, value, highlights)
        for field_name, value in fields
    ]

    if terminal_width < NARROW_LAYOUT_MIN_WIDTH:
        stacked_lines = [
            line
            for field_name, colon, value in styled_fields
            for line in wrap(f"{field_name}{colon} {value}", terminal_width)
        ]
        return "\n".join(stacked_lines)

    details_per_row = (
        min(WIDE_DETAILS_PER_ROW, len(fields))
        if terminal_width >= WIDE_LAYOUT_MIN_WIDTH
        else 1
    )
    column_count = details_per_row * COLUMNS_PER_DETAIL
    table_rows = []

    for index in range(0, len(styled_fields), details_per_row):
        table_row = []
        for field in styled_fields[index:index + details_per_row]:
            table_row.extend(field)

        table_row.extend([""] * (column_count - len(table_row)))
        table_rows.append(table_row)

    return tabulate(
        table_rows,
        tablefmt="plain",
        colalign=("left",) * column_count,
        disable_numparse=True,
        maxcolwidths=_get_column_widths(fields, terminal_width, details_per_row),
    )
