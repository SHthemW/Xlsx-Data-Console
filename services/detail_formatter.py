from typing import Iterable, Sequence, Tuple

from tabulate import tabulate

from entities.color import Colors


DETAILS_PER_ROW = 2
TABLE_COLUMN_COUNT = 6
VALUE_MAX_WIDTH = 10


def format_detail_table(
    fields: Sequence[Tuple[object, object]],
    highlighted_values: Iterable[object],
) -> str:
    """将字段详情格式化为支持中文宽字符和 ANSI 颜色的无边框表格。"""
    highlights = tuple(map(str, highlighted_values))
    table_rows = []

    for index in range(0, len(fields), DETAILS_PER_ROW):
        table_row = []
        for field_name, value in fields[index:index + DETAILS_PER_ROW]:
            field_name_text = str(field_name or "")
            value_text = str(value)
            color = (
                Colors.RED
                if any(target in value_text for target in highlights)
                else Colors.LIGHTYELLOW_EX
            )
            table_row.extend((
                f"{Colors.GREEN}{field_name_text}{Colors.RESET}",
                f"{Colors.GREEN}:{Colors.RESET}",
                f"{color}{value_text}{Colors.RESET}",
            ))

        table_row.extend([""] * (TABLE_COLUMN_COUNT - len(table_row)))
        table_rows.append(table_row)

    return tabulate(
        table_rows,
        tablefmt="plain",
        colalign=("left",) * TABLE_COLUMN_COUNT,
        disable_numparse=True,
        maxcolwidths=[None, None, VALUE_MAX_WIDTH, None, None, VALUE_MAX_WIDTH],
    )
