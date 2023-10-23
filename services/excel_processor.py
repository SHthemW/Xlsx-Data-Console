# excel_processor.py
import math
import os
import openpyxl
from entities.color import Colors
from entities.command import Keyword
from utilities.process import start_window


class ExcelProcessor:
    def __init__(self, directory: str):
        self.__directory__ = directory

    def search_files(self, field: tuple, filenames, show_detail: bool):
        def find_as(tgt_type: type) -> bool:
            _found = False
            if tgt_type(cell.value) in map(tgt_type, values):
                print(
                    Colors.GREEN + f'在文件{Colors.LIGHTYELLOW_EX}{filename:<30}{Colors.GREEN}的工作表{Colors.LIGHTYELLOW_EX}{sheet:<10}{Colors.GREEN}的第{Colors.LIGHTYELLOW_EX}{cell.row:<5}{Colors.GREEN}行第{Colors.LIGHTYELLOW_EX}{cell.column:<5}{Colors.GREEN}列查找到目标' + Colors.RESET)
                _found = True
                if show_detail:
                    print('详细信息：')
                    for i in range(0, len(row), 2):
                        left = row[i]
                        right = row[i + 1] if i + 1 < len(row) else None
                        print(
                            f"{Colors.GREEN}{worksheet[left.column_letter + '1'].value:<20}: "
                            f"{Colors.LIGHTYELLOW_EX}{str(left.value)[:10] if left.value is not None else '':<10}"
                            f"{Colors.RESET}", end="  ")
                        if not right:
                            continue
                        print(
                            f"{Colors.GREEN}{worksheet[right.column_letter + '1'].value:<20}: "
                            f"{Colors.LIGHTYELLOW_EX}{str(right.value)[:10] if right.value is not None else '':<10}"
                            f"{Colors.RESET}")
                    print("")
            return _found

        column_name, values = field
        found = False
        for filename in os.listdir(self.__directory__):
            if filename.endswith(".xlsx") and not filename.startswith("~$") and filename.lower().replace('.xlsx',
                                                                                                         '') in filenames:
                workbook = openpyxl.load_workbook(os.path.join(self.__directory__, filename))
                for sheet in workbook.sheetnames:
                    worksheet = workbook[sheet]
                    headers = [cell.value for cell in worksheet[1]]
                    column_index = None if column_name is None else (
                        headers.index(column_name) + 1 if column_name in headers else None)
                    if column_index is not None or column_name is None:
                        for row in worksheet.iter_rows(min_row=2):
                            for cell in row:
                                if cell.value is not None and (
                                        (column_index is None) or (cell.col_idx == column_index)):
                                    try:
                                        # 尝试将单元格值和字段都转换为整数进行比较
                                        if find_as(int): found = True
                                    except ValueError:
                                        # 如果转换失败，则按原来的方式进行比较
                                        if find_as(str): found = True
        if not found:
            print(Colors.YELLOW + f"未在任何文件中找到字段'{field}'" + Colors.RESET)

    def update_files(self, old_field: tuple, new_field: tuple, filenames, require_restart: bool = False):
        old_column_name, old_values = old_field
        new_column_name, new_values = new_field
        found = False
        if (old_column_name is None and new_column_name is not None) or (
                old_column_name is not None and new_column_name is None):
            print(Colors.RED + "错误：old_column_name 和 new_column_name 必须同时为空或非空" + Colors.RESET)
            return
        for filename in os.listdir(self.__directory__):
            if filename.endswith(".xlsx") and not filename.startswith("~$") and filename.lower().replace('.xlsx',
                                                                                                         '') in filenames:
                workbook = openpyxl.load_workbook(os.path.join(self.__directory__, filename))
                workbook.save(os.path.join(self.__directory__, filename))

                for sheet in workbook.sheetnames:
                    worksheet = workbook[sheet]
                    headers = [cell.value for cell in worksheet[1]]

                    old_column_index = None if old_column_name is None else (
                        headers.index(old_column_name) + 1 if old_column_name in headers else None)

                    new_column_index = None if new_column_name is None else (
                        headers.index(new_column_name) + 1 if new_column_name in headers else None)

                    # 获取新值行

                    new_value_rows = [row for row in worksheet.iter_rows(min_row=2)
                                      if row[0].row and new_column_index
                                      and str(worksheet.cell(row=row[0].row,
                                                             column=new_column_index).value)
                                      in map(str, new_values)]

                    for row in worksheet.iter_rows(min_row=2):
                        updated = False
                        for cell in row:
                            if old_column_index is not None and new_column_index is not None:
                                # 匹配行
                                if str(worksheet.cell(row=cell.row, column=old_column_index).value) \
                                        in map(str, old_values):
                                    if len(new_value_rows) != len(old_values):
                                        print(Colors.RED + f"错误：需更新的行数与new_values的数量不匹配: "
                                                           f"{len(new_value_rows)}, {len(old_values)}" + Colors.RESET)
                                        return
                                    else:
                                        try:
                                            for old_cell, new_cell in zip(row, new_value_rows[old_values.index(str(
                                                    worksheet.cell(row=cell.row, column=old_column_index).value))]):
                                                if old_cell.column == old_column_index:
                                                    continue
                                                old_cell.value = new_cell.value
                                                updated = True
                                            found = True
                                        except ValueError:
                                            for old_cell, new_cell in zip(row,
                                                                          new_value_rows[old_values.index(int(
                                                                              worksheet.cell(row=cell.row,
                                                                                             column=old_column_index).value))]):
                                                if old_cell.column == old_column_index:
                                                    continue
                                                old_cell.value = new_cell.value
                                                updated = True
                                            found = True

                            elif old_column_index is None and new_column_index is None:
                                # 匹配数值
                                if cell.value is not None:  # 添加了处理空值的判断语句
                                    try:
                                        # 尝试将单元格值和字段都转换为整数进行比较
                                        if int(cell.value) in map(int, old_values):
                                            cell.value = int(new_values[0]) if len(new_values) == 1 else int(
                                                new_values[old_values.index(int(cell.value))])
                                            updated = True
                                            found = True
                                    except ValueError:
                                        # 如果转换失败，则按原来的方式进行比较
                                        if str(cell.value) in map(str, old_values):
                                            cell.value = str(new_values[old_values.index(str(cell.value))])
                                            updated = True
                                            found = True
                        if updated:
                            print(Colors.GREEN + f'在文件{Colors.LIGHTYELLOW_EX}{filename:<30}{Colors.GREEN}的工作表'
                                                 f'{Colors.LIGHTYELLOW_EX}{sheet:<10}{Colors.GREEN}的第{Colors.LIGHTYELLOW_EX}'
                                                 f'{cell.row:<5}{Colors.GREEN}行更新了目标' + Colors.RESET)
                workbook.save(os.path.join(self.__directory__, filename))
                if require_restart:
                    start_window(os.path.join(self.__directory__, filename))
                    print(Colors.YELLOW + f"被关闭的文件已操作完毕. 启动!" + Colors.RESET)
        if not found:
            print(Colors.YELLOW + f"未在任何文件中找到字段'{old_field}'" + Colors.RESET)

    def parse_filenames(self, filenames_str):

        filenames_str = filenames_str.upper()
        all_files = [f.lower().replace('.xlsx', '') for f in os.listdir(self.__directory__) if
                     f.endswith(".xlsx") and not f.startswith("~$")]
        if Keyword.ALL.upper() in filenames_str.split():
            return all_files
        elif Keyword.EXCEPT.upper() in filenames_str.split():
            _, except_files_str = filenames_str.split(Keyword.EXCEPT.upper())
            except_files = except_files_str.lower().split()
            result = [f for f in all_files if f not in except_files]
            return result
        else:
            result = filenames_str.lower().split()
            return result

    @staticmethod
    def get_pure_filenames(filenames_str: str):
        return [name + (".xlsx" if ".xlsx" not in name else "")
                for name in filenames_str.split() if not Keyword.is_keyword(name)]
