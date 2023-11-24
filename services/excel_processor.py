# excel_processor.py
import os
import openpyxl
from entities.color import Colors
from entities.command import Keyword
from utilities.local import get_output_space_offset, get_chinese_char_comp, get_directory
from utilities.process import start_window


class ExcelProcessor:
    def __init__(self, directory: str):
        self.__directory = directory

    def search_files(self, field: tuple, filenames, show_detail: bool) -> str:

        def get_color(s, key_val) -> str:
            value_str = str(s)
            for tgt_value in map(str, key_val):
                if tgt_value in value_str:
                    return Colors.RED
            return Colors.LIGHTYELLOW_EX

        def find_as_str() -> bool:
            _found = False
            if any(tgt_value in str(cell.value) for tgt_value in map(str, values)):
                print(
                    Colors.GREEN + f'在文件{Colors.LIGHTYELLOW_EX}{filename:<30}{Colors.GREEN}的工作表{Colors.LIGHTYELLOW_EX}{sheet:<10}{Colors.GREEN}的第{Colors.LIGHTYELLOW_EX}{cell.row:<5}{Colors.GREEN}行第{Colors.LIGHTYELLOW_EX}{cell.column:<5}{Colors.GREEN}列查找到目标' + Colors.RESET)
                _found = True
                if show_detail:
                    print('详细信息：')
                    for i in range(0, len(row), 2):
                        # calc val
                        left = row[i]
                        right = row[i + 1] if i + 1 < len(row) else None
                        # print
                        left_color = get_color(left.value, values)
                        print(
                            f"{Colors.GREEN}{worksheet[left.column_letter + '1'].value or '':<20}: "
                            f"{left_color}{str(left.value)[:10] or '':<{self.pad_len(left.value, 10)}}"
                            f"{Colors.RESET}", end="   ")
                        if not right:
                            continue
                        right_color = get_color(right.value, values)
                        print(
                            f"{Colors.GREEN}{worksheet[right.column_letter + '1'].value or '':<20}: "
                            f"{right_color}{str(right.value)[:10] or '':<{self.pad_len(right.value, 10)}}"
                            f"{Colors.RESET}")
                    print("")
            return _found

        column_name, values = field
        ret_filename = None
        found = False
        for filename in os.listdir(self.__directory):
            if filename.endswith(".xlsx") and not filename.startswith("~$") and filename.lower().replace('.xlsx',
                                                                                                         '') in filenames:
                workbook = openpyxl.load_workbook(os.path.join(self.__directory, filename))
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
                                    # 尝试将单元格值和字段都转换为str进行比较
                                    if find_as_str():
                                        found = True
                                        ret_filename = filename
        if not found:
            print(Colors.YELLOW + f"未在任何文件中找到字段'{field}'" + Colors.RESET)
        return ret_filename

    def update_files(self, old_field: tuple, new_field: tuple, filenames, require_restart: bool = False) -> str:
        def update_rows(tgt_type: type) -> bool:
            success = False
            for old_cell, new_cell in zip(row, new_value_rows[old_values.index(tgt_type(
                    worksheet.cell(row=cell.row, column=old_column_index).value))]):
                if old_cell.column == old_column_index:
                    continue
                old_cell.value = new_cell.value
                success = True
            return success

        def update_vals(tgt_type: type) -> bool:
            success = False
            if tgt_type(cell.value) in map(tgt_type, old_values):
                cell.value = tgt_type(new_values[0]) if len(new_values) == 1 else tgt_type(
                    new_values[old_values.index(tgt_type(cell.value))])
                success = True
            return success

        old_column_name, old_values = old_field
        new_column_name, new_values = new_field
        found = False
        ret_filename = None
        if (old_column_name is None and new_column_name is not None) or (
                old_column_name is not None and new_column_name is None):
            print(Colors.RED + "错误：old_column_name 和 new_column_name 必须同时为空或非空" + Colors.RESET)
            return ret_filename
        for filename in os.listdir(self.__directory):
            if filename.endswith(".xlsx") and not filename.startswith("~$") and filename.lower().replace('.xlsx',
                                                                                                         '') in filenames:
                workbook = openpyxl.load_workbook(os.path.join(self.__directory, filename))
                workbook.save(os.path.join(self.__directory, filename))

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
                                        return ret_filename
                                    try:
                                        if update_rows(str): updated, found = True, True
                                    except ValueError:
                                        if update_rows(int): updated, found = True, True

                            elif old_column_index is None and new_column_index is None:
                                # 匹配数值
                                if cell.value is not None:
                                    try:
                                        if update_vals(int): updated, found = True, True
                                    except ValueError:
                                        if update_vals(str): updated, found = True, True
                        if updated:
                            print(Colors.GREEN + f'在文件{Colors.LIGHTYELLOW_EX}{filename:<30}{Colors.GREEN}的工作表'
                                                 f'{Colors.LIGHTYELLOW_EX}{sheet:<10}{Colors.GREEN}的第{Colors.LIGHTYELLOW_EX}'
                                                 f'{cell.row:<5}{Colors.GREEN}行更新了目标' + Colors.RESET)
                            ret_filename = filename
                workbook.save(os.path.join(self.__directory, filename))
                if require_restart:
                    start_window(os.path.join(self.__directory, filename))
        if not found:
            print(Colors.YELLOW + f"未在任何文件中找到字段'{old_field}'" + Colors.RESET)
        return ret_filename

    def parse_filenames(self, filenames_str):

        filenames_str = filenames_str.upper()
        all_files = [f.lower().replace('.xlsx', '') for f in os.listdir(self.__directory) if
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

    def get_pure_filenames(self, filenames_str: str):
        def find_matching_filename(input_str):
            for filename in os.listdir(self.__directory):
                if filename.split(".")[0].lower() == input_str:
                    return filename
            return None

        return [(find_matching_filename(name) if find_matching_filename(name) else
                 (name + (".xlsx" if ".xlsx" not in name else "")))
                for name in filenames_str.split() if not Keyword.is_keyword(name)]

    @staticmethod
    def pad_len(string, length):
        def ch_count(check_str):
            count = 0
            for ch in check_str:
                if '\u4e00' <= ch <= '\u9fff':
                    count += 1
            return count

        string = str(string)
        result = (length
                  - len(string.encode('GBK'))
                  + len(string)
                  + int(ch_count(string) * get_chinese_char_comp())
                  + get_output_space_offset())
        return max(0, result)


if __name__ == '__main__':
    print(ExcelProcessor(get_directory()).get_pure_filenames(input("test pure: ")))
