# excel_processor.py
import os
import openpyxl
from _config import Colors


class ExcelProcessor:
    def __init__(self, directory):
        self.__directory__ = directory

    def search_files(self, field, filenames):
        found = False
        for filename in os.listdir(self.__directory__):
            if filename.endswith(".xlsx") and not filename.startswith("~$") and filename.lower().replace('.xlsx',
                                                                                                         '') in filenames:
                try:
                    workbook = openpyxl.load_workbook(os.path.join(self.__directory__, filename))
                    for sheet in workbook.sheetnames:
                        worksheet = workbook[sheet]
                        for row in worksheet.iter_rows():
                            for cell in row:
                                if cell.value is not None:
                                    try:
                                        # 尝试将单元格值和字段都转换为整数进行比较
                                        if int(cell.value) == int(field):
                                            print(
                                                Colors.GREEN + f'在文件{Colors.LIGHTYELLOW_EX}{filename:<30}{Colors.GREEN}的工作表{Colors.LIGHTYELLOW_EX}{sheet:<10}{Colors.GREEN}的第{Colors.LIGHTYELLOW_EX}{cell.row:<5}{Colors.GREEN}行第{Colors.LIGHTYELLOW_EX}{cell.column:<5}{Colors.GREEN}列查找到目标' + Colors.RESET)
                                            found = True
                                    except ValueError:
                                        # 如果转换失败，则按原来的方式进行比较
                                        pass
                except Exception as e:
                    print(Colors.RED + f"处理文件{filename}时发生错误: {str(e)}" + Colors.RESET)
        if not found:
            print(Colors.YELLOW + f"未在任何文件中找到字段'{field}'" + Colors.RESET)

    def update_files(self, old_field, new_field, filenames):
        found = False
        for filename in os.listdir(self.__directory__):
            if filename.endswith(".xlsx") and not filename.startswith("~$") and filename.lower().replace('.xlsx',
                                                                                                         '') in filenames:
                try:
                    workbook = openpyxl.load_workbook(os.path.join(self.__directory__, filename))
                    for sheet in workbook.sheetnames:
                        worksheet = workbook[sheet]
                        for row in worksheet.iter_rows():
                            for cell in row:
                                if cell.value is not None:
                                    try:
                                        # 尝试将单元格值和字段都转换为整数进行比较
                                        if int(cell.value) == int(old_field):
                                            cell.value = int(new_field)
                                            print(
                                                Colors.GREEN + f'在文件{Colors.LIGHTYELLOW_EX}{filename:<30}{Colors.GREEN}的工作表{Colors.LIGHTYELLOW_EX}{sheet:<10}{Colors.GREEN}的第{Colors.LIGHTYELLOW_EX}{cell.row:<5}{Colors.GREEN}行第{Colors.LIGHTYELLOW_EX}{cell.column:<5}{Colors.GREEN}列更新了目标' + Colors.RESET)
                                            found = True
                                    except ValueError:
                                        # 如果转换失败，则按原来的方式进行比较
                                        pass
                    workbook.save(os.path.join(self.__directory__, filename))
                except Exception as e:
                    print(Colors.RED + f"处理文件{filename}时发生错误: {str(e)}" + Colors.RESET)
        if not found:
            print(Colors.YELLOW + f"未在任何文件中找到字段'{old_field}'" + Colors.RESET)

    def parse_filenames(self, filenames_str):
        filenames_str = filenames_str.upper()
        all_files = [f.lower().replace('.xlsx', '') for f in os.listdir(self.__directory__) if
                     f.endswith(".xlsx") and not f.startswith("~$")]
        if 'ALL' in filenames_str.split():
            return all_files
        elif 'EXCEPT' in filenames_str.split():
            _, except_files_str = filenames_str.split('EXCEPT')
            except_files = except_files_str.lower().split()
            result = [f for f in all_files if f not in except_files]
            return result
        else:
            result = filenames_str.lower().split()
            return result

    def copy_row(self, key_column, source_key, target_key, filenames):
        for filename in os.listdir(self.__directory__):
            if filename.endswith(".xlsx") and not filename.startswith("~$") and filename.lower().replace('.xlsx',
                                                                                                         '') in filenames:
                try:
                    workbook = openpyxl.load_workbook(os.path.join(self.__directory__, filename))
                    for sheet in workbook.sheetnames:
                        worksheet = workbook[sheet]
                        source_row = None
                        target_row = None
                        key_column_index = None
                        # 遍历标题行，找到主键列
                        for cell in worksheet[1]:
                            if str(cell.value) == str(key_column):
                                key_column_index = cell.column
                                break
                        if key_column_index is None:
                            print(Colors.YELLOW + f"在文件{filename}的表{sheet}中，没有找到列名为{key_column}的列。" + Colors.RESET)
                            continue
                        # 在主键列中查找源行和目标行
                        for row in worksheet.iter_rows(min_row=2):  # 从第二行开始遍历
                            if str(row[key_column_index - 1].value) == str(source_key):
                                source_row = row
                            if str(row[key_column_index - 1].value) == str(target_key):
                                target_row = row
                        if source_row and target_row:
                            for source_cell, target_cell in zip(source_row[1:], target_row[1:]):  # 从第二列开始复制
                                target_cell.value = source_cell.value
                            print(
                                Colors.GREEN + f"在文件{filename}中，已将主键为{source_key}的行复制到主键为{target_key}的行。" + Colors.RESET)
                        elif not source_row:
                            print(Colors.YELLOW + f"在文件{filename}中，没有找到主键为{source_key}的行。" + Colors.RESET)
                        elif not target_row:
                            print(Colors.YELLOW + f"在文件{filename}中，没有找到主键为{target_key}的行。" + Colors.RESET)
                    workbook.save(os.path.join(self.__directory__, filename))
                except Exception as e:
                    print(Colors.RED + f"处理文件{filename}时发生错误: {str(e)}" + Colors.RESET)




