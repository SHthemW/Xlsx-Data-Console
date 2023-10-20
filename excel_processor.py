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
            return [f for f in all_files if f not in except_files]
        else:
            return filenames_str.lower().split()
