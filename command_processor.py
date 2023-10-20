# command_processor.py
import os
from utils import Colors


class CommandProcessor:
    def __init__(self, excel_processor):
        self.__excel_proc__ = excel_processor

    def process_command(self, command) -> bool:
        if 'q' in command or 'exit' in command:
            return False
        elif 'clean' in command:
            self.clean_console()
            return True
        elif 'help' in command:
            self.print_help()
            return True

        if command[0].lower() == 'get':
            fields = command[1:-2] if 'in' in command or 'except' in command else command[1:]
            filenames_str = ' '.join(command[-2:]) if 'in' in command or 'except' in command else 'ALL'
            for field in fields:
                for f in self.parse_field(field):
                    print(f"\n正在查找字段'{f}'的结果：")
                    self.__excel_proc__.search_files(f, self.__excel_proc__.parse_filenames(filenames_str))

        elif command[0].lower() == 'update':
            if len(command) < 5 or (
                    command[2].lower() != 'to' and (command[4].lower() != 'in' and command[4].lower() != 'except')):
                print(Colors.RED + "更新操作必须指定作用域" + Colors.RESET)
                return True
            old_field, new_field, filenames_str = command[1], command[3], ' '.join(command[4:])
            for old_f in self.parse_field(old_field):
                for new_f in self.parse_field(new_field):
                    print(f"\n正在更新字段'{old_f}'至'{new_f}'的结果：")
                    self.__excel_proc__.update_files(old_f, new_f, self.__excel_proc__.parse_filenames(filenames_str))

        elif command[0].lower() == 'copyrow':
            if len(command) < 6 or (
                    command[3].lower() != 'to' and (command[5].lower() != 'in' and command[5].lower() != 'except')):
                print(Colors.RED + "复制操作必须指定作用域" + Colors.RESET)
                return True
            key_column, source_keys, target_keys, filenames_str = command[1], self.parse_field(
                command[2]), self.parse_field(command[4]), ' '.join(command[5:])
            if len(source_keys) != len(target_keys):
                print(Colors.RED + "源行和目标行的数量不匹配" + Colors.RESET)
                return True
            for source_key, target_key in zip(source_keys, target_keys):
                print(f"\n正在将主键'{source_key}'复制到主键'{target_key}'的结果：")
                self.__excel_proc__.copy_row(key_column, source_key, target_key,
                                             self.__excel_proc__.parse_filenames(filenames_str))

        else:
            print(Colors.RED + "无法识别的命令" + Colors.RESET)
        return True

    @staticmethod
    def parse_field(field):
        if '-' in field:
            try:
                start, end = map(int, field.split('-'))
                return list(range(start, end + 1))
            except ValueError:
                print(Colors.RED + f"输入'{field}'无法解析为数字区间，将作为字符串处理" + Colors.RESET)
                return [field]
        else:
            return [field]

    @staticmethod
    def print_help():
        print(f"\n{Colors.GREEN}以下是所有可用的命令：{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}get{Colors.LIGHTYELLOW_EX} [field1] [field2] ... {Colors.LIGHTMAGENTA_EX}in{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在指定的文件中查找一个或多个字段。{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}get{Colors.LIGHTYELLOW_EX} [field1] [field2] ... {Colors.LIGHTMAGENTA_EX}except{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在除指定文件外的所有文件中查找一个或多个字段。{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}update{Colors.LIGHTYELLOW_EX} [old_field] {Colors.LIGHTMAGENTA_EX}to{Colors.LIGHTYELLOW_EX} [new_field] {Colors.LIGHTMAGENTA_EX}in{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在指定的文件中将一个字段的值更新为另一个值。{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}update{Colors.LIGHTYELLOW_EX} [old_field] {Colors.LIGHTMAGENTA_EX}to{Colors.LIGHTYELLOW_EX} [new_field] {Colors.LIGHTMAGENTA_EX}except{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在除指定文件外的所有文件中将一个字段的值更新为另一个值。{Colors.RESET}")
        print(f"\n{Colors.LIGHTMAGENTA_EX}exit{Colors.RESET}")
        print(f"{Colors.RESET}  - 退出程序。{Colors.RESET}")
        print(f"\n{Colors.LIGHTMAGENTA_EX}clean{Colors.RESET}")
        print(f"{Colors.RESET}  - 清理控制台的当前显示内容。{Colors.RESET}")

    @staticmethod
    def clean_console():
        os.system('cls' if os.name == 'nt' else 'clear')
