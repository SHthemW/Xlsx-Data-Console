import os
from entities.color import Colors
from entities.command import Command, Keyword


class InstructProcessor:
    def __init__(self, excel_processor, expression_processor):
        self.__excel_proc__ = excel_processor
        self.__exp_proc__ = expression_processor

    def process_command(self, command: str) -> bool:
        # basic commands
        if Command.EXIT in command:
            return False
        elif Command.CLEAN in command:
            self.clean_console()
            return True
        elif Command.HELP in command:
            self.print_help()
            return True

        # find command used to query row in files.
        if command[0].lower() == Command.FIND:
            fields = command[1:-2] if Keyword.IN in command or Keyword.EXCEPT in command else command[1:]
            filenames_str = ' '.join(
                command[-2:]) if Keyword.IN in command or Keyword.EXCEPT in command else Keyword.ALL

            for field in fields:
                col, val = self.__exp_proc__.parse_chunk_exp(field)
                print(f"\n正在查找字段'{col}-{val}'：")
                self.__excel_proc__.search_files((col, val), self.__excel_proc__.parse_filenames(filenames_str))

        # update command used to update rows in files
        # (must assign the action scope)
        elif command[0].lower() == Command.UPDATE:
            if len(command) < 5 or (
                    command[2].lower() != Keyword.TO and (
                    command[4].lower() != Keyword.IN and command[4].lower() != Keyword.EXCEPT)):
                print(Colors.RED + "更新操作必须指定作用域" + Colors.RESET)
                return True

            old_field, new_field, filenames_str = command[1], command[3], ' '.join(command[4:])
            old_col, old_val = self.__exp_proc__.parse_chunk_exp(old_field)
            new_col, new_val = self.__exp_proc__.parse_chunk_exp(new_field)

            print(f"\n正在更新字段'{old_col}-{old_val}'至'{new_val}'：")
            self.__excel_proc__.update_files((old_col, old_val), (new_col, new_val),
                                             self.__excel_proc__.parse_filenames(filenames_str))

        # create command used to create new rows in files.
        # (must assign the action scope)
        elif command[0].lower() == Command.CREATE:
            return True

        else:
            print(Colors.RED + "无法识别的命令" + Colors.RESET)
        return True

    @staticmethod
    def print_help():
        print(f"\n{Colors.GREEN}以下是所有可用的命令：{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}{Command.FIND}{Colors.LIGHTYELLOW_EX} [field1] [field2] ... {Colors.LIGHTMAGENTA_EX}in{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在指定的文件中查找一个或多个字段。{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}{Command.FIND}{Colors.LIGHTYELLOW_EX} [field1] [field2] ... {Colors.LIGHTMAGENTA_EX}except{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在除指定文件外的所有文件中查找一个或多个字段。{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}{Command.UPDATE}{Colors.LIGHTYELLOW_EX} [old_field] {Colors.LIGHTMAGENTA_EX}to{Colors.LIGHTYELLOW_EX} [new_field] {Colors.LIGHTMAGENTA_EX}in{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在指定的文件中将一个字段的值更新为另一个值。{Colors.RESET}")
        print(
            f"\n{Colors.LIGHTMAGENTA_EX}{Command.UPDATE}{Colors.LIGHTYELLOW_EX} [old_field] {Colors.LIGHTMAGENTA_EX}to{Colors.LIGHTYELLOW_EX} [new_field] {Colors.LIGHTMAGENTA_EX}except{Colors.LIGHTYELLOW_EX} [filename1] [filename2] ...")
        print(f"{Colors.RESET}  - 在除指定文件外的所有文件中将一个字段的值更新为另一个值。{Colors.RESET}")
        print(f"\n{Colors.LIGHTMAGENTA_EX}{Command.EXIT}{Colors.RESET}")
        print(f"{Colors.RESET}  - 退出程序。{Colors.RESET}")
        print(f"\n{Colors.LIGHTMAGENTA_EX}{Command.CLEAN}{Colors.RESET}")
        print(f"{Colors.RESET}  - 清理控制台的当前显示内容。{Colors.RESET}")

    @staticmethod
    def clean_console():
        os.system('cls' if os.name == 'nt' else 'clear')


class ExpressionProcessor:
    EXP_FIELD = '-'
    EXP_CHUNK = ['{', '}']

    # -
    def parse_field_exp(self, field: str) -> list:
        if '-' in field:
            try:
                start, end = map(int, field.split(self.EXP_FIELD))
                return list(range(start, end + 1))
            except ValueError:
                print(Colors.RED + f"输入'{field}'无法解析为数字区间，将作为字符串处理" + Colors.RESET)
                return [field]
        else:
            return [field]

    # { }
    def parse_chunk_exp(self, expression: str) -> tuple:
        if expression.startswith(self.EXP_CHUNK[0]) and expression.endswith(self.EXP_CHUNK[1]):
            key_column, val = expression[1:-1].split("=")
            return key_column, self.parse_field_exp(val)
        else:
            return None, self.parse_field_exp(expression)


# debug
if __name__ == "__main__":
    exp_proc = ExpressionProcessor()
    while True:
        exp = input("\ntesting expression: ")
        rst = exp_proc.parse_chunk_exp(exp)
        print(rst)
