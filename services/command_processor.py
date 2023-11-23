import os
from entities.color import Colors
from entities.command import Command, Keyword, Expression
from services.command_helper import print_help
from utilities import process
from utilities.local import get_window_title_suffix, enable_auto_close, enable_auto_restart
from utilities.process import close_window


class InstructProcessor:
    __latest_file_name: str = None

    @property
    def latest_file_name(self) -> str:
        return f"{self.__directory}\\{self.__latest_file_name}"

    @latest_file_name.setter
    def latest_file_name(self, val):
        if not val: return
        self.__latest_file_name = self.__excel_proc.get_pure_filenames(val)[0]

    def __init__(self, directory, excel_processor, expression_processor):
        self.__directory = directory
        self.__excel_proc = excel_processor
        self.__exp_proc = expression_processor

    def process_command(self, command: list) -> bool:
        # basic commands
        if Command.EXIT in command:
            return False
        elif Command.CLEAN in command:
            self.clean_console()
            return True
        elif Command.HELP in command:
            print_help()
            return True

        # find command used to query row in files.
        if command[0].lower() == Command.FIND:
            simple_only: bool = Keyword.SIMPLE in command
            has_scope: bool = Keyword.IN in command or Keyword.EXCEPT in command

            fields = command[1:-1] if not has_scope and simple_only \
                else command[1:-2] if has_scope and not simple_only \
                else command[1:-3] if has_scope and simple_only \
                else command[1:] if not has_scope and not simple_only \
                else "invalid command"
            filenames_str = ' '.join(command[-2:]) if has_scope else Keyword.ALL
            for field in fields:
                col, val = self.__exp_proc.parse_chunk_exp(field)
                print(f"\n正在查找字段'{col}-{val}'：")
                result = self.__excel_proc.search_files((col, val), self.__excel_proc.parse_filenames(filenames_str),
                                                        not simple_only)
                self.latest_file_name = result

        # open command used to open the latest file or specified.
        elif command[0].lower() == Command.OPEN:
            if len(command) < 2:
                if not self.__latest_file_name:
                    print(Colors.RED + "未检测到最近活跃的文件, 请尝试手动指定文件名." + Colors.RESET)
                    return True
                process.start_window(self.latest_file_name)
                print(Colors.GREEN + f"最近活跃的文件为{self.latest_file_name}, 已将其打开." + Colors.RESET)

            elif len(command) == 2:
                self.latest_file_name = command[1]
                process.start_window(self.latest_file_name)
                print(Colors.GREEN + f"打开了文件{self.latest_file_name}." + Colors.RESET)

            else:
                print(Colors.RED + f"格式错误: 输入 {Command.HELP} 以查看帮助." + Colors.RESET)
            return True

        # update command used to update rows in files
        # (must assign the action scope)
        elif command[0].lower() == Command.UPDATE:
            if len(command) < 5 or (
                    command[2].lower() != Keyword.TO and (
                    command[4].lower() != Keyword.IN and command[4].lower() != Keyword.EXCEPT)):
                print(Colors.RED + "更新操作必须指定作用域" + Colors.RESET)
                return True

            old_field, new_field, filenames_str = command[1], command[3], ' '.join(command[4:])
            old_col, old_val = self.__exp_proc.parse_chunk_exp(old_field)
            new_col, new_val = self.__exp_proc.parse_chunk_exp(new_field)

            print(f"\n正在更新字段'{old_col}-{old_val}'至'{new_val}'：")
            # close
            if enable_auto_close():
                for filename in self.__excel_proc.get_pure_filenames(filenames_str):
                    window_title = filename + get_window_title_suffix()
                    close_window(window_title)
            # update
            result = self.__excel_proc.update_files((old_col, old_val),
                                                    (new_col, new_val),
                                                    self.__excel_proc.parse_filenames(filenames_str),
                                                    enable_auto_restart())
            self.latest_file_name = result

        # create command used to create new rows in files.
        # (must assign the action scope)
        elif command[0].lower() == Command.CREATE:
            return True

        else:
            print(Colors.RED + f"无法识别的命令: {command[0]}" + Colors.RESET)
        return True

    @staticmethod
    def clean_console():
        os.system('cls' if os.name == 'nt' else 'clear')


class ExpressionProcessor:

    # -
    @staticmethod
    def parse_field_exp(field: str) -> list:
        if '-' in field:
            try:
                start, end = map(int, field.split(Expression.FIELD))
                return list(range(start, end + 1))
            except ValueError:
                print(Colors.RED + f"输入'{field}'无法解析为数字区间，将作为字符串处理" + Colors.RESET)
                return [field]
        else:
            return [field]

    # { }
    def parse_chunk_exp(self, expression: str) -> tuple:
        if expression.startswith(Expression.CHUNK[0]) and expression.endswith(Expression.CHUNK[1]):
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
