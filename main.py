# main.py
import _config
from excel_processor import ExcelProcessor
from command_processor import CommandProcessor



if __name__ == "__main__":
    excel_processor = ExcelProcessor(_config.directory)
    command_processor = CommandProcessor(excel_processor)
    command_processor.clean_console()

    while True:
        command = input(_config.Colors.CYAN + "\n请输入命令（输入'help'查看可用指令）: " + _config.Colors.RESET).split()
        if not command_processor.process_command(command):
            break  # 如果process_command返回False，那么退出循环
