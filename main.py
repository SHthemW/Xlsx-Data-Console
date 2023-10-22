# main.py
from entities.color import Colors
from services.excel_processor import ExcelProcessor
from services.command_processor import InstructProcessor
from services.command_processor import ExpressionProcessor
from utilities.local import get_directory

if __name__ == "__main__":
    exp_processor = ExpressionProcessor()
    excel_processor = ExcelProcessor(get_directory())
    command_processor = InstructProcessor(excel_processor, exp_processor)
    command_processor.clean_console()

    while True:
        try:
            command = input(Colors.CYAN + "\n请输入命令（输入'help'查看可用指令）: " + Colors.RESET).split()
            if not command_processor.process_command(command):
                break  # 如果process_command返回False，那么退出循环
        except Exception as e:
            print(Colors.RED + f"程序执行时发生错误: {str(e)}" + Colors.RESET)
