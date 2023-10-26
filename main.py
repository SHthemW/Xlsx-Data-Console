# main.py
from entities.color import Colors, init_color
from services.excel_processor import ExcelProcessor
from services.command_processor import InstructProcessor
from services.command_processor import ExpressionProcessor
from utilities.local import get_directory, config_is_exist, create_config_file, CONFIG_PATH

if __name__ == "__main__":
    try:
        init_color()

        if not config_is_exist():
            print(Colors.CYAN + "\n初次使用, 请初始化配置信息: \n" + Colors.RESET)
            create_config_file()
            print(Colors.GREEN + f"\n数据配置已完成. 若需更改, 请查看目录下的{CONFIG_PATH}.\n" + Colors.RESET)

        exp_processor = ExpressionProcessor()
        excel_processor = ExcelProcessor(get_directory())
        command_processor = InstructProcessor(get_directory(), excel_processor, exp_processor)
        command_processor.clean_console()

        while True:
            command = input(Colors.CYAN + "\n请输入命令（输入'help'查看可用指令）: " + Colors.RESET).split()
            if not command_processor.process_command(command):
                break  # 如果process_command返回False，那么退出循环

    except Exception as e:
        print(Colors.RED + f"程序执行时发生错误: {str(e)}" + Colors.RESET)
        raise e

