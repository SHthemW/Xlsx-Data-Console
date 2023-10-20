from colorama import Fore
import xml.etree.ElementTree as Et

# 定义颜色
class Colors:
    RED = Fore.RED
    GREEN = Fore.GREEN
    LIGHTYELLOW_EX = Fore.LIGHTYELLOW_EX
    LIGHTMAGENTA_EX = Fore.LIGHTMAGENTA_EX
    YELLOW = Fore.YELLOW
    CYAN = Fore.CYAN
    RESET = Fore.RESET


# 定义要搜索的目录
def get_directory() -> str:
    # 使用open()函数打开文件
    with open('_config.txt', 'r') as file:
        data = file.read()

    # 使用ElementTree解析数据
    root = Et.fromstring(data)

    # 获取<datasource>标签中的文本，并赋值给directory
    return root.text.strip()





