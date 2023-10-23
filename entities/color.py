# 定义颜色
from colorama import Fore, init


class Colors:
    RED = Fore.RED
    GREEN = Fore.GREEN
    LIGHTYELLOW_EX = Fore.LIGHTYELLOW_EX
    LIGHTMAGENTA_EX = Fore.LIGHTMAGENTA_EX
    YELLOW = Fore.YELLOW
    CYAN = Fore.CYAN
    RESET = Fore.RESET


def init_color():
    init()
