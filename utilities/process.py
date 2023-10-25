import os
import win32gui
import win32con
from utilities.local import get_window_title_suffix


def close_window(title_name: str):
    def window_enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and title_name in win32gui.GetWindowText(hwnd):
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    win32gui.EnumWindows(window_enum_handler, None)


def start_window(file_path: str):
    os.startfile(file_path)


if __name__ == "__main__":
    name = input()
    print(name + get_window_title_suffix())
    close_window(name + get_window_title_suffix())
