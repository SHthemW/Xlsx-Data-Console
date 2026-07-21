import ctypes
import os
from ctypes import wintypes
from shutil import get_terminal_size


STD_OUTPUT_HANDLE = -11
FALLBACK_TERMINAL_WIDTH = 120
FALLBACK_TERMINAL_HEIGHT = 24


class _Coord(ctypes.Structure):
    _fields_ = [
        ("X", wintypes.SHORT),
        ("Y", wintypes.SHORT),
    ]


class _SmallRect(ctypes.Structure):
    _fields_ = [
        ("Left", wintypes.SHORT),
        ("Top", wintypes.SHORT),
        ("Right", wintypes.SHORT),
        ("Bottom", wintypes.SHORT),
    ]


class _ConsoleScreenBufferInfo(ctypes.Structure):
    _fields_ = [
        ("dwSize", _Coord),
        ("dwCursorPosition", _Coord),
        ("wAttributes", wintypes.WORD),
        ("srWindow", _SmallRect),
        ("dwMaximumWindowSize", _Coord),
    ]


def _get_windows_console_width() -> int | None:
    if os.name != "nt":
        return None

    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        get_std_handle = kernel32.GetStdHandle
        get_std_handle.argtypes = [wintypes.DWORD]
        get_std_handle.restype = wintypes.HANDLE

        get_console_info = kernel32.GetConsoleScreenBufferInfo
        get_console_info.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(_ConsoleScreenBufferInfo),
        ]
        get_console_info.restype = wintypes.BOOL

        output_handle = get_std_handle(STD_OUTPUT_HANDLE)
        invalid_handle = ctypes.c_void_p(-1).value
        if output_handle in (None, 0, invalid_handle):
            return None

        console_info = _ConsoleScreenBufferInfo()
        if not get_console_info(output_handle, ctypes.byref(console_info)):
            return None

        window = console_info.srWindow
        width = window.Right - window.Left + 1
        return width if width > 0 else None
    except (AttributeError, OSError, ValueError):
        return None


def get_terminal_width(fallback: int = FALLBACK_TERMINAL_WIDTH) -> int:
    windows_console_width = _get_windows_console_width()
    if windows_console_width is not None:
        return windows_console_width

    return get_terminal_size(
        fallback=(fallback, FALLBACK_TERMINAL_HEIGHT)
    ).columns
