import ctypes
import os
import unittest
from unittest.mock import patch

from utilities.terminal import (
    _ConsoleScreenBufferInfo,
    _SmallRect,
    _get_windows_console_width,
    get_terminal_width,
)


class TerminalWidthTests(unittest.TestCase):
    def test_prefers_windows_visible_console_width(self):
        with patch("utilities.terminal._get_windows_console_width", return_value=83):
            self.assertEqual(83, get_terminal_width())

    def test_falls_back_when_console_handle_is_unavailable(self):
        terminal_size = os.terminal_size((71, 24))

        with (
            patch("utilities.terminal._get_windows_console_width", return_value=None),
            patch("utilities.terminal.get_terminal_size", return_value=terminal_size),
        ):
            self.assertEqual(71, get_terminal_width())

    def test_windows_window_rectangle_uses_inclusive_coordinates(self):
        window = _SmallRect(10, 0, 89, 29)

        self.assertEqual(80, window.Right - window.Left + 1)

    def test_reads_visible_width_from_windows_console_info(self):
        def set_console_info(_, console_info_pointer):
            console_info = ctypes.cast(
                console_info_pointer,
                ctypes.POINTER(_ConsoleScreenBufferInfo),
            ).contents
            console_info.srWindow = _SmallRect(10, 0, 89, 29)
            return True

        with (
            patch("utilities.terminal.os.name", "nt"),
            patch("utilities.terminal.ctypes.WinDLL") as win_dll,
        ):
            kernel32 = win_dll.return_value
            kernel32.GetStdHandle.return_value = 123
            kernel32.GetConsoleScreenBufferInfo.side_effect = set_console_info

            self.assertEqual(80, _get_windows_console_width())


if __name__ == "__main__":
    unittest.main()
