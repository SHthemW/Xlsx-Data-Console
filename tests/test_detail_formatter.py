import re
import unittest
from unittest.mock import patch

from wcwidth import wcswidth

from entities.color import Colors
from services.detail_formatter import format_detail_table


ANSI_ESCAPE = re.compile(r"\x1b(?:\[[0-?]*[ -/]*[@-~]|\][^\x1b]*(?:\x1b\\|\x07))")


class DetailFormatterTests(unittest.TestCase):
    def assert_lines_fit_width(self, output, terminal_width):
        for line in output.splitlines():
            visible_line = ANSI_ESCAPE.sub("", line)
            self.assertLessEqual(wcswidth(visible_line), terminal_width)

    def test_aligns_chinese_values_and_long_field_names(self):
        fields = [
            ("##var", None),
            ("Id", 5060020),
            ("Category", "连体服"),
            ("ItemIdArray", None),
            ("Desc", "梦幻流萤"),
            ("Display#default=TRUE", None),
            ("OverrideTeleportEffectPath", None),
            ("OverrideTeleportSoundPath", None),
        ]

        output = ANSI_ESCAPE.sub("", format_detail_table(fields, ["流萤"]))
        lines = output.splitlines()
        separator_columns = []

        for line in lines:
            separators = [match.start() for match in re.finditer(r"(?<!\S):(?=\s)", line)]
            self.assertEqual(2, len(separators))
            separator_columns.append(tuple(wcswidth(line[:position]) for position in separators))

        self.assertTrue(all(columns == separator_columns[0] for columns in separator_columns))
        self.assertIn("梦幻流萤", output)

    def test_wraps_long_values_without_losing_content(self):
        output = ANSI_ESCAPE.sub(
            "",
            format_detail_table(
                [("Name", "花屿·流萤"), ("Items", "5010038;5010039")],
                ["流萤"],
                terminal_width=20,
            ),
        )

        compact_output = re.sub(r"\s+", "", output)
        self.assertIn("花屿·流萤", compact_output)
        self.assertIn("5010038;5010039", compact_output)

    def test_supports_an_odd_number_of_fields(self):
        output = ANSI_ESCAPE.sub("", format_detail_table([("DefaultObtainTip", "AvatarSrc.")], []))

        self.assertIn("DefaultObtainTip", output)
        self.assertIn("AvatarSrc.", output)

    def test_preserves_field_and_match_colors(self):
        output = format_detail_table([("Desc", "梦幻流萤"), ("Id", 5060020)], ["流萤"])

        self.assertIn(f"{Colors.GREEN}Desc{Colors.RESET}", output)
        self.assertIn(f"{Colors.RED}梦幻流萤{Colors.RESET}", output)
        self.assertIn(f"{Colors.LIGHTYELLOW_EX}5060020{Colors.RESET}", output)

    def test_matches_current_terminal_width(self):
        fields = [
            ("OverrideTeleportEffectPath", "Effect/非常长的特效资源路径"),
            ("OverrideTeleportSoundPath", "Sound/非常长的声音资源路径"),
            ("Desc", "梦幻流萤"),
            ("Display#default=TRUE", None),
        ]

        for terminal_width in (120, 80, 55, 40, 20, 10, 5, 2):
            with self.subTest(terminal_width=terminal_width):
                with patch("services.detail_formatter.get_terminal_width", return_value=terminal_width):
                    output = format_detail_table(fields, ["流萤"])

                self.assert_lines_fit_width(output, terminal_width)

    def test_switches_to_one_detail_per_row_on_narrow_terminals(self):
        fields = [("Category", "连体服"), ("Desc", "梦幻流萤")]

        wide_output = ANSI_ESCAPE.sub("", format_detail_table(fields, ["流萤"], terminal_width=80))
        narrow_output = ANSI_ESCAPE.sub("", format_detail_table(fields, ["流萤"], terminal_width=40))

        self.assertEqual(1, len(wide_output.splitlines()))
        self.assertEqual(2, len(narrow_output.splitlines()))

    def test_uses_wide_terminal_space_for_long_values(self):
        first_path = "Assets/Res/Entity/Item/Avatar/Suit/60031/TeleportEffect"
        second_path = "Assets/Res/Entity/Item/Avatar/Suit/60031/HarvestEffect"
        fields = [
            ("OverrideTeleportEffectPath", first_path),
            ("OverridePickEffectPath", second_path),
        ]

        output = ANSI_ESCAPE.sub(
            "",
            format_detail_table(fields, ["Assets"], terminal_width=160),
        )

        self.assertIn("Assets/Res/Entity/Item/Avatar", output)
        self.assert_lines_fit_width(output, 160)


if __name__ == "__main__":
    unittest.main()
