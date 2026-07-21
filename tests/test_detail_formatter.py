import re
import unittest

from wcwidth import wcswidth

from entities.color import Colors
from services.detail_formatter import format_detail_table


ANSI_ESCAPE = re.compile(r"\x1b(?:\[[0-?]*[ -/]*[@-~]|\][^\x1b]*(?:\x1b\\|\x07))")


class DetailFormatterTests(unittest.TestCase):
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
                [("Name", "花屿·流萤"), ("ItemIdArray", "5010038;5010039")],
                ["流萤"],
            ),
        )

        self.assertIn("花屿·流萤", output)
        self.assertIn("5010038;50", output)
        self.assertIn("10039", output)

    def test_supports_an_odd_number_of_fields(self):
        output = ANSI_ESCAPE.sub("", format_detail_table([("DefaultObtainTip", "AvatarSrc.")], []))

        self.assertIn("DefaultObtainTip", output)
        self.assertIn("AvatarSrc.", output)

    def test_preserves_field_and_match_colors(self):
        output = format_detail_table([("Desc", "梦幻流萤"), ("Id", 5060020)], ["流萤"])

        self.assertIn(f"{Colors.GREEN}Desc{Colors.RESET}", output)
        self.assertIn(f"{Colors.RED}梦幻流萤{Colors.RESET}", output)
        self.assertIn(f"{Colors.LIGHTYELLOW_EX}5060020{Colors.RESET}", output)


if __name__ == "__main__":
    unittest.main()
