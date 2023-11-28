from entities.command.interface import Command
from typing import List, final
from entities.expression import expr_range


@final
class FindCommand(Command):

    # find FIELD --simple
    # find FILED --s
    # find FIELD --s

    @property
    def cmd_keywords(self) -> List[str]:
        return ["find"]

    @property
    def help_message(self) -> str:
        return ""

    def __init__(self, *args: str):
        super().__init__(*args)
        # properties
        self.find_target: list = self.parse_filenames()
        self.simple_find: bool = any(arg in self._args for arg in ["simple", "s"])
        self.find_scopes: list = []

    def execute(self) -> bool:
        return True

    def parse_filenames(self) -> List[str]:
        result: List[str] = []
        for field in self._fields:
            if '-' in field:
                result.extend(expr_range.parse_digit_range(field))
            else:
                result.append(field)
        return result

    def debug(self):
        super(FindCommand, self).debug()
        print("target:", self.find_target)


if __name__ == "__main__":
    cmd = FindCommand("114514-114516", "191981", "--s", "--f")
    cmd.execute()
    cmd.debug()
