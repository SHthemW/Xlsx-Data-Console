import main
from entities.command.interface import Command
from typing import List, final
from entities.expression import expr_range
from entities.expression import expr_scope

@final
class FindCommand(Command):

    @property
    def cmd_keywords(self) -> List[str]:
        return ["find"]

    @property
    def help_message(self) -> str:
        return ""

    def __init__(self, args: List[str]):
        self.__file_scope: tuple = expr_scope.try_parse_scope(args)
        super().__init__(args)
        self.__find_target: list = self.parse_filenames()
        self.__simple_find: bool = any(arg in self._args for arg in ["simple", "s"])

    def execute(self) -> bool:
        for to_find in self.__find_target:
            col, val = None, to_find
            print(f"\n正在查找字段'{col}-{val}'：")
            result = main.excel_processor.search_files(
                field=(col, val),
                filenames=main.excel_processor.parse_file_scope(self.__file_scope),
                show_detail=not self.__simple_find
            )
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
        super().debug()
        print("simple:", self.__simple_find)
        print("target:", self.__find_target)
        print("scopes:", self.__file_scope)


if __name__ == "__main__":
    cmd = FindCommand(["114514-114516", "191981", "--s", "in", "abc.xlsx"])
    cmd.execute()
    cmd.debug()
