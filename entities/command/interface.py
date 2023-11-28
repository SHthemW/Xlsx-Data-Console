from abc import abstractmethod
from typing import List, final

ARGS_IDENTIFIER: str = "--"


class Command:

    # properties

    @property
    @abstractmethod
    def cmd_keywords(self) -> List[str]:
        """
        :return: all keywords that can call this command
        """
        pass

    @property
    @abstractmethod
    def help_message(self) -> str:
        return "no help message"

    # command core

    def __init__(self, *args: str):
        self._args   = [arg.replace(ARGS_IDENTIFIER, "") for arg in args if arg.startswith(ARGS_IDENTIFIER)]
        self._fields = [arg for arg in args if not arg.startswith(ARGS_IDENTIFIER)]

    @abstractmethod
    def execute(self) -> bool:
        """
        :param args: arguments of this command
        :return: execute is successful
        """
        pass

    def debug(self):
        print(f'args: {self._args}')
        print(f'fields: {self._fields}')
