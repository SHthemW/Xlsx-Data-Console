from abc import abstractmethod


class Command:
    @abstractmethod
    def cmd_keyword(self):
        pass
