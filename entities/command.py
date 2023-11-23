class Command:
    FIND: str = "find"
    EXIT: str = "exit"
    HELP: str = "help"
    OPEN: str = "open"
    CLEAN: str = "clean"
    UPDATE: str = "update"
    CREATE: str = "create"


class Keyword:
    TO: str = "to"
    IN: str = "in"
    ALL: str = "ALL"
    SIMPLE: str = "simple"
    EXCEPT: str = "except"

    @staticmethod
    def is_keyword(s: str):
        return s in [Keyword.IN, Keyword.EXCEPT, Keyword.ALL, Keyword.TO, Keyword.SIMPLE]


class Expression:
    FIELD = '-'
    CHUNK = ['{', '}']
