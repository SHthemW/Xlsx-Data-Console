class CommandName:
    FIND: str = "find"
    EXIT: str = "exit"
    HELP: str = "help"
    OPEN: str = "open"
    CLEAN: str = "clean"
    UPDATE: str = "update"
    CREATE: str = "create"


class KeywordName:
    TO: str = "to"
    IN: str = "in"
    ALL: str = "ALL"
    SIMPLE: str = "simple"
    EXCEPT: str = "except"

    @staticmethod
    def is_keyword(s: str):
        return s in [KeywordName.IN, KeywordName.EXCEPT, KeywordName.ALL, KeywordName.TO, KeywordName.SIMPLE]


class Expression:
    FIELD = '-'
    CHUNK = ['{', '}']
