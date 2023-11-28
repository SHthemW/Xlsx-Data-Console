from itertools import dropwhile
from typing import List, Tuple

IN_SCOPE_CHAR: str = "in"
EX_SCOPE_CHAR: str = "except"


def try_parse_scope(fields: List[str]) -> Tuple[bool, List[str]] | None:
    """
    parse and REMOVE scope stmt.
    :param fields: a stmt may contains scope char
    :return: if contains scope stmt, return scope info : in(true) or except(false), files.
             Else return None
    """
    if not any(f for f in fields if f in [IN_SCOPE_CHAR, EX_SCOPE_CHAR]):
        return None

    scope_stmt = list(dropwhile(lambda x: x not in [IN_SCOPE_CHAR, EX_SCOPE_CHAR], fields))

    if len(scope_stmt) <= 1:
        raise ValueError("no scope is pointed.")

    # remove scope_stmt from fields
    for elem in scope_stmt:
        fields.remove(elem)

    return scope_stmt[0] == IN_SCOPE_CHAR, scope_stmt[1::]


if __name__ == "__main__":
    print(try_parse_scope(["123", "in", "abi.xlsx", "cxk.xlsx"]))
    print(try_parse_scope(["123", "except", "cxk.xlsx"]))
    print(try_parse_scope(["123", "cxk.xlsx"]))
