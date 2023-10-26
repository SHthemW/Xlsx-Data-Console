from typing import List

from entities.color import Colors
from entities.command import Command, Expression, Keyword


def print_help():
    print(f"\n{Colors.GREEN}目前可用的命令：{Colors.RESET}")
    print_help_message(
        "Command", Command.FIND,
        ['find command is used for finding value.',
         'you can also write <scope statement> as an option.'],
        ['find [value]', 'find [value] in [filename]']
    )
    print_help_message(
        "Command", Command.OPEN,
        ['open command is used for open a file.',
         'the argument is a option to appoint a file name, no-args will automatic choose the latest file.'],
        ['open', 'open [filename]']
    )
    print_help_message(
        "Command", Command.UPDATE,
        ['update command is used for update some values or rows.',
         'use normal value to match values, and {} expr to match rows.',
         'Note: the <scope statement> expr is REQUIREMENT in update command.'],
        ['update [origValue] to [targetValue] in [filename]']
    )
    print_help_message(
        "Command", Command.EXIT,
        ['exit the program immediately.'],
    )
    print_help_message(
        "Command", Command.CLEAN,
        ['clean the console texts that already showed.'],
    )
    print(f"\n{Colors.GREEN}表达式：{Colors.RESET}")
    print_help_message(
        "Expression", Expression.FIELD,
        ['field expr used for match a range of number.'],
        ['1001-1003 => [1001, 1002, 1003]']
    )
    print_help_message(
        "Expression", " ".join(Expression.CHUNK),
        ['field expr used for match a row with given condition.'],
        ['{Id=1001} => a row with a Id-1001 cell']
    )


def print_help_message(cmd_type: str, cmd_name: str, message: List[str], example: List[str] = None):
    print(f"\n{cmd_type} " + Colors.LIGHTMAGENTA_EX + f'{cmd_name}' + Colors.RESET + ":")
    print("Describe:")
    for line in message:
        print(Colors.RESET + f'    {line}' + Colors.RESET)
    if not example:
        return
    print("Examples:")
    for line in example:
        print("    " + " ".join(
            [f'{Colors.LIGHTYELLOW_EX if (word == cmd_name or Keyword.is_keyword(word)) else Colors.RESET}'
             f'{word}{Colors.RESET}'
             for word in line.split()]))


if __name__ == '__main__':
    print_help_message(
        "Command", Command.FIND,
        ['find command is used for finding value.', 'you can also write <scope statement> as an option.'],
        ['find [value]', 'find [value] in [filename]']
    )
