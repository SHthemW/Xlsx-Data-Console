from typing import List


def parse_digit_range(field: str) -> List[str]:
    if '-' not in field:
        raise ValueError(f"invalid syntax: {field} is not a range expr")

    start_end = field.split("-")

    if len(start_end) > 2:
        raise ValueError(f"invalid syntax: {field} has too much field.")

    try:
        start, end = map(int, start_end)
    except ValueError:
        raise ValueError(f"invalid value: {field} contains not-number value, cannot parse to range.")

    step = 1 if start < end else -1

    return [str(num) for num in range(start, end + step, step)]


if __name__ == "__main__":
    print(parse_digit_range("114514-114512"))
