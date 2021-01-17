import re


def rexex(pattern: str, query: str, allow_empty: bool = True):
    r = re.compile(pattern)
    s = [
        m.groupdict()
        for m in r.finditer(query)
    ]
    d = {}
    for match in s:
        for key, value in match.items():
            if not value and not allow_empty:
                continue
            d.setdefault(key, []).append(value)
    return d
