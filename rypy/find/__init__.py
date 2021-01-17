import re


def rexex(pattern: str, query: str):
    r = re.compile(pattern)
    s = [
        m.groupdict()
        for m in r.finditer(query)
    ]
    d = {}
    for match in s:
        for key, value in match.items():
            if value:
                d.setdefault(key, []).append(value)
    return d
