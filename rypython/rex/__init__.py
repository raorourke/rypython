import re

def stringify(q: str):
    return q if isinstance(q, str) else str(q)

def capture_all(
        pattern: str,
        query: str,
        allow_empty: bool = True,
        flatten: bool = False
):
    query = stringify(query)
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
    if flatten:
        for key, value in d.items():
            if isinstance(value, list):
                if len(value) == 1:
                    d[key] = value[0]
                if len(value) == 0:
                    d[key] = None
    return d

def is_format(pattern: str, query: str):
    query = stringify(query)
    r = re.compile(pattern)
    return bool(re.match(r, query))