import re

_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "kh", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "shch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}


def transliterate(text: str) -> str:
    result = []
    for ch in text.lower():
        if ch in _MAP:
            result.append(_MAP[ch])
        elif ch.isascii() and (ch.isalnum() or ch in "_- "):
            result.append(ch)
        else:
            result.append("_")
    code = "".join(result).strip()
    code = re.sub(r"[\s\-]+", "_", code)
    code = re.sub(r"_+", "_", code)
    return code.strip("_")
