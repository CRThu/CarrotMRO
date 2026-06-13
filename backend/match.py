from rapidfuzz import process, fuzz

TOP_N = 3

def search(data: dict, query: str, limit: int = TOP_N) -> list[str]:
    """
    输入json对象和搜索的字符串，可选数量，返回匹配的列表
    """
    items = data.get("items", [])
    matches = process.extract(query, items, scorer=fuzz.token_set_ratio, limit=limit)
    return [m[0] for m in matches]
