import json
import argparse

from rapidfuzz import process, fuzz

TOP_N = 5


def match_names(names: list[str], queries: list[str], limit: int = TOP_N) -> dict[str, list[tuple[str, float]]]:
    """对字符串列表进行模糊搜索，支持多个查询"""
    results = {}
    for q in queries:
        matches = process.extract(q, names, scorer=fuzz.token_set_ratio, limit=limit)
        results[q] = [(m[0], m[1]) for m in matches]
    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="模糊搜索工具")
    parser.add_argument("queries", nargs="+", help="搜索关键词（支持多个）")
    parser.add_argument("data_path", help="数据文件路径（JSON，需包含 items 列表）")
    parser.add_argument("-n", type=int, default=5, help="每个关键词返回结果数量")
    args = parser.parse_args()

    with open(args.data_path, encoding="utf-8") as f:
        data = json.load(f)

    items = data.get("items", [])
    results = match_names(items, args.queries, limit=args.n)
    for q, matches in results.items():
        print(f"\n搜索「{q}」 top {args.n}：\n")
        for name, score in matches:
            print(f"  {score:3.0f}  {name}")
