import json
import sys
from pathlib import Path
from rapidfuzz import process, fuzz

DEFAULT_DATA = str(Path(__file__).parent.parent / "data" / "standard.json")


def search_items(query: str, data_path: str = DEFAULT_DATA, n: int = 5) -> dict:
    """
    搜索项目名称，返回 JSON 格式结果。

    参数:
        query:     搜索关键词
        data_path: JSON 数据文件路径，默认 data/standard.json
        n:         返回前 n 条结果，默认 5

    返回:
        dict: {"query": ..., "total": ..., "source": ..., "results": [{"name": ..., "score": ...}, ...]}
    """
    with open(data_path, encoding="utf-8") as f:
        items = json.load(f)["items"]

    results = process.extract(query, items, scorer=fuzz.token_set_ratio, limit=n)

    return {
        "query": query,
        "total": n,
        "source": data_path,
        "results": [
            {"name": name, "score": score}
            for name, score, _ in results
        ]
    }


def main():
    if len(sys.argv) < 2:
        print("用法: uv run python search.py <关键词> [-n N] [-d <数据文件>]")
        print("示例: uv run python search.py key")
        print("      uv run python search.py key -n 5")
        print("      uv run python search.py key -d data/2026.json")
        sys.exit(1)

    # 解析参数
    query = sys.argv[1]
    n = 5
    data_path = DEFAULT_DATA

    if "-n" in sys.argv:
        idx = sys.argv.index("-n")
        if idx + 1 < len(sys.argv):
            n = int(sys.argv[idx + 1])

    if "-d" in sys.argv:
        idx = sys.argv.index("-d")
        if idx + 1 < len(sys.argv):
            data_path = sys.argv[idx + 1]

    # 调用搜索方法
    result = search_items(query, data_path, n)

    # 友好显示
    print(f"\n搜索「{result['query']}」 top {result['total']}（数据来源：{result['source']}）：\n")
    for item in result["results"]:
        print(f"  {item['score']:3.0f}  {item['name']}")


if __name__ == "__main__":
    main()
