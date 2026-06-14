import openpyxl
import csv
import json
import argparse
from io import BytesIO
from pathlib import Path


def parse_ratecard_file(content: bytes, filename: str) -> dict:
    """解析 Excel/CSV 文件，返回定价表数据结构。

    表头行使用前缀标记列的匹配属性：
      ! = 必须匹配项（报价单比对时必须匹配）
      ? = 可选匹配项（报价单比对时可选匹配，允许差异）
      无前缀 = 忽略列，不导入

    返回格式:
    {
      "columns": [{"name": str, "strict": bool, "alias": str|None}, ...],
      "items": [{col_name: value, ...}, ...]
    }
    """
    filename_lower = filename.lower()
    if filename_lower.endswith(".xlsx") or filename_lower.endswith(".xls"):
        return _parse_excel(content)
    elif filename_lower.endswith(".csv"):
        return _parse_csv(content)
    else:
        raise ValueError("不支持的文件格式，请上传 .xlsx / .xls / .csv 文件")


def _parse_excel(content: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
    ws = wb.active
    rows = [tuple(cell for cell in row) for row in ws.iter_rows(values_only=True)]
    return _parse_rows(rows)


def _parse_csv(content: bytes) -> dict:
    text = content.decode("utf-8-sig")
    reader = csv.reader(text.splitlines())
    rows = [tuple(cell for cell in row) for row in reader]
    return _parse_rows(rows)


def _parse_rows(rows: list[tuple]) -> dict:
    if not rows:
        return {"columns": [], "items": []}

    header_idx = _find_header_row(rows)
    if header_idx is None:
        raise ValueError("未找到标记行（表头需包含 ! 或 ? 前缀）")

    columns = _parse_columns(rows[header_idx])
    _assign_name_alias(columns)

    items = _read_data_rows(rows, header_idx, columns)

    return {
        "columns": [
            {"name": c["name"], "strict": c["strict"], "alias": c["alias"]}
            for c in columns
        ],
        "items": items,
    }


def _find_header_row(rows: list[tuple]) -> int | None:
    for i, row in enumerate(rows):
        for cell in row:
            if cell is not None:
                s = str(cell).strip()
                if s.startswith("!") or s.startswith("?"):
                    return i
    return None


def _parse_columns(header_row: tuple) -> list[dict]:
    columns = []
    for col_idx, cell in enumerate(header_row):
        if cell is None:
            continue
        s = str(cell).strip()
        if not s.startswith("!") and not s.startswith("?"):
            continue
        strict = s.startswith("!")
        col_name = s.lstrip("!?").strip()
        if not col_name:
            continue
        alias = "name" if "名称" in col_name else None
        columns.append({
            "col_idx": col_idx,
            "name": col_name,
            "strict": strict,
            "alias": alias,
        })
    return columns


def _assign_name_alias(columns: list[dict]):
    if any(c["alias"] == "name" for c in columns):
        return
    for col in columns:
        if col["strict"]:
            col["alias"] = "name"
            return


def _read_data_rows(rows: list[tuple], header_idx: int, columns: list[dict]) -> list[dict]:
    items = []
    for row in rows[header_idx + 1:]:
        if not _is_valid_row(row, columns):
            continue
        item = {}
        for col in columns:
            val = row[col["col_idx"]] if col["col_idx"] < len(row) else None
            item[col["name"]] = str(val).strip() if val is not None else ""
        items.append(item)
    return items


def _is_valid_row(row: tuple, columns: list[dict]) -> bool:
    for col in columns:
        if not col["strict"]:
            continue
        val = row[col["col_idx"]] if col["col_idx"] < len(row) else None
        if val is None or str(val).strip() == "":
            return False
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="解析带 !/? 标记的定价表 Excel/CSV")
    parser.add_argument("-i", "--input", required=True, help="输入文件路径")
    parser.add_argument("-o", "--output", default=None, help="输出 JSON 路径（默认: 同目录同名 .json）")
    args = parser.parse_args()

    input_path = Path(args.input)
    content = input_path.read_bytes()
    result = parse_ratecard_file(content, input_path.name)

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix(".json")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"输入: {input_path}")
    print(f"输出: {output_path}")
    print(f"列数: {len(result['columns'])}，有效行: {len(result['items'])}")
    for col in result["columns"]:
        tag = "!" if col["strict"] else "?"
        alias_tag = f" (alias={col['alias']})" if col["alias"] else ""
        print(f"  {tag}{col['name']}{alias_tag}")


def extract_names(columns: list[dict], items: list[dict]) -> list[str]:
    """从定价表数据中提取 name 列的值"""
    name_col = None
    for col in columns:
        if col.get("alias") == "name":
            name_col = col["name"]
            break
    if name_col is None:
        for col in columns:
            if "名称" in col["name"]:
                name_col = col["name"]
                break
    if name_col is None and columns:
        for col in columns:
            if col.get("strict"):
                name_col = col["name"]
                break
    if name_col is None:
        return []
    return [item[name_col] for item in items if name_col in item]
