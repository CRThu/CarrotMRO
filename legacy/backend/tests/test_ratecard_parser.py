"""ratecard_parser.py 测试"""

import csv
from io import BytesIO

import openpyxl

from ratecard_parser import parse_ratecard_file, extract_names


def _make_excel(header: list[str], rows: list[list[str]]) -> bytes:
    """创建测试用 Excel 文件"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for row in rows:
        ws.append(row)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _make_csv(header: list[str], rows: list[list[str]]) -> bytes:
    """创建测试用 CSV 文件"""
    import io
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(header)
    for row in rows:
        writer.writerow(row)
    return buf.getvalue().encode("utf-8-sig")
    return buf.getvalue()


class TestParseRatecardFile:
    def test_parse_excel(self):
        """测试 Excel 解析"""
        content = _make_excel(
            ["!项目名称", "!单位", "?综合单价"],
            [
                ["螺栓M10", "个", "2.50"],
                ["垫片M10", "个", "0.80"],
            ],
        )
        result = parse_ratecard_file(content, "test.xlsx")

        assert len(result["columns"]) == 3
        assert result["columns"][0]["name"] == "项目名称"
        assert result["columns"][0]["strict"] is True
        assert result["columns"][2]["name"] == "综合单价"
        assert result["columns"][2]["strict"] is False
        assert len(result["items"]) == 2

    def test_parse_csv(self):
        """测试 CSV 解析"""
        content = _make_csv(
            ["!项目名称", "!单位", "?综合单价"],
            [
                ["螺栓M10", "个", "2.50"],
                ["垫片M10", "个", "0.80"],
            ],
        )
        result = parse_ratecard_file(content, "test.csv")

        assert len(result["columns"]) == 3
        assert len(result["items"]) == 2

    def test_skip_rows_with_empty_strict_column(self):
        """测试跳过 strict 列为空的行"""
        content = _make_excel(
            ["!项目名称", "!单位"],
            [
                ["螺栓M10", "个"],
                ["", "个"],  # strict 列为空，应跳过
                ["垫片M10", ""],  # strict 列为空，应跳过
                ["螺丝M8", "个"],
            ],
        )
        result = parse_ratecard_file(content, "test.xlsx")

        assert len(result["items"]) == 2
        assert result["items"][0]["项目名称"] == "螺栓M10"
        assert result["items"][1]["项目名称"] == "螺丝M8"

    def test_no_header_raises_error(self):
        """测试无标记行时抛出异常"""
        content = _make_excel(
            ["项目名称", "单位"],
            [["螺栓M10", "个"]],
        )
        try:
            parse_ratecard_file(content, "test.xlsx")
            assert False, "应该抛出 ValueError"
        except ValueError as e:
            assert "未找到标记行" in str(e)

    def test_unsupported_format_raises_error(self):
        """测试不支持的文件格式"""
        try:
            parse_ratecard_file(b"test", "test.txt")
            assert False, "应该抛出 ValueError"
        except ValueError as e:
            assert "不支持的文件格式" in str(e)


class TestExtractNames:
    def test_extract_with_explicit_column(self):
        """测试使用显式指定的列名"""
        columns = [
            {"name": "项目名称", "strict": True, "alias": None},
            {"name": "单位", "strict": True, "alias": None},
        ]
        items = [
            {"项目名称": "螺栓M10", "单位": "个"},
            {"项目名称": "垫片M10", "单位": "个"},
        ]
        names = extract_names(columns, items, name_column="项目名称")
        assert names == ["螺栓M10", "垫片M10"]

    def test_extract_with_alias(self):
        """测试使用 alias 查找 name 列"""
        columns = [
            {"name": "物料名称", "strict": True, "alias": "name"},
            {"name": "单位", "strict": True, "alias": None},
        ]
        items = [
            {"物料名称": "螺栓M10", "单位": "个"},
            {"物料名称": "垫片M10", "单位": "个"},
        ]
        names = extract_names(columns, items)
        assert names == ["螺栓M10", "垫片M10"]

    def test_extract_empty_when_no_name_column(self):
        """测试无 name 列时返回空列表"""
        columns = [
            {"name": "单位", "strict": True, "alias": None},
        ]
        items = [{"单位": "个"}]
        names = extract_names(columns, items)
        assert names == []

    def test_extract_empty_items(self):
        """测试空数据时返回空列表"""
        columns = [{"name": "项目名称", "strict": True, "alias": "name"}]
        names = extract_names(columns, [])
        assert names == []
