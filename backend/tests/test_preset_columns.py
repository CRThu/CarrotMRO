"""preset_columns.py 测试"""

from preset_columns import get_preset_columns, get_preset_labels, PRESET_COLUMNS


def test_get_preset_columns():
    """测试获取预制列"""
    columns = get_preset_columns()
    assert isinstance(columns, list)
    assert len(columns) > 0
    assert all("key" in col for col in columns)
    assert all("label" in col for col in columns)
    assert all("required" in col for col in columns)
    assert all("type" in col for col in columns)


def test_preset_columns_has_required_fields():
    """测试预制列包含必要字段"""
    columns = get_preset_columns()
    labels = [col["label"] for col in columns]
    assert "项目名称" in labels
    assert "数量" in labels
    assert "单位" in labels


def test_get_preset_labels():
    """测试获取预制列标签"""
    labels = get_preset_labels()
    assert isinstance(labels, list)
    assert "项目名称" in labels
    assert "数量" in labels
    assert "单位" in labels
    assert "单价" in labels
    assert "备注" in labels


def test_preset_columns_required_flag():
    """测试必填标记"""
    columns = get_preset_columns()
    required_cols = [col for col in columns if col["required"]]
    required_labels = [col["label"] for col in required_cols]
    assert "项目名称" in required_labels
    assert "数量" in required_labels
    assert "单位" in required_labels


def test_preset_columns_unique_keys():
    """测试 key 唯一性"""
    columns = get_preset_columns()
    keys = [col["key"] for col in columns]
    assert len(keys) == len(set(keys))


def test_preset_columns_unique_labels():
    """测试 label 唯一性"""
    columns = get_preset_columns()
    labels = [col["label"] for col in columns]
    assert len(labels) == len(set(labels))
