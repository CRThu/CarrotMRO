"""match.py 测试"""

from match import match_names


def test_match_single_query():
    """测试单个查询"""
    names = ["螺栓M10", "螺栓M12", "垫片M10", "螺丝M8"]
    results = match_names(names, ["螺栓"], limit=5)
    assert "螺栓" in results
    assert len(results["螺栓"]) > 0
    # 螺栓相关的结果应该排在前面
    assert any("螺栓" in name for name, _ in results["螺栓"])


def test_match_multiple_queries():
    """测试多个查询"""
    names = ["螺栓M10", "垫片M10", "螺丝M8"]
    results = match_names(names, ["螺栓", "垫片"], limit=5)
    assert "螺栓" in results
    assert "垫片" in results


def test_match_limit():
    """测试结果数量限制"""
    names = [f"项目{i}" for i in range(20)]
    results = match_names(names, ["项目"], limit=3)
    assert len(results["项目"]) <= 3


def test_match_empty_query():
    """测试空查询"""
    names = ["螺栓M10", "垫片M10"]
    results = match_names(names, [], limit=5)
    assert results == {}


def test_match_empty_names():
    """测试空名称列表"""
    results = match_names([], ["螺栓"], limit=5)
    assert "螺栓" in results
    assert results["螺栓"] == []


def test_match_no_match():
    """测试无匹配结果"""
    names = ["螺栓M10", "垫片M10"]
    results = match_names(names, ["完全不相关"], limit=5)
    assert "完全不相关" in results
    # 可能有低分结果，但不应该报错
