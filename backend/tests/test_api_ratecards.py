"""定价表 API 测试"""

import json


def test_get_ratecards_empty(client):
    """测试获取空定价表列表"""
    response = client.get("/api/ratecards")
    assert response.status_code == 200
    assert response.json()["ratecards"] == []


def test_create_ratecard(client):
    """测试创建定价表"""
    response = client.post("/api/ratecards/test-ratecard")
    assert response.status_code == 200
    assert "创建成功" in response.json()["message"]


def test_create_ratecard_duplicate(client):
    """测试创建重复定价表"""
    client.post("/api/ratecards/test-ratecard")
    response = client.post("/api/ratecards/test-ratecard")
    assert response.status_code == 400


def test_get_ratecard_data(client):
    """测试获取定价表数据"""
    client.post("/api/ratecards/test-ratecard")
    response = client.get("/api/ratecards/test-ratecard")
    assert response.status_code == 200
    data = response.json()
    assert "columns" in data
    assert "items" in data


def test_get_ratecard_not_found(client):
    """测试获取不存在的定价表"""
    response = client.get("/api/ratecards/nonexistent")
    assert response.status_code == 404


def test_match_ratecard_empty_queries(client):
    """测试空查询"""
    response = client.post("/api/match", json={
        "ratecard_name": "test",
        "queries": [],
    })
    assert response.status_code == 400
