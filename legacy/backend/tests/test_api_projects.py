"""项目 API 测试"""

import json


def test_get_projects_empty(client):
    """测试获取空项目列表"""
    response = client.get("/api/projects")
    assert response.status_code == 200
    assert response.json()["projects"] == []


def test_create_project(client):
    """测试创建项目"""
    response = client.post("/api/projects/test-project")
    assert response.status_code == 200
    data = response.json()
    assert data["message"] == "项目创建成功"
    assert data["data"]["name"] == "test-project"


def test_create_project_duplicate(client):
    """测试创建重复项目"""
    client.post("/api/projects/test-project")
    response = client.post("/api/projects/test-project")
    assert response.status_code == 400


def test_get_project_info(client):
    """测试获取项目信息"""
    client.post("/api/projects/test-project")
    response = client.get("/api/projects/test-project")
    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "test-project"
    assert "column_mappings" in data


def test_get_project_not_found(client):
    """测试获取不存在的项目"""
    response = client.get("/api/projects/nonexistent")
    assert response.status_code == 404


def test_get_preset_columns(client):
    """测试获取预制列"""
    response = client.get("/api/preset-columns")
    assert response.status_code == 200
    data = response.json()
    assert "columns" in data
    assert len(data["columns"]) > 0
    labels = [col["label"] for col in data["columns"]]
    assert "项目名称" in labels
