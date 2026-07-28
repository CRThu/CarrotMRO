"""报价单 API 测试"""


def test_get_quotations_not_found(client):
    """测试获取不存在项目的报价单"""
    response = client.get("/api/projects/nonexistent/quotations")
    assert response.status_code == 404


def test_create_quotation_not_found(client):
    """测试在不存在的项目中创建报价单"""
    response = client.post("/api/projects/nonexistent/quotations")
    assert response.status_code == 404


def test_quotation_crud(client):
    """测试报价单完整 CRUD"""
    # 创建项目
    client.post("/api/projects/test-project")

    # 创建报价单
    response = client.post("/api/projects/test-project/quotations")
    assert response.status_code == 200
    filename = response.json()["file"]

    # 获取报价单列表
    response = client.get("/api/projects/test-project/quotations")
    assert response.status_code == 200
    assert filename in response.json()["files"]

    # 获取报价单数据
    response = client.get(f"/api/projects/test-project/quotations/{filename}")
    assert response.status_code == 200
    data = response.json()
    assert "items" in data
    assert "created_at" in data
    assert "last_edit_time" in data

    # 保存报价单
    data["items"] = [{"项目名称": "测试", "数量": "10", "单位": "个", "单价": "5.00"}]
    response = client.put(f"/api/projects/test-project/quotations/{filename}", json=data)
    assert response.status_code == 200

    # 删除报价单
    response = client.delete(f"/api/projects/test-project/quotations/{filename}")
    assert response.status_code == 200

    # 确认已删除
    response = client.get("/api/projects/test-project/quotations")
    assert filename not in response.json()["files"]
