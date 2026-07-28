"""模板 API 测试"""


def test_list_templates_empty(client):
    """测试获取空模板列表"""
    response = client.get("/api/templates")
    assert response.status_code == 200
    assert response.json()["files"] == []


def test_upload_template_invalid_format(client):
    """测试上传无效格式的模板"""
    response = client.post(
        "/api/templates",
        files={"file": ("test.txt", b"not an excel file", "text/plain")},
    )
    assert response.status_code == 400


def test_get_template_not_found(client):
    """测试获取不存在的模板"""
    response = client.get("/api/templates/nonexistent.xlsx")
    assert response.status_code == 404


def test_delete_template_not_found(client):
    """测试删除不存在的模板"""
    response = client.delete("/api/templates/nonexistent.xlsx")
    assert response.status_code == 404
