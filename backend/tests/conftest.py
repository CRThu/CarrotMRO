import sys
from pathlib import Path
from unittest.mock import patch

import pytest
from fastapi.testclient import TestClient

# 确保 backend 目录在 sys.path 中
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


@pytest.fixture
def client(tmp_path):
    """创建测试客户端，使用临时数据目录"""
    with patch("config.settings") as mock_settings:
        mock_settings.data_dir = tmp_path
        mock_settings.project_dir = tmp_path / "projects"
        mock_settings.ratecard_dir = tmp_path / "ratecard"
        mock_settings.template_dir = tmp_path / "template"
        mock_settings.project_subdir = "projects"
        mock_settings.template_subdir = "template"
        mock_settings.xiaomi_mimo_api_key = ""
        mock_settings.litellm_model = "xiaomi_mimo/mimo-v2-flash"

        # 创建必要的目录
        (tmp_path / "projects").mkdir(exist_ok=True)
        (tmp_path / "ratecard").mkdir(exist_ok=True)
        (tmp_path / "template").mkdir(exist_ok=True)

        from main import app
        yield TestClient(app)


@pytest.fixture
def sample_ratecard_data():
    """示例定价表数据"""
    return {
        "columns": [
            {"name": "项目名称", "strict": True, "alias": None},
            {"name": "单位", "strict": True, "alias": None},
            {"name": "综合单价", "strict": True, "alias": None},
        ],
        "items": [
            {"项目名称": "螺栓M10", "单位": "个", "综合单价": "2.50"},
            {"项目名称": "垫片M10", "单位": "个", "综合单价": "0.80"},
        ],
    }
