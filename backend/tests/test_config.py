"""config.py 测试"""

from pathlib import Path
from unittest.mock import patch


def test_settings_defaults():
    """测试默认配置值"""
    from config import Settings
    settings = Settings()
    assert settings.litellm_model == "xiaomi_mimo/mimo-v2-flash"
    assert settings.project_subdir == "projects"
    assert settings.template_subdir == "template"


def test_data_dir_default():
    """测试默认数据目录（当 DATA_ROOT 未设置时）"""
    with patch.dict("os.environ", {"DATA_ROOT": ""}, clear=False):
        # 清除缓存的 settings
        import importlib
        import config
        importlib.reload(config)
        settings = config.Settings()
        # 默认应该是 backend 的父目录下的 data
        expected = Path(__file__).resolve().parent.parent.parent / "data"
        assert settings.data_dir == expected


def test_data_dir_custom():
    """测试自定义数据目录"""
    with patch.dict("os.environ", {"DATA_ROOT": "/tmp/test_data"}):
        from config import Settings
        import importlib
        import config
        importlib.reload(config)
        settings = config.Settings()
        assert settings.data_dir == Path("/tmp/test_data")


def test_project_dir():
    """测试项目目录"""
    from config import Settings
    settings = Settings()
    assert settings.project_dir == settings.data_dir / "projects"


def test_ratecard_dir():
    """测试定价表目录"""
    from config import Settings
    settings = Settings()
    assert settings.ratecard_dir == settings.data_dir / "ratecard"


def test_template_dir():
    """测试模板目录"""
    from config import Settings
    settings = Settings()
    assert settings.template_dir == settings.data_dir / "template"
