"""photo_ocr.py 测试"""

import json
from unittest.mock import patch, MagicMock

import pytest

from photo_ocr import _build_prompt, _extract_json, ocr_images


class TestBuildPrompt:
    def test_build_prompt_with_columns(self):
        """测试构建 prompt"""
        prompt = _build_prompt(["项目名称", "数量", "单位"])
        assert "项目名称" in prompt
        assert "数量" in prompt
        assert "单位" in prompt
        assert "JSON" in prompt


class TestExtractJson:
    def test_extract_valid_json(self):
        """测试提取有效 JSON"""
        text = '{"items": [{"name": "test"}], "remarks": ""}'
        result = _extract_json(text)
        assert result is not None
        assert "items" in result

    def test_extract_json_in_code_block(self):
        """测试从代码块中提取 JSON"""
        text = '```json\n{"items": [{"name": "test"}], "remarks": ""}\n```'
        result = _extract_json(text)
        assert result is not None
        assert "items" in result

    def test_extract_json_with_surrounding_text(self):
        """测试从周围文本中提取 JSON"""
        text = 'Here is the result:\n{"items": [{"name": "test"}], "remarks": ""}\nDone.'
        result = _extract_json(text)
        assert result is not None
        assert "items" in result

    def test_extract_invalid_json(self):
        """测试无效 JSON 返回 None"""
        result = _extract_json("not json at all")
        assert result is None


class TestOcrImages:
    def test_no_columns_returns_error(self):
        """测试未配置列时返回错误"""
        result = ocr_images([b"fake_image"], [])
        assert result["success"] is False
        assert "未配置识别列" in result["error"]

    def test_no_api_key_raises_error(self):
        """测试未配置 API Key 时抛出异常"""
        with patch("photo_ocr.settings") as mock_settings:
            mock_settings.xiaomi_mimo_api_key = ""
            with pytest.raises(RuntimeError, match="MiMo 未配置"):
                ocr_images([b"fake_image"], ["项目名称"])

    @patch("builtins.__import__")
    def test_ocr_success(self, mock_import):
        """测试 OCR 成功"""
        mock_litellm = MagicMock()
        mock_response = MagicMock()
        mock_response.choices = [MagicMock()]
        mock_response.choices[0].message.content = json.dumps({
            "items": [{"项目名称": "螺栓M10", "数量": "100", "单位": "个"}],
            "remarks": ""
        })
        mock_litellm.completion.return_value = mock_response

        def import_side_effect(name, *args, **kwargs):
            if name == "litellm":
                return mock_litellm
            return __builtins__.__import__(name, *args, **kwargs) if hasattr(__builtins__, '__import__') else __import__(name, *args, **kwargs)

        mock_import.side_effect = import_side_effect

        with patch("photo_ocr.settings") as mock_settings:
            mock_settings.xiaomi_mimo_api_key = "test_key"
            mock_settings.litellm_model = "xiaomi_mimo/mimo-v2-flash"
            result = ocr_images([b"\xff\xd8\xff\xe0fake_jpeg"], ["项目名称", "数量", "单位"])

        assert result["success"] is True
        assert "data" in result
        assert len(result["data"]["items"]) == 1
