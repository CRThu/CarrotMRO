"""
照片上传 Google Gemini AI 识别，返回 JSON。

用法:
    from photo_ai import recognize_photo

    result = recognize_photo("照片.jpg")
    print(result)  # JSON dict

配置在 .env 中:
    GEMINI_API_KEY  - API Key（必填）
    GEMINI_PROMPT   - 你调好的提示词（必填）
    PROXY           - 代理地址（可选）
"""

import os
import json
import re
from pathlib import Path
from datetime import datetime
from typing import Any

from dotenv import load_dotenv
import google.generativeai as genai

# ══════════════════════════════════════════════
# 加载 .env 配置
# ══════════════════════════════════════════════
load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY", "")
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-2.0-flash")
PROMPT = os.getenv("GEMINI_PROMPT", "请详细描述这张图片的内容")

# 代理设置（一个 PROXY 变量同时设置 http 和 https）
PROXY = os.getenv("PROXY", "")
if PROXY:
    os.environ["http_proxy"] = PROXY
    os.environ["https_proxy"] = PROXY

# ══════════════════════════════════════════════
# 初始化 Gemini
# ══════════════════════════════════════════════
if API_KEY and API_KEY != "你的API_KEY":
    genai.configure(api_key=API_KEY)
    _model = genai.GenerativeModel(MODEL_NAME)
else:
    _model = None


def _ensure_configured():
    if _model is None:
        raise RuntimeError(
            "Gemini 未配置！请先在 .env 中设置 GEMINI_API_KEY=你的API_KEY\n"
            "获取 Key: https://aistudio.google.com/app/apikey"
        )


def recognize_photo(
    photo_path: str | Path,
    json_mode: bool = True,
    response_schema: dict | None = None,
) -> dict:
    """
    上传一张照片给 Gemini AI 识别，返回 JSON 结果。

    json_mode=True 时使用 response_mime_type="application/json"
    效果等同于 AI Studio 的 JSON 模式开关，模型会严格输出合法 JSON。

    参数:
        photo_path:     照片文件路径 (jpg/png/webp/heic)
        json_mode:      是否强制 JSON 输出（默认 True，类似 AI Studio 的 JSON 模式）
        response_schema: 可选的 JSON Schema 定义输出结构（更精确）

    返回:
        {
            "success": True/False,
            "photo": "文件名",
            "timestamp": "...",
            "model": "gemini-2.0-flash",
            "prompt": ".env 中配置的提示词",
            "data": {JSON对象} 或 None,   # json_mode=True 时直接是解析好的 dict
            "error": None 或 "错误信息"
        }
    """
    _ensure_configured()

    path = Path(photo_path)
    if not path.exists():
        return {
            "success": False,
            "photo": str(path),
            "error": f"文件不存在: {photo_path}",
        }

    try:
        import PIL.Image
        img = PIL.Image.open(path)
    except Exception as e:
        return {
            "success": False,
            "photo": str(path),
            "error": f"图片读取失败: {e}",
        }

    # =============================================
    # 关键：response_mime_type 强制 JSON 输出
    # 效果 = AI Studio 里的 JSON 模式开关
    # =============================================
    generation_config = {}
    if json_mode:
        generation_config["response_mime_type"] = "application/json"
        if response_schema:
            generation_config["response_schema"] = response_schema

    try:
        response = _model.generate_content(
            [PROMPT, img],
            generation_config=generation_config,
        )
        text = response.text.strip()
    except Exception as e:
        return {
            "success": False,
            "photo": path.name,
            "error": f"API 调用失败: {e}",
        }

    result = {
        "success": True,
        "photo": path.name,
        "timestamp": datetime.now().isoformat(),
        "model": MODEL_NAME,
        "prompt": PROMPT,
        "data": None,
        "error": None,
    }

    # JSON 模式下，API 保证返回合法 JSON
    if json_mode:
        try:
            result["data"] = json.loads(text)
        except json.JSONDecodeError as e:
            result["data"] = {"raw_text": text}
            result["warning"] = f"JSON 解析失败: {e}"
    else:
        result["text"] = text

    return result


def recognize_photo_base64(
    image_data: bytes,
    json_mode: bool = True,
    response_schema: dict | None = None,
) -> dict:
    """
    通过二进制字节数据识别图片（适合摄像头/前端直接传入）。

    参数同 recognize_photo()。
    """
    _ensure_configured()

    try:
        import PIL.Image
        from io import BytesIO
        img = PIL.Image.open(BytesIO(image_data))
    except Exception as e:
        return {
            "success": False,
            "photo": "base64_input",
            "error": f"图片字节数据解析失败: {e}",
        }

    generation_config = {}
    if json_mode:
        generation_config["response_mime_type"] = "application/json"
        if response_schema:
            generation_config["response_schema"] = response_schema

    try:
        response = _model.generate_content(
            [PROMPT, img],
            generation_config=generation_config,
        )
        text = response.text.strip()
    except Exception as e:
        return {
            "success": False,
            "photo": "base64_input",
            "error": f"API 调用失败: {e}",
        }

    result = {
        "success": True,
        "photo": "base64_input",
        "timestamp": datetime.now().isoformat(),
        "model": MODEL_NAME,
        "prompt": PROMPT,
        "data": None,
        "error": None,
    }

    if json_mode:
        try:
            result["data"] = json.loads(text)
        except json.JSONDecodeError as e:
            result["data"] = {"raw_text": text}
            result["warning"] = f"JSON 解析失败: {e}"
    else:
        result["text"] = text

    return result


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: uv run python photo_ai.py <图片路径>")
        sys.exit(1)
    path = sys.argv[1]
    result = recognize_photo(path, json_mode=False)
    print(json.dumps(result, ensure_ascii=False, indent=2))
