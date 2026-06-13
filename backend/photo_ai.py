"""
照片上传 Google Gemini AI 识别，返回 JSON。
使用最新的 google.genai SDK。
"""

import os
import json
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from google import genai
from google.genai import types

# 加载 .env 配置
load_dotenv(dotenv_path=Path(__file__).parent.parent / ".env")

API_KEY = os.getenv("GEMINI_API_KEY", "")
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-2.0-flash")
PROMPT = os.getenv("GEMINI_PROMPT", "请详细描述这张图片的内容")

# 初始化 Client
client = genai.Client(api_key=API_KEY) if API_KEY and API_KEY != "你的API_KEY" else None

def _ensure_configured():
    if client is None:
        raise RuntimeError("Gemini 未配置！请在 .env 中设置 GEMINI_API_KEY")

def recognize_photo_base64(image_data: bytes) -> dict:
    """通过二进制数据识别图片"""
    _ensure_configured()
    
    try:
        # 使用新的 SDK 生成内容
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_bytes(data=image_data, mime_type="image/jpeg"),
                        types.Part.from_text(text=PROMPT),
                    ],
                )
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
            ),
        )
        
        result = {
            "success": True,
            "data": json.loads(response.text),
            "timestamp": datetime.now().isoformat()
        }
        return result
    except Exception as e:
        return {"success": False, "error": str(e)}

if __name__ == "__main__":
    print("SDK 已更新，请通过后端接口调用。")
