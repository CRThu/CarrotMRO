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
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-3.1-flash-lite")
# 专注于任务指令，不再包含 JSON 结构描述
PROMPT_INSTRUCTIONS = os.getenv("GEMINI_PROMPT", """
你是一个专业的工程报价单OCR识别助手。
你的任务是阅读并理解图片中的报价单内容，提取其中的项目信息。

任务要求：
1. 提取项目名称、数量、单位、参考单价（若存在）。
2. 不要对项目进行序号编号。
3. 不要遗漏项目，也不要虚构项目。
4. 将所有不确定或有疑问的内容汇总在备注字段中。
""")

# 定义严格的响应 Schema (强制约束格式)
RESPONSE_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "items": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "name": types.Schema(type=types.Type.STRING, description="项目名称"),
                    "quantity": types.Schema(type=types.Type.STRING, description="数量"),
                    "unit": types.Schema(type=types.Type.STRING, description="单位"),
                    "unit_price": types.Schema(type=types.Type.NUMBER, description="参考单价，若无则为null", nullable=True),
                },
                required=["name", "quantity", "unit"],
            ),
        ),
        "remarks": types.Schema(type=types.Type.STRING, description="不确定的或有疑问的内容"),
    },
    required=["items", "remarks"],
)

# 初始化 Client
client = genai.Client(api_key=API_KEY) if API_KEY and API_KEY != "你的API_KEY" else None

def _ensure_configured():
    if client is None:
        raise RuntimeError("Gemini 未配置！请在 .env 中设置 GEMINI_API_KEY")

def recognize_photos_base64(image_list: list[bytes]) -> dict:
    """通过二进制数据列表识别多张图片"""
    _ensure_configured()
    
    # 构造内容列表
    contents = [types.Content(role="user", parts=[types.Part.from_text(text=PROMPT_INSTRUCTIONS)])]
    for img_data in image_list:
        contents[0].parts.append(types.Part.from_bytes(data=img_data, mime_type="image/jpeg"))
    
    try:
        # 使用新的 SDK 生成内容
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=RESPONSE_SCHEMA,
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
