"""
照片上传 Google Gemini AI 识别，返回 JSON。
使用最新的 google.genai SDK。
"""

import json
from datetime import datetime
from google import genai
from google.genai import types
from config import settings

API_KEY = settings.gemini_api_key
MODEL_NAME = settings.gemini_model

# 初始化 Client
client = genai.Client(api_key=API_KEY) if API_KEY and API_KEY != "你的API_KEY" else None


def _ensure_configured():
    if client is None:
        raise RuntimeError("Gemini 未配置！请在 .env 中设置 GEMINI_API_KEY")


def _build_prompt(columns: list[str]) -> str:
    """根据列名动态生成 prompt"""
    col_list = "、".join(columns)
    return f"""你是一个专业的工程报价单OCR识别助手。
你的任务是阅读并理解图片中的报价单内容，提取其中的项目信息。

任务要求：
1. 提取以下字段：{col_list}。
2. 不要对项目进行序号编号。
3. 不要遗漏项目，也不要虚构项目。
4. 将所有不确定或有疑问的内容汇总在备注字段中。
"""


def _build_schema(columns: list[str]) -> types.Schema:
    """根据列名动态生成响应 Schema"""
    # 判断哪些列是必填的（数量、单位通常是必填）
    required_columns = [c for c in columns if c in ("数量", "单位")]

    properties = {}
    for col in columns:
        if col in ("单价", "参考单价", "综合单价"):
            properties[col] = types.Schema(
                type=types.Type.NUMBER,
                description=col,
                nullable=True,
            )
        else:
            properties[col] = types.Schema(
                type=types.Type.STRING,
                description=col,
            )

    return types.Schema(
        type=types.Type.OBJECT,
        properties={
            "items": types.Schema(
                type=types.Type.ARRAY,
                items=types.Schema(
                    type=types.Type.OBJECT,
                    properties=properties,
                    required=required_columns,
                ),
            ),
            "remarks": types.Schema(type=types.Type.STRING, description="不确定的或有疑问的内容"),
        },
        required=["items", "remarks"],
    )


def ocr_images(image_list: list[bytes], columns: list[str] | None = None) -> dict:
    """通过二进制数据列表识别多张图片

    Args:
        image_list: 图片二进制数据列表
        columns: 要识别的列名列表，如 ["物料名称", "数量", "单位", "单价"]
                 若为 None 则使用默认列名
    """
    _ensure_configured()

    # 默认列名（兼容旧调用）
    if not columns:
        columns = ["物料名称", "数量", "单位", "单价"]

    prompt = _build_prompt(columns)
    schema = _build_schema(columns)

    # 构造内容列表
    contents = [types.Content(role="user", parts=[types.Part.from_text(text=prompt)])]
    for img_data in image_list:
        contents[0].parts.append(types.Part.from_bytes(data=img_data, mime_type="image/jpeg"))

    try:
        # 使用新的 SDK 生成内容
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=schema,
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
