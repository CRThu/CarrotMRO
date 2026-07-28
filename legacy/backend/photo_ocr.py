"""
照片上传 OCR 识别，返回 JSON。
使用 litellm 调用 Xiaomi MiMo 模型。
"""

import json
import re
import base64
from datetime import datetime
from config import settings


def _ensure_configured():
    if not settings.xiaomi_mimo_api_key or settings.xiaomi_mimo_api_key == "你的API_KEY":
        raise RuntimeError("MiMo 未配置！请在 .env 中设置 XIAOMI_MIMO_API_KEY")


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

请严格按照以下 JSON 格式返回结果，不要包含其他内容：
{{
  "items": [
    {{{", ".join(f'"{col}": ""' for col in columns)}}}
  ],
  "remarks": ""
}}
"""


def _extract_json(text: str) -> dict | None:
    """从响应文本中提取 JSON"""
    # 尝试直接解析
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 尝试提取 ```json ... ``` 块
    match = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass

    # 尝试提取 { ... } 块
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass

    return None


def ocr_images(image_list: list[bytes], columns: list[str]) -> dict:
    """通过二进制数据列表识别多张图片

    Args:
        image_list: 图片二进制数据列表
        columns: 要识别的列名列表，如 ["物料名称", "数量", "单位", "单价"]
                 必须提供，否则报错
    """
    if not columns:
        return {"success": False, "error": "未配置识别列，请先在项目中选择列"}

    _ensure_configured()

    import litellm

    prompt = _build_prompt(columns)

    # 构造消息：text + images
    content: list[dict] = [{"type": "text", "text": prompt}]
    for img_data in image_list:
        b64 = base64.b64encode(img_data).decode()
        # 检测图片格式
        mime_type = "image/jpeg"
        if img_data[:4] == b'\x89PNG':
            mime_type = "image/png"
        elif img_data[:4] == b'RIFF':
            mime_type = "image/webp"
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:{mime_type};base64,{b64}"}
        })

    try:
        response = litellm.completion(
            model=settings.litellm_model,
            messages=[{"role": "user", "content": content}],
            temperature=0.1,
        )

        text = response.choices[0].message.content
        data = _extract_json(text)

        if data is None:
            return {"success": False, "error": f"无法解析模型返回的 JSON: {text[:200]}"}

        # 确保返回格式正确
        if "items" not in data:
            data = {"items": data if isinstance(data, list) else [data], "remarks": ""}
        if "remarks" not in data:
            data["remarks"] = ""

        return {
            "success": True,
            "data": data,
            "timestamp": datetime.now().isoformat()
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


if __name__ == "__main__":
    print("请通过后端接口调用 OCR。")
