"""全局预制列定义。

预制列是系统中固定的标准列清单，用户在项目配置中手动将模板列映射到这些预制列。
每个场景（OCR、定价表、报价单）各自有独立的映射。
"""

PRESET_COLUMNS = [
    {"key": "name",        "label": "项目名称", "required": True,  "type": "string"},
    {"key": "quantity",    "label": "数量",     "required": True,  "type": "number"},
    {"key": "unit",        "label": "单位",     "required": True,  "type": "string"},
    {"key": "unit_price",  "label": "单价",     "required": False, "type": "number"},
    {"key": "total_price", "label": "合价",     "required": False, "type": "number", "computed": True},
    {"key": "spec",        "label": "规格型号", "required": False, "type": "string"},
    {"key": "brand",       "label": "品牌",     "required": False, "type": "string"},
    {"key": "remark",      "label": "备注",     "required": False, "type": "string"},
]

# 映射场景类型
MAPPING_SCOPES = ["ocr", "ratecard", "quotation"]


def get_preset_columns() -> list[dict]:
    return PRESET_COLUMNS


def get_preset_labels() -> list[str]:
    return [col["label"] for col in PRESET_COLUMNS]
