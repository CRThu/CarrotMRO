from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from datetime import datetime
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import os
import json
import shutil
import asyncio
import uuid
from typing import List, Dict, Any
from pathlib import Path
from io import BytesIO
from config import settings
from photo_ocr import ocr_images
from ratecard_parser import parse_ratecard_file, extract_names
from match import match_names
from quotation_template import (
    load_template,
    get_template_info,
    get_template_columns,
    export_quotation,
    import_quotation,
)

app = FastAPI()

# 允许跨域（本地开发时用）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 数据目录
PROJECT_BASE_DIR = settings.project_dir
RATECARD_BASE_DIR = settings.ratecard_dir
TEMPLATE_DIR = settings.template_dir

# 确保目录存在
PROJECT_BASE_DIR.mkdir(exist_ok=True, parents=True)
RATECARD_BASE_DIR.mkdir(exist_ok=True, parents=True)
TEMPLATE_DIR.mkdir(exist_ok=True, parents=True)

# 在内存中维护任务状态
tasks_db: Dict[str, Dict[str, Any]] = {}

# 定价表 name 列缓存
_names_cache: Dict[str, List[str]] = {}


def invalidate_names(ratecard_name: str):
    _names_cache.pop(ratecard_name, None)


# ========== 后台 OCR 任务函数（仅项目用） ==========

async def run_ocr_task(task_id: str, project_name: str, content_list: List[bytes], columns: List[str]):
    try:
        loop = asyncio.get_running_loop()
        result = await loop.run_in_executor(None, ocr_images, content_list, columns)

        # OCR 识别失败（如 API 限流）时不生成文件
        if not result.get("success"):
            tasks_db[task_id] = {"status": "error", "message": result.get("error", "OCR 识别失败")}
            return

        project_dir = PROJECT_BASE_DIR / project_name
        existing_files = list(project_dir.glob("ocr-*.json"))
        new_index = len(existing_files) + 1
        json_filename = f"ocr-{new_index}.json"
        json_path = project_dir / json_filename

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        tasks_db[task_id] = {"status": "done", "result": result, "file": json_filename, "columns": columns}
    except Exception as e:
        tasks_db[task_id] = {"status": "error", "message": str(e)}


# ========== 工具函数：读取 / 写入定价表 data.json ==========

def get_ratecard_data_path(ratecard_name: str) -> Path:
    return RATECARD_BASE_DIR / f"{ratecard_name}.json"


def read_ratecard_data(ratecard_name: str) -> dict:
    """读取定价表的 data.json，不存在则返回空表"""
    path = get_ratecard_data_path(ratecard_name)
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"columns": [], "items": []}


def write_ratecard_data(ratecard_name: str, data: dict):
    """覆写定价表的 data.json"""
    path = get_ratecard_data_path(ratecard_name)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ========== 路由：协议定价表 ==========

@app.get("/api/ratecards")
async def get_ratecards():
    """获取所有定价表列表"""
    ratecards = []
    if RATECARD_BASE_DIR.exists():
        for f in RATECARD_BASE_DIR.glob("*.json"):
            ratecards.append(f.stem)
    return {"ratecards": ratecards}


@app.post("/api/ratecards/{ratecard_name}")
async def create_ratecard(ratecard_name: str):
    """创建定价表 JSON 文件"""
    path = get_ratecard_data_path(ratecard_name)
    if path.exists():
        raise HTTPException(status_code=400, detail="协议定价表已存在")
    write_ratecard_data(ratecard_name, {"columns": [], "items": []})
    return {"message": "协议定价表创建成功"}


@app.get("/api/ratecards/{ratecard_name}")
async def get_ratecard_data(ratecard_name: str):
    """获取定价表 JSON 内容"""
    path = get_ratecard_data_path(ratecard_name)
    if not path.exists():
        raise HTTPException(status_code=404, detail="定价表不存在")
    return read_ratecard_data(ratecard_name)


@app.post("/api/ratecards/{ratecard_name}/import")
async def import_ratecard(ratecard_name: str, file: UploadFile = File(...)):
    """上传 Excel/CSV 文件，按 !/? 标记解析后写入 JSON"""
    path = get_ratecard_data_path(ratecard_name)
    if not path.exists():
        raise HTTPException(status_code=404, detail="定价表不存在")

    content = await file.read()
    try:
        data = parse_ratecard_file(content, file.filename or "")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    write_ratecard_data(ratecard_name, data)
    invalidate_names(ratecard_name)
    return data


@app.post("/api/match")
async def match_ratecard(body: dict):
    """模糊搜索定价表 name 列"""
    ratecard_name = body.get("ratecard_name")
    queries = body.get("queries", [])
    limit = body.get("limit", 5)

    if not ratecard_name:
        raise HTTPException(status_code=400, detail="ratecard_name 不能为空")
    if not queries:
        raise HTTPException(status_code=400, detail="queries 不能为空")

    path = get_ratecard_data_path(ratecard_name)
    if not path.exists():
        raise HTTPException(status_code=404, detail="定价表不存在")

    names = _names_cache.get(ratecard_name)
    if names is None:
        data = read_ratecard_data(ratecard_name)
        names = extract_names(data.get("columns", []), data.get("items", []))
        _names_cache[ratecard_name] = names

    results = match_names(names, queries, limit=limit)
    return results


# ========== 路由：项目 ==========

@app.get("/api/projects")
async def get_projects():
    projects = []
    if PROJECT_BASE_DIR.exists():
        for d in PROJECT_BASE_DIR.iterdir():
            if d.is_dir():
                projects.append(d.name)
    return {"projects": projects}


@app.get("/api/projects/{project_name}")
async def get_project_info(project_name: str):
    """获取项目基本信息（project.json）"""
    json_path = PROJECT_BASE_DIR / project_name / "project.json"
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


@app.get("/api/projects/{project_name}/ocr-files")
async def get_ocr_files(project_name: str):
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    files = [f.name for f in project_dir.glob("ocr-*.json")]
    files.sort(key=lambda x: int(x.split('-')[1].split('.')[0]))
    return {"files": files}


@app.post("/api/projects/{project_name}")
async def create_project(project_name: str):
    project_dir = PROJECT_BASE_DIR / project_name
    if project_dir.exists():
        raise HTTPException(status_code=400, detail="项目已存在")
    project_dir.mkdir(parents=True)
    project_info = {
        "name": project_name,
        "created_at": datetime.now().isoformat(),
        "ratecard_name": None,
        "template_name": None,
        "selected_columns": [],
    }
    json_path = project_dir / "project.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(project_info, f, ensure_ascii=False, indent=2)
    return {"message": "项目创建成功", "data": project_info}


@app.post("/api/projects/{project_name}/ocr")
async def ocr_images_async(project_name: str, files: List[UploadFile] = File(...)):
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")

    # 读取项目配置获取选中的列
    json_path = project_dir / "project.json"
    if not json_path.exists():
        raise HTTPException(status_code=400, detail="项目配置不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        project_info = json.load(f)
    columns = project_info.get("selected_columns")
    if not columns:
        raise HTTPException(status_code=400, detail="项目未配置识别列，请先关联模板并选择列")

    content_list = [await file.read() for file in files]
    task_id = str(uuid.uuid4())
    tasks_db[task_id] = {"status": "processing"}
    asyncio.create_task(run_ocr_task(task_id, project_name, content_list, columns))
    return {"task_id": task_id}


@app.get("/api/tasks/{task_id}")
async def get_task_status(task_id: str):
    if task_id not in tasks_db:
        return {"status": "not_found"}
    return tasks_db[task_id]


@app.patch("/api/projects/{project_name}/ratecard")
async def update_project_ratecard(project_name: str, body: dict):
    project_dir = PROJECT_BASE_DIR / project_name
    json_path = project_dir / "project.json"
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        project_info = json.load(f)
    project_info["ratecard_name"] = body.get("ratecard_name") or None
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(project_info, f, ensure_ascii=False, indent=2)
    return {"message": "定价表更新成功"}


@app.get("/api/projects/{project_name}/ocr-files/{filename}")
async def get_ocr_data(project_name: str, filename: str):
    json_path = PROJECT_BASE_DIR / project_name / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 获取项目的列配置
    project_json_path = PROJECT_BASE_DIR / project_name / "project.json"
    if project_json_path.exists():
        with open(project_json_path, "r", encoding="utf-8") as f:
            project_info = json.load(f)
        data["columns"] = project_info.get("selected_columns", [])
    else:
        data["columns"] = []

    return data


@app.put("/api/projects/{project_name}/ocr-files/{filename}")
async def save_ocr(project_name: str, filename: str, data: dict):
    json_path = PROJECT_BASE_DIR / project_name / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    output = {
        "success": True,
        "data": data,
        "timestamp": datetime.now().isoformat()
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    return {"message": "保存成功"}


@app.delete("/api/projects/{project_name}/ocr-files/{filename}")
async def delete_ocr_file(project_name: str, filename: str):
    json_path = PROJECT_BASE_DIR / project_name / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    os.remove(json_path)
    return {"message": "删除成功"}


# ========== 路由：报价单 ==========

@app.get("/api/projects/{project_name}/quotations")
async def get_quotations(project_name: str):
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    files = [f.name for f in project_dir.glob("quotation-*.json")]
    files.sort(key=lambda x: int(x.split('-')[1].split('.')[0]))
    return {"files": files}


@app.post("/api/projects/{project_name}/quotations")
async def create_quotation(project_name: str):
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    existing_files = list(project_dir.glob("quotation-*.json"))
    new_index = len(existing_files) + 1
    json_filename = f"quotation-{new_index}.json"
    json_path = project_dir / json_filename
    now = datetime.now().isoformat()
    quotation_data = {
        "columns": [],
        "items": [],
        "created_at": now,
        "last_edit_time": now
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(quotation_data, f, ensure_ascii=False, indent=2)
    return {"message": "报价单创建成功", "file": json_filename}


@app.get("/api/projects/{project_name}/quotations/{filename}")
async def get_quotation_data(project_name: str, filename: str):
    json_path = PROJECT_BASE_DIR / project_name / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


@app.put("/api/projects/{project_name}/quotations/{filename}")
async def save_quotation(project_name: str, filename: str, data: dict):
    json_path = PROJECT_BASE_DIR / project_name / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    data["last_edit_time"] = datetime.now().isoformat()
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"message": "保存成功"}


@app.delete("/api/projects/{project_name}/quotations/{filename}")
async def delete_quotation(project_name: str, filename: str):
    json_path = PROJECT_BASE_DIR / project_name / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    os.remove(json_path)
    return {"message": "删除成功"}


# ========== 路由：模板管理 ==========

@app.get("/api/templates")
async def list_templates():
    """列出所有模板文件"""
    files = []
    if TEMPLATE_DIR.exists():
        for f in TEMPLATE_DIR.glob("*.xlsx"):
            files.append(f.name)
    return {"files": files}


@app.post("/api/templates")
async def upload_template(file: UploadFile = File(...)):
    """上传新模板"""
    if not file.filename or not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="仅支持 .xlsx 文件")
    content = await file.read()
    try:
        template = load_template(content)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"模板解析失败: {e}")
    dest = TEMPLATE_DIR / file.filename
    dest.write_bytes(content)
    return {"message": "模板上传成功", "file": file.filename, "info": get_template_info(template)}


@app.get("/api/templates/{filename}")
async def download_template(filename: str):
    """下载模板文件"""
    path = TEMPLATE_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="模板不存在")
    return StreamingResponse(
        path.open("rb"),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.delete("/api/templates/{filename}")
async def delete_template(filename: str):
    """删除模板"""
    path = TEMPLATE_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="模板不存在")
    os.remove(path)
    return {"message": "删除成功"}


@app.get("/api/templates/{filename}/info")
async def get_template_info_api(filename: str):
    """获取模板解析信息"""
    path = TEMPLATE_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="模板不存在")
    try:
        template = load_template(path.read_bytes())
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"模板解析失败: {e}")
    return get_template_info(template)


# ========== 路由：项目模板关联 ==========

@app.patch("/api/projects/{project_name}/template")
async def update_project_template(project_name: str, body: dict):
    """设置/取消关联模板"""
    json_path = PROJECT_BASE_DIR / project_name / "project.json"
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        project_info = json.load(f)
    template_name = body.get("template_name")
    if template_name:
        template_path = TEMPLATE_DIR / template_name
        if not template_path.exists():
            raise HTTPException(status_code=404, detail="模板不存在")
    project_info["template_name"] = template_name or None
    # 关联模板时清空选中的列（需要重新选择）
    project_info["selected_columns"] = []
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(project_info, f, ensure_ascii=False, indent=2)
    return {"message": "模板关联更新成功"}


@app.patch("/api/projects/{project_name}/columns")
async def update_project_columns(project_name: str, body: dict):
    """设置项目选中的列"""
    json_path = PROJECT_BASE_DIR / project_name / "project.json"
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        project_info = json.load(f)

    # 验证：必须关联模板
    template_name = project_info.get("template_name")
    if not template_name:
        raise HTTPException(status_code=400, detail="项目未关联模板，请先关联模板")

    # 验证：选中的列必须在模板可用列中
    template_path = TEMPLATE_DIR / template_name
    template = load_template(template_path.read_bytes())
    available_columns = get_template_columns(template)

    selected = body.get("columns", [])
    invalid = [c for c in selected if c not in available_columns]
    if invalid:
        raise HTTPException(status_code=400, detail=f"无效的列名: {', '.join(invalid)}")

    project_info["selected_columns"] = selected
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(project_info, f, ensure_ascii=False, indent=2)
    return {"message": "列配置更新成功", "selected_columns": selected}


@app.get("/api/projects/{project_name}/columns")
async def get_project_columns(project_name: str):
    """获取项目的列配置"""
    json_path = PROJECT_BASE_DIR / project_name / "project.json"
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    with open(json_path, "r", encoding="utf-8") as f:
        project_info = json.load(f)

    template_name = project_info.get("template_name")
    available_columns = []
    if template_name:
        template_path = TEMPLATE_DIR / template_name
        if template_path.exists():
            template = load_template(template_path.read_bytes())
            available_columns = get_template_columns(template)

    return {
        "available_columns": available_columns,
        "selected_columns": project_info.get("selected_columns", []),
    }


# ========== 路由：报价单导出 ==========

@app.post("/api/projects/{project_name}/quotations/{filename}/export")
async def export_quotation_excel(project_name: str, filename: str):
    """用关联模板导出报价单为 Excel"""
    project_dir = PROJECT_BASE_DIR / project_name
    json_path = project_dir / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="报价单不存在")

    # 读取项目信息获取关联模板
    project_json_path = project_dir / "project.json"
    if not project_json_path.exists():
        raise HTTPException(status_code=400, detail="项目配置不存在")
    with open(project_json_path, "r", encoding="utf-8") as f:
        project_info = json.load(f)
    template_name = project_info.get("template_name")
    if not template_name:
        raise HTTPException(status_code=400, detail="项目未关联模板，请先关联模板")

    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists():
        raise HTTPException(status_code=404, detail="模板文件不存在")

    # 读取报价单数据
    with open(json_path, "r", encoding="utf-8") as f:
        quotation_data = json.load(f)

    # 构建 groups 数据
    items = quotation_data.get("items", [])
    if not items:
        raise HTTPException(status_code=400, detail="报价单无数据")

    # 按组分组（通过 _group 字段或默认单组）
    groups = _build_groups_from_items(items)

    # 导出
    try:
        template_content = template_path.read_bytes()
        excel_bytes = export_quotation(template_content, groups)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"导出失败: {e}")

    # 生成输出文件名
    base_name = filename.replace(".json", "")
    output_filename = f"{base_name}.xlsx"

    return StreamingResponse(
        BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={output_filename}"},
    )


def _build_groups_from_items(items: list[dict]) -> list[dict]:
    """将报价单 items 按组分组。

    支持两种模式:
    1. items 中包含 _group 字段 → 按 _group 分组
    2. 无 _group 字段 → 所有 items 作为单组
    """
    # 检查是否有 _group 字段
    has_group = any("_group" in item for item in items)

    if not has_group:
        return [{"name": "", "items": items}]

    # 按 _group 分组，保持顺序
    groups: dict[str, list[dict]] = {}
    group_order: list[str] = []
    for item in items:
        group_name = item.get("_group", "")
        if group_name not in groups:
            groups[group_name] = []
            group_order.append(group_name)
        # 复制 item 并移除 _group 字段
        clean_item = {k: v for k, v in item.items() if not k.startswith("_")}
        groups[group_name].append(clean_item)

    return [{"name": name, "items": groups[name]} for name in group_order]


# ========== 路由：报价单导入 ==========

@app.post("/api/projects/{project_name}/quotations/import")
async def import_quotation_excel(project_name: str, file: UploadFile = File(...)):
    """上传 Excel 解析为新报价单"""
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")

    # 读取项目信息获取关联模板
    project_json_path = project_dir / "project.json"
    if not project_json_path.exists():
        raise HTTPException(status_code=400, detail="项目配置不存在")
    with open(project_json_path, "r", encoding="utf-8") as f:
        project_info = json.load(f)
    template_name = project_info.get("template_name")
    if not template_name:
        raise HTTPException(status_code=400, detail="项目未关联模板，请先关联模板")

    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists():
        raise HTTPException(status_code=404, detail="模板文件不存在")

    # 读取上传的 Excel
    excel_content = await file.read()

    # 导入解析
    try:
        template_content = template_path.read_bytes()
        groups = import_quotation(template_content, excel_content)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"导入解析失败: {e}")

    if not groups:
        raise HTTPException(status_code=400, detail="未解析到有效数据")

    # 将 groups 转换为 items（扁平化，带 _group 字段）
    items = []
    for group in groups:
        for item in group.get("items", []):
            items.append({**item, "_group": group.get("name", "")})

    # 创建新报价单
    existing_files = list(project_dir.glob("quotation-*.json"))
    new_index = len(existing_files) + 1
    json_filename = f"quotation-{new_index}.json"
    json_path = project_dir / json_filename
    now = datetime.now().isoformat()
    quotation_data = {
        "columns": [],
        "items": items,
        "created_at": now,
        "last_edit_time": now,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(quotation_data, f, ensure_ascii=False, indent=2)

    return {"message": "导入成功", "file": json_filename, "item_count": len(items)}


# 挂载前端静态文件 (打包后)
static_dir = os.path.join(os.path.dirname(__file__), "static")
if os.path.isdir(static_dir):
    app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
