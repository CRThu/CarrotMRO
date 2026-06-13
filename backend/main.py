from fastapi import FastAPI, UploadFile, File, HTTPException
from datetime import datetime
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import os
import json
import shutil
from pathlib import Path
from typing import List
from photo_ocr import recognize_photos_base64

app = FastAPI()

# 允许跨域（本地开发时用）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 数据根目录
DATA_ROOT = Path(os.getenv("DATA_ROOT", "data"))
# 项目数据存放目录 (以前是 project, 现在改为 projects)
PROJECT_BASE_DIR = DATA_ROOT / os.getenv("PROJECT_SUBDIR", "projects")

# 确保目录存在
PROJECT_BASE_DIR.mkdir(exist_ok=True, parents=True)
TEMPLATE_DIR = DATA_ROOT / os.getenv("TEMPLATE_SUBDIR", "template")

# 确保模板目录存在
TEMPLATE_DIR.mkdir(exist_ok=True, parents=True)

# 路由：获取所有项目
@app.get("/api/projects")
async def get_projects():
    projects = []
    # 查找所有的项目目录
    if PROJECT_BASE_DIR.exists():
        for d in PROJECT_BASE_DIR.iterdir():
            if d.is_dir():
                projects.append(d.name)
    return {"projects": projects}

import asyncio
import uuid
from typing import Dict, Any

# 在内存中维护任务状态 (实际生产环境建议用 Redis 或 SQLite 替代)
tasks_db: Dict[str, Dict[str, Any]] = {}

# 真正的 OCR 执行逻辑
async def run_ocr_task(task_id: str, project_name: str, content_list: List[bytes]):
    try:
        # 在子线程运行耗时的同步 OCR 操作，避免阻塞事件循环
        loop = asyncio.get_running_loop()
        result = await loop.run_in_executor(None, recognize_photos_base64, content_list)
        
        project_dir = PROJECT_BASE_DIR / project_name
        existing_files = list(project_dir.glob("ocr-*.json"))
        new_index = len(existing_files) + 1
        json_filename = f"ocr-{new_index}.json"
        json_path = project_dir / json_filename
        
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            
        tasks_db[task_id] = {"status": "done", "result": result, "file": json_filename}
    except Exception as e:
        tasks_db[task_id] = {"status": "error", "message": str(e)}

# 路由：OCR 提交任务
@app.post("/api/ocr/{project_name}")
async def ocr_images_async(project_name: str, files: List[UploadFile] = File(...)):
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
        
    content_list = [await file.read() for file in files]
    task_id = str(uuid.uuid4())
    tasks_db[task_id] = {"status": "processing"}
    
    # 在后台启动任务
    asyncio.create_task(run_ocr_task(task_id, project_name, content_list))
    
    return {"task_id": task_id}

# 路由：查询任务状态
@app.get("/api/task-status/{task_id}")
async def get_task_status(task_id: str):
    if task_id not in tasks_db:
        return {"status": "not_found"}
    return tasks_db[task_id]

# 路由：创建新项目
@app.put("/api/create-project/{project_name}")
async def create_project(project_name: str):
    project_dir = PROJECT_BASE_DIR / project_name
    
    if project_dir.exists():
        raise HTTPException(status_code=400, detail="项目已存在")
    
    # 创建项目目录
    project_dir.mkdir(parents=True)
    
    # 在目录下创建 project.json
    project_info = {
        "name": project_name,
        "created_at": datetime.now().isoformat()
    }
    
    json_path = project_dir / "project.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(project_info, f, ensure_ascii=False, indent=2)
        
    return {"message": "项目创建成功", "data": project_info}

# 路由：获取项目下的所有 OCR 记录文件
@app.get("/api/ocr-files/{project_name}")
async def get_ocr_files(project_name: str):
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
    
    files = [f.name for f in project_dir.glob("ocr-*.json")]
    # 按名称排序，通常是 ocr-1, ocr-2
    files.sort(key=lambda x: int(x.split('-')[1].split('.')[0]))
    return {"files": files}

# 路由：获取具体的 OCR 结果内容
@app.get("/api/ocr-data/{project_name}/{filename}")
async def get_ocr_data(project_name: str, filename: str):
    json_path = PROJECT_BASE_DIR / project_name / filename
    if not json_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data

@app.post("/api/save-ocr/{project_name}/{filename}")
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

# 挂载前端静态文件 (打包后)
static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
