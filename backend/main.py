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

# 路由：OCR 接口
@app.post("/api/ocr/{project_name}")
async def ocr_images(project_name: str, files: List[UploadFile] = File(...)):
    # 检查项目目录是否存在
    project_dir = PROJECT_BASE_DIR / project_name
    if not project_dir.exists():
        raise HTTPException(status_code=404, detail="项目不存在")
        
    # 读取所有图片数据
    content_list = []
    for file in files:
        content_list.append(await file.read())
    
    # 调用 photo_ocr 的批量接口
    result = recognize_photos_base64(content_list)
    
    # 确定文件命名 (ocr-1.json, ocr-2.json, ...)
    existing_files = list(project_dir.glob("ocr-*.json"))
    new_index = len(existing_files) + 1
    json_filename = f"ocr-{new_index}.json"
    json_path = project_dir / json_filename
    
    # 保存结果
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
        
    return {"message": "OCR完成", "data": result, "file": str(json_path)}

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
