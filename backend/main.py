from fastapi import FastAPI, UploadFile, File, HTTPException
from datetime import datetime
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import os
import json
import shutil
from pathlib import Path
# 导入你的核心模块
from photo_ai import recognize_photo_base64

app = FastAPI()

# 允许跨域（本地开发时用）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

DATA_ROOT = Path(os.getenv("DATA_ROOT", "data"))
PROJECT_DIR = DATA_ROOT / os.getenv("PROJECT_SUBDIR", "project")
TEMPLATE_DIR = DATA_ROOT / os.getenv("TEMPLATE_SUBDIR", "template")

# 确保目录存在
PROJECT_DIR.mkdir(exist_ok=True, parents=True)
TEMPLATE_DIR.mkdir(exist_ok=True, parents=True)

# 路由：获取所有项目
@app.get("/api/projects")
async def get_projects():
    projects = []
    for f in PROJECT_DIR.glob("*.json"):
        projects.append(f.stem)
    return {"projects": projects}

# 路由：OCR 接口
@app.post("/api/ocr")
async def ocr_image(file: UploadFile = File(...)):
    # 读取图片数据
    content = await file.read()
    
    # 调用 photo_ai 的 base64 接口
    result = recognize_photo_base64(content)
    
    # 保存结果到本地 JSON
    json_filename = f"{Path(file.filename).stem}.json"
    json_path = PROJECT_DIR / json_filename
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
        
    return {"message": "OCR完成", "data": result, "file": str(json_path)}

# 路由：创建新项目
@app.put("/api/create-project/{project_name}")
async def create_project(project_name: str):
    json_path = PROJECT_DIR / f"{project_name}.json"
    
    if json_path.exists():
        raise HTTPException(status_code=400, detail="Project already exists")
        
    project_data = {
        "name": project_name,
        "created_at": datetime.now().isoformat()
    }
    
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(project_data, f, ensure_ascii=False, indent=2)
        
    return {"message": "项目创建成功", "data": project_data}

# 路由：简单模拟 XLSX 生成（后续接入你的逻辑）
@app.post("/api/export-xlsx")
async def export_xlsx():
    # 这里接入你的 process_excel 逻辑
    return {"message": "XLSX 生成成功", "download_url": "/downloads/report.xlsx"}

# 挂载前端静态文件 (打包后)
static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
