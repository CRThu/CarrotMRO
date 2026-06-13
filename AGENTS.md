# CarrotMRO 项目说明

CarrotMRO 是一个基于 OCR 技术的报价单自动处理系统，旨在通过识别图片形式的报价单，将其数据化，并生成标准格式的 Excel 报告。

## 系统架构

本系统采用前后端分离架构：
- **后端 (Backend)**: 基于 FastAPI 构建，提供核心 API 接口，负责 OCR 处理、项目管理、数据存储及 Excel 生成。
- **前端 (Frontend)**: 基于 React 构建，提供用户界面，用于创建项目和上传需要 OCR 的图片。

## 目录结构

```text
.
├── backend/            # 后端 FastAPI 应用
│   ├── main.py         # 核心 API 路由与服务入口
│   └── photo_ai.py     # OCR 识别逻辑封装
├── frontend/           # 前端 React 应用
│   └── src/            # 前端源代码
├── {DATA_ROOT}/        # 数据根目录 (环境变量配置)
│   ├── project/        # 项目数据存储 (环境变量 PROJECT_SUBDIR 配置)
│   └── template/       # 模板文件存储 (环境变量 TEMPLATE_SUBDIR 配置)
└── (其他根目录脚本)     # 辅助工具和运行脚本
```

## 核心功能

### 1. 项目管理
- **功能**: 创建新项目以隔离不同业务的数据。
- **API**: `PUT /api/create-project/{project_name}`
- **说明**: 创建后会在 `{DATA_ROOT}/{PROJECT_SUBDIR}/` 下生成 `{project_name}.json` 文件，用于存放项目的持久化状态信息。

### 2. OCR 识别
- **功能**: 上传图片并识别其中的文字内容。
- **API**: `POST /api/ocr`
- **说明**: 接收图片文件，调用 `photo_ai.py` 进行识别，并将结果保存至对应项目的 JSON 文件中。

### 3. XLSX 导出
- **功能**: 将识别后的数据导出为结构化的 Excel 报表。
- **API**: `POST /api/export-xlsx`
- **说明**: 利用 `project-template/` 中的模板生成最终报告。

## 运行指南

1.  **环境准备**: 确保已安装 Python 环境及所需的依赖库（请查阅 `pyproject.toml`）。
2.  **启动**:
    - 开发环境下，可分别启动后端服务和前端开发服务器。
    - 使用提供的 `run.bat` 脚本可以快速启动整个系统。

## 数据流向
用户在前端创建项目 -> 后端生成 `data/projectname.json` -> 用户上传图片 -> 后端调用 OCR 识别 -> 识别数据存入 `data/projectname.json` -> 用户触发导出 -> 后端生成 Excel 报表。
