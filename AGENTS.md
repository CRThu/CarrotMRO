# CarrotMRO 项目说明

CarrotMRO 是一个 MRO（维护、维修、运营）综合管理系统，包含多个业务模块。当前系统包含**项目管理**（含 OCR 图片识别报价单）和**协议定价表管理**两大核心模块。

## 系统架构

本系统采用前后端分离架构：
- **后端 (Backend)**: 基于 FastAPI 构建，提供核心 API 接口，负责 OCR 处理、项目管理、定价表管理、数据存储及 Excel 生成。
- **前端 (Frontend)**: 基于 React 构建，提供用户界面，包含项目工作台和定价表编辑器两个独立视图。

## 目录结构

```text
.
├── backend/            # 后端 FastAPI 应用
│   ├── main.py         # 核心 API 路由与服务入口
│   ├── photo_ocr.py    # OCR 识别逻辑封装
│   ├── match.py        # 项目名称匹配逻辑
│   └── file_utils.py   # 文件读写辅助工具
├── frontend/           # 前端 React 应用
│   └── src/            # 前端源代码
├── data/               # 数据存储目录
│   ├── projects/       # 项目数据及 OCR 结果存储
│   └── ratecard/       # 协议定价表 JSON 文件（*.json）
├── run.bat             # 快速启动脚本
├── build.bat           # 前端构建脚本
└── (其他根目录脚本)     # 辅助工具
```

## 核心功能

### 1. 项目管理
- **功能**: 创建项目、上传报价单图片进行 OCR 识别、查看和编辑识别结果。
- **API**:
  | 方法 | 路径 | 说明 |
  |------|------|------|
  | `GET` | `/api/projects` | 获取所有项目列表 |
  | `POST` | `/api/projects/{name}` | 创建新项目（生成目录 + `project.json`） |
  | `GET` | `/api/projects/{name}` | 获取项目基本信息 |
  | `PATCH` | `/api/projects/{name}/ratecard` | 关联/解除关联定价表 |
  | `POST` | `/api/projects/{name}/ocr` | 上传图片启动 OCR 异步识别 |
  | `GET` | `/api/tasks/{task_id}` | 轮询 OCR 任务状态 |
  | `GET` | `/api/projects/{name}/ocr-files` | 获取项目下的 OCR 结果文件列表 |
  | `GET` | `/api/projects/{name}/ocr-files/{file}` | 读取 OCR 结果数据 |
  | `PUT` | `/api/projects/{name}/ocr-files/{file}` | 保存 OCR 表格编辑结果 |
  | `DELETE` | `/api/projects/{name}/ocr-files/{file}` | 删除指定的 OCR 结果文件 |
- **说明**: 创建后会在 `data/projects/{project_name}/` 下生成 `project.json`，记录项目名称、创建时间和关联的定价表名称。OCR 识别成功后，结果存储为 `ocr-1.json`、`ocr-2.json` ... 每个文件包含识别出的物料清单表格数据。若 OCR 识别失败（如 API 限流），**不会生成结果文件**，任务状态返回 `{"status": "error", "message": "..."}`。

### 2. 协议定价表管理
- **功能**: 创建和管理协议定价表，支持从 Excel/CSV 文件导入数据。导入后可在页面上预览，如需修改需重新导入覆盖。
- **存储**: 每个定价表对应 `data/ratecard/{名称}.json` 一个文件，不涉及目录嵌套。
- **数据格式**: `{ items: [{ name, quantity, unit, unit_price }], remarks }`。

#### 定价表 API

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/api/ratecards` | 获取所有定价表名称列表 |
| `POST` | `/api/ratecards/{name}` | 创建新定价表 |
| `GET` | `/api/ratecards/{name}` | 获取定价表的表格数据 |
| `POST` | `/api/ratecards/{name}/import` | 上传 Excel/CSV 文件并自动解析覆盖数据 |

导入功能使用 `openpyxl` 解析 `.xlsx/.xls` 文件，支持自动表头识别；也支持 `.csv`（UTF-8）。

## 数据流向

```
[项目模块]
  上传图片 → OCR 异步识别 → 保存 ocr-{n}.json → 前端 DataTable 展示 ↔ 编辑保存

[定价表模块]
  新建定价表 → 创建 data/ratecard/{name}.json（空）
  导入 Excel/CSV → 解析写入同一 JSON → 前端 DataTable 预览（如需修改需重新导入）
```

## 运行与开发指南
1.  **环境管理**:
    - **后端**: 本项目使用 `uv` 进行依赖管理与环境同步，请确保已安装 `uv`。
    - **前端**: 本项目使用 `bun` 进行依赖管理与脚本运行，请确保已安装 `bun`。
2.  **环境准备**:
    - **后端**: 确保已安装 Python 环境及所需的依赖库（请查阅 `pyproject.toml`）。
    - **前端**: 确保已进入 `frontend/` 目录，执行 `bun install` 安装依赖。
3.  **启动**:
    - 开发环境下，可分别启动后端服务和前端开发服务器。
    - 使用提供的 `run.bat` 脚本可以快速启动整个系统。

## 文档维护规范
为了确保文档始终反映系统实际状态，遵循以下原则：
- 在进行任何涉及项目结构、API 接口、核心逻辑或数据流向的变更后，必须同步检查并更新本 `AGENTS.md` 文件。
- **编辑警示**: 在使用 `edit` 工具修改代码时，请确保 `oldText` 块精准且最小化，严禁通过包含大量无关的前后文来增加匹配难度，防止误伤不相关的代码区域。
- 严禁在文档中保留过时的路径、名称或逻辑描述。
