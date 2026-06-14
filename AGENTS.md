# CarrotMRO 项目说明

CarrotMRO 是一个 MRO（维护、维修、运营）综合管理系统，包含多个业务模块。当前系统包含**项目管理**（含 OCR 图片识别报价单）和**协议定价表管理**两大核心模块。

## 系统架构

本系统采用前后端分离架构：
- **后端 (Backend)**: 基于 FastAPI 构建，提供核心 API 接口，负责 OCR 处理、项目管理、定价表管理、数据存储及 Excel 生成。
- **前端 (Frontend)**: 基于 React 构建，提供用户界面。采用树形侧边栏导航 + 右侧工作区布局，包含项目工作台和定价表编辑器两个独立视图。

## 目录结构

```text
.
├── backend/            # 后端 FastAPI 应用
│   ├── main.py         # 核心 API 路由与服务入口
│   ├── config.py       # 配置管理（环境变量、目录路径）
│   ├── photo_ocr.py    # OCR 识别逻辑封装
│   ├── match.py        # 项目名称匹配逻辑
│   ├── file_utils.py   # 文件读写辅助工具
│   ├── ratecard_parser.py # 定价表 Excel/CSV 解析器
│   ├── search.py       # 模糊搜索工具（独立运行）
│   ├── pyproject.toml  # Python 依赖配置
│   └── uv.lock         # 依赖锁文件
├── frontend/           # 前端 React + Vite 应用
│   ├── src/
│   │   ├── App.tsx             # 应用入口，state 管理与布局路由
│   │   ├── api.ts              # API 接口封装
│   │   ├── types.ts            # TypeScript 类型定义
│   │   └── components/
│   │       ├── Sidebar.tsx         # 侧边栏导航组件
│   │       ├── SidebarTree.tsx     # 树形导航基础组件（SidebarSection、TreeItem）
│   │       ├── ProjectWorkspace.tsx # 项目工作区
│   │       ├── RateCardWorkspace.tsx # 定价表工作区
│   │       ├── DataTable.tsx       # 通用数据表格组件
│   │       └── ui/                 # shadcn/ui 基础组件
│   └── vite.config.js  # Vite 配置（含 /api 代理）
├── data/               # 数据存储目录（自动定位，可通过 .env 覆盖）
│   ├── projects/       # 项目数据及 OCR 结果存储
│   └── ratecard/       # 协议定价表 JSON 文件（*.json）
├── .env.example        # 环境变量模板
├── dev.bat             # 开发环境启动（FastAPI + Vite 并行）
├── run.bat             # 生产环境启动（仅 FastAPI）
└── build.bat           # 前端构建 → backend/static/
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
  | `GET` | `/api/tasks/{task_id}` | 轮询 OCR 任务状态（完成时返回 `columns` 和 `result`） |
  | `GET` | `/api/projects/{name}/ocr-files` | 获取项目下的 OCR 结果文件列表 |
  | `GET` | `/api/projects/{name}/ocr-files/{file}` | 读取 OCR 结果数据 |
  | `PUT` | `/api/projects/{name}/ocr-files/{file}` | 保存 OCR 表格编辑结果 |
  | `DELETE` | `/api/projects/{name}/ocr-files/{file}` | 删除指定的 OCR 结果文件 |
- **说明**: 创建后会在 `data/projects/{project_name}/` 下生成 `project.json`，记录项目名称、创建时间和关联的定价表名称。OCR 识别成功后，结果存储为 `ocr-1.json`、`ocr-2.json` ... 每个文件包含识别出的物料清单表格数据（`items`）。若 OCR 识别失败（如 API 限流），**不会生成结果文件**，任务状态返回 `{"status": "error", "message": "..."}`。OCR 接口统一返回 `columns` 定义（后端 `OCR_COLUMNS` 单一事实来源）。

### 2. 协议定价表管理
- **功能**: 创建和管理协议定价表，支持从 Excel/CSV 文件导入数据。导入后可在页面上预览，如需修改需重新导入覆盖。
- **存储**: 每个定价表对应 `data/ratecard/{名称}.json` 一个文件，不涉及目录嵌套。

#### 导入格式规范

Excel/CSV 的表头行使用前缀标记列的匹配属性：

| 前缀 | 含义 | 说明 |
|------|------|------|
| `!` | 必须匹配项 | 后续报价单比对时，此列必须匹配 |
| `?` | 可选匹配项 | 后续报价单比对时，此列可选匹配，允许差异 |
| 无前缀 | 忽略列 | 不导入，不参与比对 |

**示例表头**:
```
序号  !项目名称  !单位  数量  !综合单价  ?税率
```

**Alias 自动映射**:
- 列名包含"名称" → 自动 alias 为 `name`（必须存在）
- 若无列匹配"名称" → 第一个 `!` 列作为 `name`
- 其他列不做自动映射，保留原始列名

**行跳过逻辑**: 任意一个 `!` 列为空 → 跳过该行（分类行/小计行自动过滤）

**输出数据结构**:
```json
{
  "columns": [
    {"name": "项目名称", "strict": true, "alias": "name"},
    {"name": "单位", "strict": true, "alias": null},
    {"name": "综合单价", "strict": true, "alias": null},
    {"name": "税率", "strict": false, "alias": null}
  ],
  "items": [
    {"项目名称": "xxx", "单位": "只", "综合单价": "88", "税率": "8%"}
  ]
}
```

#### 定价表 API

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/api/ratecards` | 获取所有定价表名称列表 |
| `POST` | `/api/ratecards/{name}` | 创建新定价表 |
| `GET` | `/api/ratecards/{name}` | 获取定价表的表格数据 |
| `POST` | `/api/ratecards/{name}/import` | 上传 Excel/CSV 文件并自动解析覆盖数据 |

导入功能使用 `openpyxl` 解析 `.xlsx/.xls` 文件；也支持 `.csv`（UTF-8）。解析时自动识别 `!`/`?` 前缀标记行作为表头。

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
    - **开发模式**: 双击 `dev.bat`，自动并行启动 FastAPI（端口 8000，热重载）和 Vite（端口 5173，代理 `/api`）。
    - **生产模式**: 双击 `run.bat` 启动 FastAPI（端口 8000），前端需先 `build.bat` 构建。
    - **仅构建前端**: 双击 `build.bat`，将 React 构建产物复制到 `backend/static/`。

## 文档维护规范
为了确保文档始终反映系统实际状态，遵循以下原则：
- 在进行任何涉及项目结构、API 接口、核心逻辑或数据流向的变更后，必须同步检查并更新本 `AGENTS.md` 文件。
- **编辑警示**: 在使用 `edit` 工具修改代码时，请确保 `oldText` 块精准且最小化，严禁通过包含大量无关的前后文来增加匹配难度，防止误伤不相关的代码区域。
- 严禁在文档中保留过时的路径、名称或逻辑描述。

## 用户偏好

### 沟通
- 使用中文交流和回复。

### 技术栈
- **Python 包管理**: uv
- **Python 数据处理**: polars（优先于 pandas）
- **Python Web 框架**: FastAPI
- **前端框架**: React
- **CSS**: Tailwind CSS
- **UI 组件库**: shadcn/ui（base-nova 风格，@base-ui/react 原语）
- **前端包管理/运行**: bun
- **前端构建**: Vite

### 前端架构原则
- **组件抽象化防耦合**: 侧边栏、导航等可复用 UI 区域必须抽离为独立组件，避免 App.tsx 堆积过多职责。App.tsx 只负责 state 管理和顶层布局路由，具体 UI 由子组件承载。
- **扩展性优先**: 组件设计时考虑后续模块扩展，抽象出通用模式（如 SidebarSection、TreeItem），而非一次性硬编码。

### 开发原则
- **单一事实来源 (Single Source of Truth)**: 同一数据的定义必须只存在于一个地方。例如列定义、字段映射等配置，后端定义一次，前端通过 API 获取使用，禁止前端再维护一份副本。