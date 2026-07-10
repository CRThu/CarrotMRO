# CarrotMRO 项目说明

CarrotMRO 是一个 MRO（维护、维修、运营）综合管理系统，包含多个业务模块。当前系统包含**项目管理**（含 OCR 图片识别报价单、报价单管理）和**协议定价表管理**两大核心模块。

## 系统架构

本系统采用前后端分离架构：
- **后端 (Backend)**: 基于 FastAPI 构建，提供核心 API 接口，负责 OCR 处理、项目管理、定价表管理、数据存储及 Excel 生成。
- **前端 (Frontend)**: 基于 React 构建，提供用户界面。采用树形侧边栏导航 + 右侧工作区布局，包含项目工作台、报价单编辑器和定价表查看器三个独立视图。

## 目录结构

```text
.
├── backend/            # 后端 FastAPI 应用
│   ├── main.py         # 核心 API 路由与服务入口
│   ├── config.py       # 配置管理（环境变量、目录路径）
│   ├── photo_ocr.py    # OCR 识别逻辑封装（litellm + Xiaomi MiMo）
│   ├── preset_columns.py # 全局预制列定义
│   ├── match.py        # 模糊搜索工具（match_names）
│   ├── ratecard_parser.py # 定价表 Excel/CSV 解析器（含 extract_names）
│   ├── quotation_template.py # 报价单 Excel 模板解析与导出
│   ├── pyproject.toml  # Python 依赖配置
│   ├── uv.lock         # 依赖锁文件
│   └── tests/          # 后端单元测试（pytest）
├── frontend/           # 前端 React + Vite 应用
│   ├── src/
│   │   ├── App.tsx             # 应用入口，state 管理与布局路由
│   │   ├── api.ts              # API 接口封装
│   │   ├── types.ts            # TypeScript 类型定义
│   │   ├── test/               # 前端测试配置
│   │   └── components/
│   │       ├── Sidebar.tsx         # 侧边栏导航组件
│   │       ├── SidebarTree.tsx     # 树形导航基础组件（SidebarSection、TreeItem）
│   │       ├── ProjectWorkspace.tsx # 项目工作区（OCR 数据编辑）
│   │       ├── ProjectConfigWorkspace.tsx # 项目配置工作区（关联定价表、列映射）
│   │       ├── RateCardWorkspace.tsx # 定价表工作区
│   │       ├── QuotationWorkspace.tsx # 报价单工作区（含匹配功能）
│   │       ├── MatchPopover.tsx    # 匹配候选选择弹窗（点击图标选择清单名称）
│   │       ├── TaskNotification.tsx # 通用异步任务通知（浮动提示）
│   │       ├── ErrorBoundary.tsx   # 错误边界（显示完整堆栈信息）
│   │       ├── DataTable.tsx       # 通用数据表格组件
│   │       └── ui/                 # shadcn/ui 基础组件
│   ├── vitest.config.js  # Vitest 测试配置
│   └── vite.config.js  # Vite 配置（含 /api 代理）
├── e2e/                # 端到端测试（Playwright）
├── data/               # 数据存储目录（自动定位，可通过 .env 覆盖）
│   ├── projects/       # 项目数据及 OCR 结果存储
│   ├── ratecard/       # 协议定价表 JSON 文件（*.json）
│   └── template/       # 报价单 Excel 模板（*.xlsx）
├── .env.example        # 环境变量模板
├── dev.bat             # 开发环境启动（FastAPI + Vite 并行）
├── run.bat             # 生产环境启动（仅 FastAPI）
└── build.bat           # 前端构建 → backend/static/
```

## 核心功能

### 1. 项目管理
- **功能**: 创建项目、上传报价单图片进行 OCR 识别、查看和编辑识别结果、管理报价单。
- **API**:
  | 方法 | 路径 | 说明 |
  |------|------|------|
  | `GET` | `/api/projects` | 获取所有项目列表 |
  | `POST` | `/api/projects/{name}` | 创建新项目（生成目录 + `project.json`） |
  | `GET` | `/api/projects/{name}` | 获取项目基本信息 |
  | `GET` | `/api/preset-columns` | 获取全局预制列列表 |
  | `PATCH` | `/api/projects/{name}/ratecard` | 关联/解除关联定价表 |
  | `PATCH` | `/api/projects/{name}/template` | 关联/解除关联模板 |
  | `GET` | `/api/projects/{name}/columns` | 获取项目列配置（可用列 + 已选列 + 列映射） |
  | `PATCH` | `/api/projects/{name}/columns` | 设置项目选中的列和列映射 |
  | `POST` | `/api/projects/{name}/ocr` | 上传图片启动 OCR 异步识别 |
  | `GET` | `/api/tasks/{task_id}` | 轮询 OCR 任务状态（完成时返回 `columns` 和 `result`） |
  | `GET` | `/api/projects/{name}/ocr-files` | 获取项目下的 OCR 结果文件列表 |
  | `GET` | `/api/projects/{name}/ocr-files/{file}` | 读取 OCR 结果数据 |
  | `PUT` | `/api/projects/{name}/ocr-files/{file}` | 保存 OCR 表格编辑结果 |
  | `DELETE` | `/api/projects/{name}/ocr-files/{file}` | 删除指定的 OCR 结果文件 |
  | `GET` | `/api/projects/{name}/quotations` | 获取项目下的报价单文件列表 |
  | `POST` | `/api/projects/{name}/quotations` | 创建新报价单（自动命名 quotation-{n}.json） |
  | `GET` | `/api/projects/{name}/quotations/{file}` | 读取报价单数据 |
  | `PUT` | `/api/projects/{name}/quotations/{file}` | 保存报价单（自动更新 last_edit_time） |
  | `DELETE` | `/api/projects/{name}/quotations/{file}` | 删除指定的报价单文件 |
  | `POST` | `/api/projects/{name}/quotations/{file}/export` | 用关联模板导出 Excel |
  | `POST` | `/api/projects/{name}/quotations/import` | 上传 Excel 导入为新报价单 |
- **说明**: 创建后会在 `data/projects/{project_name}/` 下生成 `project.json`，记录项目名称、创建时间、关联的定价表名称、关联的模板名称和选中的列名列表。OCR 识别使用动态列配置：用户关联模板后，从模板可用列中勾选需要识别的列，OCR 只识别这些列。报价单存储为 `quotation-1.json`、`quotation-2.json` ...，每个文件包含 `created_at`、`last_edit_time` 和 `items`。报价单列名由模板定义，用户通过勾选确认。支持按 `_group` 字段分组导出，无 `_group` 则作为单组导出。

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

**Alias 说明**:
- `alias` 字段保留但不再自动设置（已移除自动映射逻辑）
- 列名由用户在项目配置中手动映射到预制列
- 其他列不做自动映射，保留原始列名

**行跳过逻辑**: 任意一个 `!` 列为空 → 跳过该行（分类行/小计行自动过滤）

**输出数据结构**:
```json
{
  "columns": [
    {"name": "项目名称", "strict": true, "alias": null},
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
| `POST` | `/api/match` | 模糊匹配定价表 name 列（支持多关键词） |

导入功能使用 `openpyxl` 解析 `.xlsx/.xls` 文件；也支持 `.csv`（UTF-8）。解析时自动识别 `!`/`?` 前缀标记行作为表头。

**搜索接口参数**:
```json
{
  "ratecard_name": "定价表名称",
  "queries": ["搜索词1", "搜索词2"],
  "limit": 5
}
```

### 3. 模板管理

- **功能**: 上传、下载、删除报价单 Excel 模板。
- **存储**: 每个模板对应 `data/template/{文件名}.xlsx`。
- **列提取**: 模板上传时自动解析 `{item.xxx}` 占位符，提取 `xxx` 作为可用列名。

#### 模板 API

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/api/templates` | 列出所有模板文件 |
| `POST` | `/api/templates` | 上传新模板（.xlsx） |
| `GET` | `/api/templates/{filename}` | 下载模板文件 |
| `DELETE` | `/api/templates/{filename}` | 删除模板 |
| `GET` | `/api/templates/{filename}/info` | 获取模板解析信息（可用列名列表） |

### 4. 报价单模板导出

- **功能**: 基于 Excel 模板导出报价单，自动填充数据、复制样式、更新公式行号。
- **模块**: `backend/quotation_template.py`

#### 模板占位符

占位符采用命名空间前缀区分循环层级：

| 占位符 | 说明 | 输出次数 |
|--------|------|----------|
| `{group.name}` | 组名 | 每组一次 |
| `{group.num}` | 组内序号（从 1 开始） | 每组一次 |
| `{item.xxx}` | 数据字段（xxx 为模板自定义名称） | 每个 item 一次 |
| `{row}` | 当前行号（公式引用） | 导出时替换 |
| `{row-N}` | 当前行号减 N | 导出时替换 |
| `{row+N}` | 当前行号加 N | 导出时替换 |

**示例模板**:
```
行5: {group.name}                           ← 组名行（每组复制一次）
行6: {group.num} | {item.name} | {item.unit} | {item.quantity} | {item.unit_price} | =D{row}*E{row}  ← 数据行（每个 item 复制一次）
行8: 合计 | =SUM(F5:F{row-1})               ← 页脚（公式行号自动更新）
行10: 总价 | =F{row-2}*(1+F{row-1})
```

#### 导出数据格式

```json
[
  {
    "name": "组名",
    "items": [
      {"name": "项目名称", "unit": "单位", "quantity": "数量", "unit_price": "单价"}
    ]
  }
]
```

items 中的 key 为纯字段名（不带 `item.` 前缀），导出时自动映射到模板占位符。

#### 导出流程

1. 解析模板，识别 group_row 和 data_row
2. 根据数据量计算需要插入的行数，在 data_row 后插入空行
3. 遍历每个组：复制组名行样式 + 填充占位符，再逐行复制数据行样式 + 填充
4. 扫描页脚区域，更新包含 `{row}` 的公式行号
5. 如果组名行有合并单元格，为每个组重新创建合并区域

#### 函数接口

| 函数 | 说明 |
|------|------|
| `load_template(content: bytes)` | 从 Excel 文件内容加载并解析模板 |
| `load_template_from_file(path)` | 从文件路径加载模板 |
| `get_template_info(template)` | 获取模板摘要信息（用于 API 返回） |
| `get_template_columns(template)` | 提取模板中所有可用的数据列名（从 {item.xxx} 中提取 xxx） |
| `export_quotation(template_content, groups)` | 将数据填充到模板，生成 Excel 文件 |
| `import_quotation(template_content, excel_content)` | 从 Excel 文件中提取报价单数据，按模板结构分组返回 |

## 数据流向

```
[项目模块]
  关联模板 → 解析模板可用列 → 用户勾选需要的列 → 存储到 project.json
  上传图片 → OCR 动态识别（使用用户勾选的列名） → 保存 ocr-{n}.json → 前端 DataTable 展示 ↔ 编辑保存

[报价单模块]
  点击"+"创建报价单 → 生成 quotation-{n}.json（含 created_at、last_edit_time）→ 前端 QuotationWorkspace 展示
  从 OCR 导入 → 使用模板定义的列名 → 每行显示匹配状态图标（🔵pending/🟢matched/🟠custom）
  点击匹配图标 → 调用 /api/match 获取 top5 候选 → 弹出 MatchPopover → 点击选择填充单价+清单名称
  无匹配结果 → 保持 pending → 用户手动选择"不匹配—自定义"
  导出 Excel → 读取关联模板 → 按 _group 字段分组 → 填充模板占位符 → 生成 Excel 文件
  导入 Excel → 读取关联模板 → 识别组名行/数据行 → 解析数据 → 创建新报价单

[定价表模块]
  新建定价表 → 创建 data/ratecard/{name}.json（空）
  导入 Excel/CSV → 解析写入同一 JSON → 前端 DataTable 预览（如需修改需重新导入）

[模板模块]
  上传模板 → 存储到 data/template/{filename}.xlsx
  关联模板 → 更新 project.json 的 template_name 字段
  导出时 → 读取关联模板 → 填充数据 → 返回 Excel
  导入时 → 读取关联模板 → 解析 Excel → 返回分组数据
```

## 数据规范

- **统一 key**: 所有 items 的 key 统一使用 `col.name`（中文列名），前后端一致。`col.alias` 仅用于跨表匹配（如报价单名称匹配定价表时的关联）。
- **OCR 数据映射**: litellm + Xiaomi MiMo 模型返回 JSON 数据，后端按 `column_mappings` 中的配置将模板列名映射为预制列名。
- **预制列映射**: 每个项目支持三个独立的映射（ocr、ratecard、quotation），存储在 `project.json` 的 `column_mappings` 字段中。

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
4.  **测试**:
    - **后端单元测试**: `cd backend && uv run pytest -v`
    - **前端单元测试**: `cd frontend && bunx vitest run`
    - **端到端测试**: `cd e2e && npx playwright test`

## 文档维护规范
为了确保文档始终反映系统实际状态，遵循以下原则：
- 在进行任何涉及项目结构、API 接口、核心逻辑或数据流向的变更后，必须同步检查并更新本 `AGENTS.md` 文件。
- **编辑警示**: 在使用 `edit` 工具修改代码时，请确保 `oldText` 块精准且最小化，严禁通过包含大量无关的前后文来增加匹配难度，防止误伤不相关的代码区域。
- 严禁在文档中保留过时的路径、名称或逻辑描述。

## 开发规范
- **禁止读取 data/ 目录**: agent 禁止读取或列出 `data/` 目录下的任何文件，该目录包含用户实际数据，仅由运行时程序访问。
- **禁止读取 .env 文件**: agent 禁止读取 `.env` 文件，该文件包含敏感配置信息（API 密钥、数据库连接等）。

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
- **最小改动原则**: 引入新依赖或重构时，优先选择与现有架构兼容的方案，减少上下游文件的连锁修改。新功能通过扩展配置（如声明式 props）实现，而非修改已有组件的内部逻辑。改动范围应局限在必要文件，避免波及不相关的模块。
- **显式优于隐式**: 数据流和组件行为必须可追踪。优先使用显式 props 传递和配置声明，避免隐式约定（如依赖内部状态推断、中间层不可见转换）。组件职责单一，不隐藏副作用，不依赖调用顺序。