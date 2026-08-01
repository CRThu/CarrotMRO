# CarrotMRO 项目说明

CarrotMRO 是一个 MRO（维护、维修、运营）综合管理系统，包含多个业务模块。当前系统包含**项目管理**（含项目设置、报价单管理、图片 OCR 智能提取）和**协议定价表管理**两大核心模块。

## 系统架构

本系统采用全栈 TypeScript 架构，支持单文件 `.exe` 发布与桌面/网页双模式运行：
- **后端 (Server)**: 基于 Node.js / Express 构建，运行于桌面主进程，提供 RESTful API 接口，数据管理采用本地纯 JSON 文件存储 (`data/`)。
- **前端 (Frontend)**: 基于 React 19 + Vite 6 + Tailwind CSS v4 构建，采用树形侧边栏导航 + 右侧工作区布局。
- **运行/桌面壳**: 支持 **Bun Compile 超高速单文件打包** (`bun run package:exe`) 或 **Electron 原生 GUI 桌面应用模式** (`bun run dev:electron`)，端口 3000 天然开放供系统浏览器随时访问。

## 目录结构

```text
.
├── src/                # 前端 React 19 应用
│   ├── App.tsx         # 应用入口，State 管理与工作区路由
│   ├── api.ts          # RESTful API 请求工具封装
│   ├── types.ts        # TypeScript 全局数据模型与 10 项标准预制列定义
│   └── components/     # UI 视图组件
│       ├── Sidebar.tsx         # 侧边栏导航组件 (项目 settings/报价单/定价表)
│       ├── SidebarTree.tsx     # 树形导航基础组件
│       ├── ProjectConfigWorkspace.tsx # 项目设置工作区（绑定定价单/模板、配置 ocr_columns 与 quotation_columns）
│       ├── RateCardWorkspace.tsx # 定价表工作区（含 Excel/CSV 表头映射对齐导入弹窗）
│       ├── QuotationWorkspace.tsx # 报价单工作区（含图片 OCR 识别导入、公式自动联动计算、定价单匹配）
│       ├── SettingsWorkspace.tsx  # 系统设置工作区（多模型接入与独立 Key 配置）
│       ├── MatchPopover.tsx    # 匹配候选选择弹窗
│       └── TaskNotification.tsx # 通用异步任务通知
├── server/             # 后端 Node.js API 服务
│   └── index.ts        # Express RESTful API 路由与 Excel/CSV 解析对齐
├── electron/           # Electron 桌面应用壳
│   └── main.ts         # 原生 GUI 桌面窗口主进程
├── data/               # 本地 JSON 数据与 Excel 模板存储
│   ├── projects/       # 项目数据目录：{项目名称}/settings.json 与 quotation-*.json
│   ├── ratecard/       # 协议定价表 JSON 文件
│   ├── template/       # 报价单 Excel 模板（*.xlsx）
│   └── settings.json   # 多服务商系统设置与 API Key 存储 JSON
└── package.json        # 项目依赖与 Bun 命令集中配置文件
```

## 全局 10 项标准预制列 (Preset Columns)

系统定义了 10 项标准的全局预制列（规格型号与品牌等属性统一包含在`说明`中）：
```typescript
export const PRESET_COLUMNS = [
  '项目组',
  '项目名称',
  '单位',
  '数量',
  '不含税单价',
  '不含税总价',
  '税率',
  '含税单价',
  '含税总价',
  '说明'
] as const;
```

---

## 核心功能

### 1. 项目管理与配置
- **存储结构**: 每个项目存储为 `data/projects/{项目名称}/` 目录，内部包含 `settings.json`：
  ```json
  {
    "name": "某维保项目",
    "created_at": "2026-07-28T12:00:00.000Z",
    "ratecard_name": "2026协议定价表.json",
    "template_name": "标准模板.xlsx",
    "ocr_columns": ["项目名称", "单位", "数量", "不含税单价", "说明"],
    "quotation_columns": ["项目组", "项目名称", "单位", "数量", "不含税单价", "不含税总价", "税率", "含税单价", "含税总价", "说明"],
    "match_validation_rules": {
      "strict_name_match": true,
      "check_columns": ["项目名称", "单位"],
      "fill_columns": ["单位", "不含税单价", "含税单价", "税率", "说明"]
    }
  }
  ```
- **定价单匹配带入与输出前一键校验 (解耦流程)**:
  - **匹配带入 (`fill_columns`)**: 在报价单中点击物料匹配时，仅将定价表中 `fill_columns` 中勾选的列覆盖带入报价单，并自动联动计算公式。
  - **成品输出前一键校验 (`check_columns`)**: 在导出 Excel 或输出成品前，用户可点击顶部【校验】按钮，全表对比已匹配条目与原定价表数据。若发现 `check_columns` 中勾选的严格校验列被误修改，将在“识别提示与复核备注栏”集中提示警告。
- **报价单创建**: 支持在项目下创建多个报价单 (`quotation-1.json`, `quotation-2.json` ...)。
- **图片 OCR 识别导入**: 沉淀为报价单内部的核心导入功能。在报价单编辑页面直接上传图片，调用大模型按照项目配置的 `ocr_columns` 提取数据，并自动填充到当前报价单。后端支持 `stream: true` 真实打字流控制台推送、协议层 `response_format: { type: "json_object" }` 强制 JSON Mode 约束与 3 次指数退避自动重试机制。
- **自动保存机制**:
  - 系统/服务商配置：选择服务商卡片即选即存，参数编辑防抖保存。
  - 报价单编辑：OCR 生成完成自动立即持久化写盘，单元格/行列变动 1.2 秒防抖自动保存，顶部提供实时保存状态栏（`✓ 已自动保存` / `⏱️ 正在保存...`）。
- **公式自动联动**: 报价单编辑时自动执行公式联动：
  - `不含税总价` = `数量` × `不含税单价`
  - `含税单价` = `不含税单价` × (1 + `税率`)
  - `含税总价` = `数量` × `含税单价`

- **项目相关 API**:
  | 方法 | 路径 | 说明 |
  |------|------|------|
  | `GET` | `/api/projects` | 获取所有项目名称列表 |
  | `POST` | `/api/projects/{name}` | 创建新项目目录并初始化 `settings.json` |
  | `GET` | `/api/projects/{name}` | 读取项目的 `settings.json` 配置 |
  | `PATCH` | `/api/projects/{name}/settings` | 更新项目的 `settings.json` 配置 |
  | `POST` | `/api/projects/{name}/ocr` | 上传图片按项目的 `ocr_columns` 提取数据 (支持真实打字流) |
  | `GET` | `/api/projects/{name}/quotations` | 获取项目下的报价单列表 |
  | `POST` | `/api/projects/{name}/quotations` | 创建新报价单（quotation-{n}.json） |
  | `GET` | `/api/projects/{name}/quotations/{file}` | 读取指定报价单数据 |
  | `PUT` | `/api/projects/{name}/quotations/{file}` | 保存报价单数据 |
  | `DELETE` | `/api/projects/{name}/quotations/{file}` | 删除报价单 |

### 2. 协议定价表管理
- **功能**: 管理协议定价表。上传 Excel/CSV 时先解析表头提供映射弹窗，系统仅对 100% 精确同名的列自动预先勾选，非同名的原始列保持留空由用户手动下拉指定，确认后将数据清洗归一化并持久化到 JSON。
- **存储**: 保存为干净的 `data/ratecard/{名称}.json`：
  ```json
  {
    "columns": ["项目名称", "单位", "不含税单价", "说明"],
    "items": [
      { "项目名称": "铜芯电缆", "单位": "米", "不含税单价": "45.00", "说明": "国标" }
    ]
  }
  ```

- **定价表 API**:
  | 方法 | 路径 | 说明 |
  |------|------|------|
  | `GET` | `/api/ratecards` | 获取所有定价表名称列表 |
  | `POST` | `/api/ratecards/{name}` | 创建空定价表 |
  | `GET` | `/api/ratecards/{name}` | 读取定价表数据 |
  | `POST` | `/api/ratecards/{name}/import-preview` | 解析上传的 Excel/CSV 文件原始表头与示例数据 |
  | `POST` | `/api/ratecards/{name}/import` | 提交表头映射配置，将数据归一化后保存到 JSON |
  | `POST` | `/api/match` | 按`项目名称`在定价表中检索匹配单价 |

---

## 运行与测试指南
1. **启动开发**:
   - `bun run dev` (同时启动 Vite 前端与 Node/Express 服务)
2. **测试运行**:
   - `bun run test` (全量运行 56 项自动化单元测试与 API 契约测试)

## 文档维护规范
- 当修改核心 API、项目配置模型或标准列规范后，需同步更新本 `AGENTS.md`。
- Agent 禁止读取 `data/` 和 `.env` 敏感文件。