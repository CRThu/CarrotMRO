# CarrotMRO

一个基于 AI 的 MRO（维护、维修、运营）综合管理系统，支持 OCR 识别报价单、智能匹配定价表和报价单管理。

![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.14+-blue.svg)
![React](https://img.shields.io/badge/react-18+-61dafb.svg)

## ✨ 功能特性

- **OCR 报价单识别** - 上传报价单图片，AI 自动识别并提取物料信息
- **智能匹配** - 模糊搜索定价表，自动匹配物料名称和单价
- **报价单管理** - 创建、编辑、导入导出报价单
- **定价表管理** - 支持 Excel/CSV 导入，灵活的列配置系统
- **项目管理** - 按项目组织报价单和定价表关联

## 🛠️ 技术栈

### 后端
- **Web 框架**: FastAPI
- **OCR 识别**: Google Gemini API
- **数据处理**: Polars
- **模糊匹配**: RapidFuzz
- **Excel 解析**: OpenPyXL

### 前端
- **框架**: React 18+
- **构建工具**: Vite
- **样式**: Tailwind CSS
- **UI 组件**: shadcn/ui
- **包管理**: Bun

## 🚀 快速开始

### 环境要求

- Python 3.14+
- Node.js 18+ / Bun
- uv (Python 包管理器)

### 安装步骤

1. **克隆仓库**
   ```bash
   git clone https://github.com/your-username/CarrotMRO.git
   cd CarrotMRO
   ```

2. **配置环境变量**
   ```bash
   cp .env.example .env
   ```
   编辑 `.env` 文件，必须配置：
   | 变量 | 说明 |
   |------|------|
   | `GEMINI_API_KEY` | Google Gemini API Key，从 [AI Studio](https://aistudio.google.com/app/apikey) 获取 |
   | `GEMINI_MODEL` | 模型名称，默认 `gemini-3.1-flash-lite` |

3. **安装后端依赖**
   ```bash
   cd backend
   uv sync
   ```

4. **安装前端依赖**
   ```bash
   cd frontend
   bun install
   ```

### 运行项目

**开发模式**（推荐）
```bash
# Windows
dev.bat

# 或手动启动
# 终端 1: 后端
cd backend && uvicorn main:app --reload --port 8000

# 终端 2: 前端
cd frontend && bun dev
```

**生产模式**
```bash
# 构建前端
build.bat

# 启动服务
run.bat
```

## 📖 使用说明

### 1. 项目管理

- 在侧边栏创建新项目
- 上传报价单图片进行 OCR 识别
- 查看和编辑识别结果

### 2. 报价单管理

- 创建新报价单
- 从 OCR 结果导入物料信息
- 点击匹配图标自动填充单价
- 支持手动编辑和自定义

### 3. 定价表管理

- 创建定价表并导入 Excel/CSV 文件
- 支持 `!`（必须匹配）和 `?`（可选匹配）列标记
- 自动识别物料名称列

## 📡 API 文档

启动后端服务后访问: `http://localhost:8000/docs`

### 核心接口

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/api/projects` | 获取所有项目 |
| `POST` | `/api/projects/{name}` | 创建项目 |
| `POST` | `/api/projects/{name}/ocr` | 上传图片 OCR 识别 |
| `GET` | `/api/projects/{name}/quotations` | 获取报价单列表 |
| `POST` | `/api/projects/{name}/quotations` | 创建报价单 |
| `GET` | `/api/ratecards` | 获取定价表列表 |
| `POST` | `/api/ratecards/{name}/import` | 导入定价表 |
| `POST` | `/api/match` | 模糊匹配物料 |

## 📁 项目结构

```
CarrotMRO/
├── backend/              # FastAPI 后端
│   ├── main.py          # API 路由
│   ├── config.py        # 配置管理
│   ├── photo_ocr.py     # OCR 处理
│   ├── match.py         # 模糊匹配
│   └── ratecard_parser.py  # 定价表解析
├── frontend/            # React 前端
│   ├── src/
│   │   ├── App.tsx      # 应用入口
│   │   ├── api.ts       # API 封装
│   │   └── components/  # 组件
│   └── vite.config.js   # Vite 配置
├── data/                # 数据存储
│   ├── projects/        # 项目数据
│   └── ratecard/        # 定价表
├── dev.bat              # 开发启动
├── run.bat              # 生产启动
└── build.bat            # 构建脚本
```

## 📄 许可证

本项目采用 [Apache License 2.0](LICENSE) 许可证。

```
Copyright 2024 CarrotMRO Contributors

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```

---

⭐ 如果这个项目对你有帮助，请给个 Star 支持一下！
