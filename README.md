# CarrotMRO

一个基于 AI 的 MRO（维护、维修、运营）综合管理系统，包含项目管理（含 OCR 图片识别报价单、报价单管理）和协议定价表管理两大核心模块。

![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)
![TypeScript](https://img.shields.io/badge/TypeScript-5.7+-blue.svg)
![React](https://img.shields.io/badge/React-19-61dafb.svg)
![Node.js](https://img.shields.io/badge/Node.js-22+-green.svg)
![Bun](https://img.shields.io/badge/Bun-1.1+-black.svg)

## ✨ 功能特性

- **项目管理与配置** - 关联协议定价表、导出 Excel 模板、自定义动态列映射。
- **单一主数据源公式引擎与物理只读锁定** - 采用**单一主数据源原则**，以 `不含税单价` 为唯一核心源全自动推导 `含税单价`、`不含税总价` 与 `含税总价`；推导列在前端 UI 表格全量应用原生 `readOnly` 物理锁定，杜绝数据冲突。
- **原生 Excel 动态公式导出** - 导出 `.xlsx` 报价单时，自动生成 Excel 原生乘法公式（如 `=D5*E5`）与底部合计求和公式（如 `=SUM(F5:F10)`），用户在 Excel / WPS 中修改数量单价时由 Excel 计算引擎自动联动刷新。
- **OCR 报价单识别** - 上传报价单图片，内置真实 AI 打字流控制台推送、协议层 `response_format` 强制 JSON Mode 约束与 3 次退避自动重试机制。
- **协议定价表管理** - 支持 Excel/CSV 导入，按 100% 精确列名自动预选，支持可视化手动表头对齐映射到 10 项标准预制列。
- **报价单智能匹配与自动保存** - 模糊匹配定价表填充单价，按模板一键导出 Excel，支持编辑过程防抖自动保存与状态监控。
- **系统设置与多模型 API 接入** - 支持 Google Gemini、DeepSeek、Xiaomi MiMo 及 Custom 模式，兼容 OpenAI 标准 API 协议，支持多 Key 独立配置与即选即存自动持久化。
- **极简命令行 / CLI 随时启动** - 支持直接 `npx carrotmro` 全局启动，数据天然存放在命令执行的同级 `data/` 目录中。


## 🛠️ 全栈技术栈

- **全栈语言**: TypeScript / Node.js
- **后端服务**: Express
- **前端框架**: React 19
- **前端构建与热重载**: Vite 6
- **CSS 样式**: Tailwind CSS v4
- **UI 组件库**: shadcn/ui（base-nova 风格）
- **包管理 & 运行驱动**: Bun / npm

## 🚀 快速开始

### 环境要求

- Node.js 18+ 或 Bun 1.1+

### 启动运行方式

#### 1. 命令行/NPM 全局随处运行
可以在任意目录下直接执行命令，自动在当前目录下读写 `./data` 并打开浏览器：
```bash
npx carrotmro
```

#### 2. 本地开发与调试
```bash
# 1. 安装依赖
bun install   # 或 npm install

# 2. 启动开发服务（后端 API 与前端热重载）
bun run dev

# 3. 运行 Electron 桌面界面开发调试
bun run dev:electron

# 4. 运行生产构建并自动打开浏览器
bun run start
```

### 📦 打包与发布

```bash
# 1. 打包 Windows 独立单文件桌面程序 (.exe)
bun run package:exe
# 构建产物位于 release/CarrotMRO-Portable-2.0.0.exe

# 2. 发布到 NPM
npm publish
```

## 📁 目录结构

```text
CarrotMRO/
├── bin/cli.js          # CLI 命令行入口
├── src/                # React 19 前端源码
├── server/             # Node.js / Express 后端服务
├── data/               # 本地 JSON 数据存储目录
├── legacy/             # 归档的旧版算法与参考代码
├── package.json        # 依赖与脚本配置
└── AGENTS.md           # 项目架构规范
```

## 📄 开源协议

[Apache License 2.0](LICENSE)
