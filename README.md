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

## 📥 安装与使用方式

### 环境要求
- Node.js 18+ 或 Bun 1.1+ (推荐 Node.js v20+)

---

### 1. NPX 零安装一键随处运行 (推荐)
无需提前全局安装任何软件包，在终端任意目录下直接执行以下命令，系统会自动创建/读取同级 `./data` 数据目录并自动在默认浏览器中打开系统：
```bash
npx carrotmro
```
> **自定义端口**：如需指定端口，可使用 `--port` 或 `-p` 参数：
> ```bash
> npx carrotmro --port 15000
> ```

---

### 2. 包管理器全局安装 (Global Installation)
通过 NPM、PNPM 或 Bun 将 `carrotmro` 注册为全局系统命令，即可在终端随时唤起：

```bash
# 使用 npm 全局安装
npm install -g carrotmro

# 使用 pnpm 全局安装
pnpm add -g carrotmro

# 使用 bun 全局安装
bun add -g carrotmro

# 安装完成后在任意目录下执行：
carrotmro
```

---

### 3. 源码克隆与本地二次开发 (Local Development)
适用于开发人员扩展新功能、修改算法或自定义 UI 样式：

```bash
# 1. 克隆 GitHub 源码仓库
git clone https://github.com/CRThu/CarrotMRO.git
cd CarrotMRO

# 2. 安装项目依赖
bun install   # 或 npm install

# 3. 启动开发调试服务 (前端 Vite 热重载 + 后端 API)
bun run dev   # 或 npm run dev

# 4. 启动 Electron 桌面客户端开发模式
bun run dev:electron

# 5. 执行全栈生产编译打包
bun run build # 或 npm run build
```

### 📦 NPM 发布命令

```bash
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
