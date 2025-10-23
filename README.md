# 📊 周报生成器

基于 Git 提交记录自动生成工作周报，支持多项目、可视化界面。

## ✨ 功能特点

- 🎯 支持多个 Git 项目同时生成周报
- 🤖 使用 DeepSeek AI 智能解析提交信息
- 🎨 现代化的 Web 可视化界面
- 📁 **系统文件浏览器集成**（支持 macOS/Windows/Linux）
- 📅 智能日期范围（默认本周周一至周五工作日）
- 🖥️ **详细的控制台输出**（显示所有 Git 提交记录）
- 📥 历史周报管理与下载
- 📝 自动分类任务和问题
- 📊 基于 Excel 模板生成周报
- 🎨 **完整的颜色一致性**（绿色表头、红色标题、白色数据区）
- 💾 **智能缓存功能**（自动保存姓名和项目路径，无需重复输入）

## 🚀 快速开始

### 1. 安装依赖

```bash
npm install
```

### 2. 配置 DeepSeek API Key（必须）

**推荐方式：使用 `.env` 文件**

```bash
# 复制模板文件
cp .env.example .env

# 编辑 .env 文件，将 sk-your-api-key-here 替换为你的真实 API Key
# DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxx
```

**备选方式：设置环境变量**

```bash
export DEEPSEEK_API_KEY="sk-your-api-key-here"
```

> 💡 获取 API Key：访问 https://platform.deepseek.com 注册并获取
> 
> ⚠️ **安全提示**：`.env` 文件已被 `.gitignore` 忽略，不会被提交到版本控制，请放心使用

### 3. 启动 Web 服务器

**方式一：使用启动脚本（推荐）**

```bash
./start.sh
```

启动脚本会自动检查环境变量并启动服务器。

**方式二：直接启动**

```bash
npm run server
```

> 💡 程序会自动从 `.env` 文件加载配置

### 4. 打开浏览器

访问：http://localhost:3000

## 📖 使用方法

### 方式一：可视化界面（推荐）

1. 在浏览器中打开 `http://localhost:3000`
2. 输入周报负责人姓名
3. 选择日期范围（留空则默认本周）
4. 添加一个或多个 Git 项目路径
   - **方式1**：点击"📁 浏览"按钮，通过系统文件选择器选择目录
   - **方式2**：直接输入项目的绝对路径
5. 点击"生成周报"按钮
6. 等待处理完成后自动下载

### 方式二：命令行模式

如果需要使用原始的命令行方式：

1. 确保已配置 `.env` 文件（参考上面的配置步骤）
2. 修改 `config.json` 配置文件（或让程序自动生成）
3. 运行：`npm start`

## 📁 项目结构

```
weekly-report-generator/
├── index.js              # 原始命令行版本
├── server.js             # Web 服务器
├── start.sh              # 启动脚本
├── config.json           # 用户配置（自动生成）
├── public/
│   └── index.html        # 可视化界面
├── output/               # 生成的周报存放目录
├── 周报模版.xlsx         # Excel 模板
├── .env.example          # 环境变量模板
├── .env                  # 环境变量配置（需自行创建，不会被提交）
├── .gitignore            # Git 忽略规则
├── package.json          # 项目配置
└── CHANGELOG.md          # 更新日志
```

## 🔧 配置说明

### Excel 模板格式

模板需要包含以下内容：

- **标题行（第1行）**：会自动填充为"姓名 年月日期范围 工作周报"
- **重点任务表格（第4行开始）**：包含序号、重点需求或任务、事项说明等列
- **日常问题表格（第12行开始）**：包含序号、问题分类、具体描述等列

### 项目路径

- 必须是 Git 仓库的绝对路径
- 示例：`/Users/username/projects/my-project`
- 支持同时添加多个项目

**两种添加方式：**

1. **使用文件浏览器**（推荐）
   - 点击"📁 浏览"按钮
   - 在弹出的系统文件选择器中选择目录
   - 支持 macOS、Windows、Linux 系统

2. **手动输入路径**
   - 直接在输入框中粘贴或输入绝对路径

## 🎯 AI 解析说明

系统会自动调用 DeepSeek AI 分析每条 Git 提交记录，并：

- 判断是"任务"还是"问题"
- 自动分类（开发新功能、修复bug、优化性能等）
- 提取简洁的工作描述
- 识别关联的需求/BUG编号

## 📝 生成的周报包含

1. **重点任务跟进**
   - 序号、任务名称、事项说明
   - 启动日期、完成日期、负责人
   - 协同人、完成进度、备注

2. **日常问题处理**
   - 序号、问题分类、具体描述
   - 提出日期、解决方案、解决日期

## 🌟 界面特色

- 📱 响应式设计，支持各种屏幕尺寸
- 🎨 渐变紫色主题，视觉美观
- ✨ 流畅的动画效果
- 📂 项目路径管理（添加/删除）
- 📜 历史周报列表与下载
- ⚡ 实时状态提示

## 🛠️ 技术栈

- **后端**：Node.js + Express
- **前端**：原生 HTML/CSS/JavaScript
- **AI**：DeepSeek API (OpenAI SDK)
- **Excel**：ExcelJS
- **Git**：通过命令行调用

## 🔧 命令说明

- `./start.sh` - 使用启动脚本启动（自动检查环境变量）
- `npm run server` - 直接启动 Web 服务器
- `npm start` - 运行原命令行版本

## ❓ 常见问题

### Q: 如何获取 DeepSeek API Key？

A: 访问 https://platform.deepseek.com 注册并获取 API Key

### Q: 忘记设置 API Key 会怎样？

A: 程序会在启动时显示警告，生成周报时会因无法调用 DeepSeek API 而失败。请确保配置了 `.env` 文件或设置了环境变量。

### Q: .env 文件和环境变量哪个优先级更高？

A: 如果同时存在，环境变量的优先级更高。但推荐使用 `.env` 文件方式，更安全和方便。

### Q: 可以不填日期吗？

A: 可以，留空会默认使用本周的周一到周日

### Q: 支持哪些 Git 平台？

A: 支持所有 Git 仓库（GitHub、GitLab、Gitee等），只要是本地 Git 项目即可

### Q: 生成的文件在哪里？

A: 在项目根目录下的 `output/` 文件夹中

### Q: 文件浏览器在 Linux 上无法使用？

A: Linux 系统需要安装 zenity 工具：
```bash
sudo apt-get install zenity  # Debian/Ubuntu
sudo yum install zenity       # CentOS/RHEL
```

### Q: 点击"浏览"按钮没有反应？

A: 
- 确保服务器正在运行
- 检查浏览器控制台是否有错误信息
- macOS 可能需要授予终端访问文件权限（系统偏好设置 > 安全性与隐私）

## 📄 License

MIT

## 👨‍💻 作者

陈毅

