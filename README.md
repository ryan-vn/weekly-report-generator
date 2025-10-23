# 📊 周报生成器

一个基于 Git 提交记录和 DeepSeek AI 的智能周报生成工具，支持 Web 界面和命令行模式。

## ✨ 主要功能

- 🤖 **智能分析**: 使用 DeepSeek AI 分析 Git 提交记录，自动生成专业周报
- 📊 **项目聚合**: 按项目分组，识别代码模块，将相关提交聚合成高质量任务
- 🌐 **Web 界面**: 直观的 Web 界面，支持实时预览和编辑
- 📧 **邮件发送**: 自动发送周报到指定邮箱
- 📁 **多项目支持**: 支持同时分析多个 Git 项目
- 📅 **灵活时间**: 支持自定义时间范围或使用本周
- 📝 **Excel 输出**: 生成标准格式的 Excel 周报文件
- 🔧 **配置管理**: 支持保存和恢复配置

## 🚀 快速开始

### 1. 安装依赖

```bash
npm install
```

### 2. 配置 API Key

**推荐方式：使用 .env 文件**

1. 复制环境变量模板：
```bash
cp .env.example .env
```

2. 编辑 `.env` 文件，填入你的 DeepSeek API Key：
```bash
# DeepSeek API 配置
DEEPSEEK_API_KEY=sk-your-api-key-here

# 邮件服务配置（可选）
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your-email@example.com
SMTP_PASS=your-email-password
MAIL_FROM_NAME=周报生成器
MAIL_FROM_EMAIL=your-email@example.com
MAIL_TO_DEFAULT=manager@example.com
MAIL_CC_DEFAULT=team@example.com,hr@example.com
```

**获取 API Key**: 访问 [DeepSeek 平台](https://platform.deepseek.com) 获取 API Key

### 3. 启动服务

**方式一：使用启动脚本（推荐）**
```bash
chmod +x start.sh
./start.sh
```

**方式二：直接启动**
```bash
npm start
```

启动后访问：http://localhost:3000

## 📧 邮件配置

### SMTP 服务器配置

在 `.env` 文件中配置邮件服务器：

```bash
# SMTP 服务器设置
SMTP_HOST=smtp.example.com        # 邮件服务器地址
SMTP_PORT=587                     # 端口（587 或 465）
SMTP_SECURE=false                 # 是否使用 SSL（465 端口设为 true）
SMTP_USER=your-email@example.com  # 发送邮箱账号
SMTP_PASS=your-email-password     # 邮箱密码或应用密码

# 邮件发送设置
MAIL_FROM_NAME=周报生成器         # 发件人名称
MAIL_FROM_EMAIL=your-email@example.com  # 发件人邮箱
MAIL_TO_DEFAULT=manager@example.com     # 默认收件人
MAIL_CC_DEFAULT=team@example.com,hr@example.com  # 默认抄送
```

### 常用邮件服务商配置

**Gmail**
```bash
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your-gmail@gmail.com
SMTP_PASS=your-app-password  # 使用应用密码，不是登录密码
```

**QQ 邮箱**
```bash
SMTP_HOST=smtp.qq.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your-qq@qq.com
SMTP_PASS=your-authorization-code  # 使用授权码
```

**163 邮箱**
```bash
SMTP_HOST=smtp.163.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your-email@163.com
SMTP_PASS=your-authorization-code
```

### 邮件发送功能

1. 在 Web 界面中勾选"📧 自动发送邮件"
2. 填写收件人和抄送邮箱
3. 自定义邮件主题和内容（可选）
4. 生成周报时会自动发送邮件

## 🎯 使用方式

### Web 界面模式

1. 访问 http://localhost:3000
2. 填写周报负责人姓名
3. 选择日期范围（留空默认本周）
4. 添加 Git 项目路径
5. 配置邮件发送（可选）
6. 点击"生成周报"
7. 预览和编辑周报内容
8. 下载 Excel 文件

### 命令行模式

```bash
node index.js
```

按提示输入：
- 周报负责人姓名
- 开始日期
- 结束日期  
- Git 项目路径

## 📁 项目结构

```
weekly-report-generator/
├── 📄 index.js              # 命令行版本
├── 🌐 server.js             # Web 服务器
├── 📱 public/index.html     # Web 界面
├── ⚙️ config.json           # 配置文件
├── 📋 周报模版.xlsx         # Excel 模板
├── 📤 output/               # 输出目录
├── 🔧 .env.example          # 环境变量模板
├── 🚀 start.sh              # 启动脚本
└── 📖 README.md             # 说明文档
```

## 🔧 配置说明

### config.json

```json
{
  "userName": "xx",
  "projectPaths": [
    "/Users/vincent/project/weekly-report-generator"
  ],
  "deepseekModel": "deepseek-chat",
  "dateFormat": "yyyy年MM月dd日"
}
```

### 环境变量优先级

1. `.env` 文件中的 `DEEPSEEK_API_KEY`
2. `config.json` 中的 `deepseekApiKey`
3. 系统环境变量 `DEEPSEEK_API_KEY`

## 🤖 AI 分析特性

### 智能模块识别
- 自动识别代码模块（用户模块、订单模块、支付模块等）
- 根据文件路径和提交信息分析功能

### 任务聚合
- 将同一模块的多次提交合并为一个任务
- 生成专业、简洁的工作描述
- 突出工作价值，避免技术细节

### 分析示例
```
原始提交: 15 条
↓ AI 智能分析
生成任务: 3 条
聚合率: 80%
```

## 📊 输出格式

### Excel 周报包含：
- 序号
- 重点需求或任务
- 事项说明（包含关键改动）
- 启动日期
- 预计完成日期
- 负责人
- 协同人或部门
- 完成进度
- 备注

### 邮件内容：
- 专业的 HTML 格式邮件
- 自动生成邮件主题和内容
- 支持自定义邮件内容
- Excel 文件作为附件

## ❓ 常见问题

### Q: 如何获取 DeepSeek API Key？
A: 访问 [DeepSeek 平台](https://platform.deepseek.com)，注册账号并创建 API Key。

### Q: 支持哪些邮件服务商？
A: 支持所有标准 SMTP 服务商，包括 Gmail、QQ 邮箱、163 邮箱、企业邮箱等。

### Q: 如何配置 Gmail？
A: 需要开启两步验证并使用应用密码，不是登录密码。

### Q: 邮件发送失败怎么办？
A: 检查 SMTP 配置、网络连接、邮箱密码或授权码是否正确。

### Q: 可以同时分析多个项目吗？
A: 可以，在项目路径中添加多个 Git 项目路径即可。

### Q: 如何自定义邮件内容？
A: 在 Web 界面的邮件配置中可以自定义邮件主题和内容。

## 🔒 安全提示

- ⚠️ **不要**将 `.env` 文件提交到版本控制系统
- ✅ **使用** `.env.example` 作为配置模板
- 🔐 **保护**你的 API Key 和邮箱密码
- 🛡️ **定期**更换密码和 API Key

## 📝 更新日志

### v2.0.0 - 智能分析升级
- ✨ 新增按项目分组的智能分析模式
- 🤖 AI 自动识别代码模块和功能
- 📊 提交聚合，减少冗余任务
- 📧 新增邮件自动发送功能
- 🎨 优化 Web 界面和用户体验

### v1.0.0 - 基础功能
- 🚀 基础周报生成功能
- 📊 Excel 文件输出
- 🌐 Web 界面支持
- 📁 多项目分析

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

**享受智能周报生成的便利！** 🎉