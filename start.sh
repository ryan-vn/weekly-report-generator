#!/bin/bash

# 周报生成器启动脚本

echo "🚀 启动周报生成器..."
echo ""

# 检查 .env 文件是否存在
if [ -f .env ]; then
    echo "✅ 找到 .env 配置文件"
else
    echo "⚠️  未找到 .env 文件"
    echo ""
    echo "   推荐创建 .env 文件来配置 API Key:"
    echo "   1. 复制模板: cp .env.example .env"
    echo "   2. 编辑 .env 文件，填入你的 DeepSeek API Key"
    echo ""
    echo "   或者你也可以通过环境变量设置:"
    echo "   export DEEPSEEK_API_KEY='sk-your-api-key-here'"
    echo ""
    echo "   按 Ctrl+C 取消启动，或直接回车继续（可能会因缺少 API Key 而失败）"
    read -t 5
fi

echo ""
echo "📝 提示: 如果启动后提示缺少 API Key，请创建 .env 文件或设置环境变量"
echo ""
echo "启动 Web 服务器..."
echo ""

# 启动服务器（dotenv 会自动加载 .env 文件）
npm run server

