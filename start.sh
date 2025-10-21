#!/bin/bash

# 周报生成器启动脚本

echo "🚀 启动周报生成器..."
echo ""

# 检查 .env 文件
if [ -f .env ]; then
    echo "✅ 找到 .env 文件，加载环境变量..."
    export $(cat .env | grep -v '^#' | xargs)
else
    echo "⚠️  未找到 .env 文件"
    echo "   你可以："
    echo "   1. 创建 .env 文件: cp .env.example .env"
    echo "   2. 或直接设置环境变量: export DEEPSEEK_API_KEY='your-key'"
    echo ""
fi

# 检查是否设置了 API Key
if [ -z "$DEEPSEEK_API_KEY" ]; then
    echo "❌ 错误: DEEPSEEK_API_KEY 未设置"
    echo "   请先设置 DeepSeek API Key"
    echo ""
    echo "   方法1 - 创建 .env 文件:"
    echo "   cp .env.example .env"
    echo "   # 然后编辑 .env 文件，填入你的 API Key"
    echo ""
    echo "   方法2 - 直接设置环境变量:"
    echo "   export DEEPSEEK_API_KEY='sk-your-api-key-here'"
    echo ""
    exit 1
fi

echo "✅ DeepSeek API Key 已配置"
echo ""
echo "启动 Web 服务器..."
echo ""

# 启动服务器
npm run server

