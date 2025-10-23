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
echo "📦 检查并安装依赖..."
if [ ! -d "node_modules" ] || [ "package.json" -nt "node_modules" ]; then
    echo "   正在安装依赖包..."
    npm install
    if [ $? -eq 0 ]; then
        echo "✅ 依赖安装完成"
    else
        echo "❌ 依赖安装失败，请检查网络连接和 npm 配置"
        exit 1
    fi
else
    echo "✅ 依赖已是最新版本"
fi

echo ""
echo "📝 提示: 如果启动后提示缺少 API Key，请创建 .env 文件或设置环境变量"
echo ""
echo "启动 Web 服务器..."
echo ""

# 启动服务器（dotenv 会自动加载 .env 文件）
npm run server &

# 等待服务器启动
echo "⏳ 等待服务器启动..."
sleep 3

# 自动打开浏览器
echo "🌐 正在打开浏览器..."
if command -v open >/dev/null 2>&1; then
    # macOS
    open http://localhost:3000
elif command -v xdg-open >/dev/null 2>&1; then
    # Linux
    xdg-open http://localhost:3000
elif command -v start >/dev/null 2>&1; then
    # Windows (Git Bash)
    start http://localhost:3000
else
    echo "⚠️  无法自动打开浏览器，请手动访问: http://localhost:3000"
fi

echo ""
echo "✅ 服务器已启动，浏览器已打开"
echo "   访问地址: http://localhost:3000"
echo "   按 Ctrl+C 停止服务器"
echo ""

# 等待用户中断
wait

