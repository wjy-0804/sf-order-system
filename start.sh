#!/bin/bash
# 顺丰快递模版录入系统 - 一键启动脚本

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PORT=5001

echo "🚀 启动顺丰快递模版录入系统..."

# 检查端口是否已被占用
if lsof -Pi :$PORT -sTCP:LISTEN -t >/dev/null 2>&1; then
    echo "⚠️  端口 $PORT 已被占用，正在停止旧进程..."
    lsof -ti:$PORT | xargs kill -9 2>/dev/null
    sleep 1
fi

# 启动服务
cd "$SCRIPT_DIR"
python3 app.py &
SERVER_PID=$!

sleep 2

# 打开浏览器
echo "✅ 服务启动成功！正在打开浏览器..."
open "http://localhost:$PORT"

echo ""
echo "📋 系统信息："
echo "   地址：http://localhost:$PORT"
echo "   进程：$SERVER_PID"
echo ""
echo "按 Ctrl+C 停止服务"

# 等待中断信号
trap "echo ''; echo '🛑 停止服务...'; kill $SERVER_PID 2>/dev/null; exit 0" INT
wait $SERVER_PID
