#!/bin/bash
# 一键部署脚本 - 在云服务器上执行
set -e

APP_DIR="/opt/park-platform"
REPO_URL="https://github.com/替换为你的仓库.git"

echo "===== 青川智慧工业园 自动部署 ====="
echo "时间: $(date '+%Y-%m-%d %H:%M:%S')"

if [ -d "$APP_DIR" ]; then
    echo ">>> 拉取最新代码..."
    cd "$APP_DIR"
    git pull origin main
else
    echo ">>> 首次部署，克隆仓库..."
    git clone "$REPO_URL" "$APP_DIR"
    cd "$APP_DIR"
fi

echo ">>> 安装依赖..."
npm install --production

echo ">>> 创建数据目录..."
mkdir -p .data

echo ">>> 重启服务..."
if command -v pm2 &>/dev/null; then
    pm2 restart park-platform || pm2 start ecosystem.config.js
    pm2 save
else
    echo "!!! PM2 未安装，请先执行 setup.sh"
    exit 1
fi

echo "✅ 部署完成"
pm2 status
