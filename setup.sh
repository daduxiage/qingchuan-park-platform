#!/bin/bash
# 服务器首次初始化脚本
set -e

echo "===== 服务器环境初始化 ====="

# 1. 安装 Node.js 22
if ! command -v node &>/dev/null; then
    echo ">>> 安装 Node.js..."
    curl -fsSL https://deb.nodesource.com/setup_22.x | bash -
    apt-get install -y nodejs
fi

# 2. 安装 PM2
if ! command -v pm2 &>/dev/null; then
    echo ">>> 安装 PM2..."
    npm install -g pm2
fi

# 3. 安装 git
if ! command -v git &>/dev/null; then
    echo ">>> 安装 Git..."
    apt-get install -y git
fi

# 4. 创建应用目录
mkdir -p /opt/park-platform

# 5. PM2 开机自启
pm2 startup systemd -u root --hp /root
pm2 save

echo "✅ 初始化完成"
echo "node: $(node -v)"
echo "npm:  $(npm -v)"
echo "pm2:  $(pm2 -v)"
echo ""
echo "下一步："
echo "  1. git clone <你的仓库> /opt/park-platform"
echo "  2. cd /opt/park-platform && npm install"
echo "  3. pm2 start ecosystem.config.js"
