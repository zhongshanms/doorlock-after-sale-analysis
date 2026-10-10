#!/bin/bash
# ============================================================
#  亚马逊门锁售后分析系统 —— 服务器端初始化（宝塔 / 通用 Linux）
#  作用: 创建站点目录 + Git 裸仓库 + post-receive 钩子(推送后自动发布)
#  用法: bash 服务器初始化.sh <站点目录名> [裸仓库名]
#  示例: bash 服务器初始化.sh doorlock.example.com
#        生成的站点目录: /www/wwwroot/doorlock.example.com
#        生成的裸仓库:   /www/repo/doorlock.example.com.git
# ============================================================
set -e

SITE_NAME="${1:?用法: bash 服务器初始化.sh <站点目录名> [裸仓库名]}"
REPO_NAME="${2:-$SITE_NAME}"
SITE_DIR="/www/wwwroot/$SITE_NAME"
REPO_DIR="/www/repo/$REPO_NAME.git"

echo "站点目录: $SITE_DIR"
echo "裸仓库:   $REPO_DIR"
echo

mkdir -p "$SITE_DIR" "$(dirname "$REPO_DIR")"

if [ -d "$REPO_DIR" ]; then
  echo "[跳过] 裸仓库已存在"
else
  git init --bare "$REPO_DIR" >/dev/null
  echo "[OK] 已创建裸仓库"
fi

# ── post-receive: 推送到 main 后自动 checkout 到站点目录 ──
cat > "$REPO_DIR/hooks/post-receive" <<EOF
#!/bin/bash
TARGET="$SITE_DIR"
GIT_DIR="$REPO_DIR"
while read oldrev newrev ref; do
  if [ "\$ref" = "refs/heads/main" ]; then
    git --work-tree="\$TARGET" --git-dir="\$GIT_DIR" checkout -f main >/dev/null 2>&1
    cd "\$TARGET" && git reset --hard main >/dev/null 2>&1 || true
    echo "[post-receive] 工作区已同步到 main"
  fi
done
EOF
chmod +x "$REPO_DIR/hooks/post-receive"
echo "[OK] post-receive 钩子已写入"
echo

echo "============================================"
echo " 服务器端就绪，接下来："
echo "============================================"
echo "1) 宝塔面板 → 网站 → 添加站点，根目录填: $SITE_DIR"
echo "   （只需静态站点，无需 PHP/数据库；SSL 可在面板一键申请）"
echo "2) 新电脑推送用地址（写入 deploy-config.json 的 server_url）:"
echo "   root@<服务器IP>:$SITE_DIR"
echo "3) 新电脑免密推送：把新电脑的 ~/.ssh/id_rsa.pub 追加到服务器"
echo "   ~/.ssh/authorized_keys，或直接执行:"
echo "   ssh-copy-id root@<服务器IP>"
echo
echo "验证: 新电脑执行一次『一键更新.bat』后，本目录应出现 index.html 等文件"
