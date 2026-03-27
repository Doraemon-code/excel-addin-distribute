#!/bin/sh
# scripts/install_hook.sh
#
# 一键安装 pre-commit hook 到当前 git 仓库
# 用法：sh scripts/install_hook.sh

set -e

REPO_ROOT=$(git rev-parse --show-toplevel 2>/dev/null)
if [ -z "$REPO_ROOT" ]; then
    echo "❌ 请在 git 仓库目录中运行此脚本"
    exit 1
fi

HOOK_DIR="$REPO_ROOT/.git/hooks"
HOOK_SRC="$REPO_ROOT/hooks/pre-commit"
HOOK_DST="$HOOK_DIR/pre-commit"

# 检查源文件是否存在
if [ ! -f "$HOOK_SRC" ]; then
    echo "❌ 找不到 hooks/pre-commit，请确认文件结构正确"
    exit 1
fi

# 备份已有 hook
if [ -f "$HOOK_DST" ]; then
    BACKUP="$HOOK_DST.bak_$(date +%Y%m%d_%H%M%S)"
    cp "$HOOK_DST" "$BACKUP"
    echo "💾 已备份原有 hook → $BACKUP"
fi

# 复制并赋权
cp "$HOOK_SRC" "$HOOK_DST"
chmod +x "$HOOK_DST"

echo ""
echo "✅ pre-commit hook 已安装成功！"
echo ""
echo "📋 接下来的步骤："
echo "   1. 确认已安装依赖：pip install oletools"
echo "   2. 将 .xlam 文件加入 git 追踪：git add MyAddin.xlam"
echo "   3. 正常提交：git commit -m 'your message'"
echo "      → 提交时会自动导出 VBA 代码到 src/vba/"
echo ""
