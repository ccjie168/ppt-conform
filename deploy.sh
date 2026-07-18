#!/bin/bash
# 自动化部署脚本：提交代码 → 推送 GitHub → 触发 Streamlit Cloud 自动更新
# 用法: ./deploy.sh "提交说明"
# 环境变量: GITHUB_TOKEN (必需) - GitHub Personal Access Token
#   可在 ~/.bashrc 中添加: export GITHUB_TOKEN="ghp_your_token_here"

set -e

cd "$(dirname "$0")"

# 颜色输出
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${YELLOW}🚀 开始自动化部署流程${NC}"

# 1. 检查是否有变更
if git diff --quiet && git diff --cached --quiet && [ -z "$(git ls-files --others --exclude-standard)" ]; then
    echo -e "${YELLOW}⚠️  没有检测到代码变更，无需部署${NC}"
    exit 0
fi

# 2. 检查 GitHub Token（从环境变量读取，避免泄露到代码库）
if [ -z "$GITHUB_TOKEN" ]; then
    echo -e "${RED}❌ 未设置 GITHUB_TOKEN 环境变量${NC}"
    echo -e "${YELLOW}   请运行: export GITHUB_TOKEN=\"ghp_your_token_here\"${NC}"
    echo -e "${YELLOW}   或在 ~/.bashrc 中添加该命令使其永久生效${NC}"
    exit 1
fi

GITHUB_USER="ccjie168"
REPO_NAME="ppt-conform"

# 3. 添加变更文件（排除敏感文件）
echo -e "${YELLOW}📦 步骤 1/4: 暂存代码变更${NC}"
git add -A
git status --short

# 4. 提交
COMMIT_MSG="${1:-auto: update $(date '+%Y-%m-%d %H:%M:%S')}"
echo -e "${YELLOW}💾 步骤 2/4: 提交代码${NC}"
git commit -m "$COMMIT_MSG"

# 5. 推送（使用环境变量中的 Token 认证）
echo -e "${YELLOW}📤 步骤 3/4: 推送到 GitHub${NC}"
git remote set-url origin "https://${GITHUB_USER}:${GITHUB_TOKEN}@github.com/${GITHUB_USER}/${REPO_NAME}.git"
git push origin master 2>&1 | sed "s/${GITHUB_TOKEN}/***TOKEN***/g"
# 清理 URL 中的 Token
git remote set-url origin "https://github.com/${GITHUB_USER}/${REPO_NAME}.git"

echo -e "${GREEN}✅ 代码已推送到 GitHub${NC}"
echo -e "${GREEN}   仓库: https://github.com/${GITHUB_USER}/${REPO_NAME}${NC}"

# 6. 提示 Streamlit Cloud 自动部署
echo -e "${YELLOW}☁️  步骤 4/4: Streamlit Cloud 自动部署${NC}"
echo -e "${YELLOW}   Streamlit Community Cloud 会自动检测 GitHub 仓库更新${NC}"
echo -e "${YELLOW}   通常在 1-3 分钟内完成重新部署${NC}"
echo ""
echo -e "${GREEN}🎉 部署流程完成！${NC}"
echo -e "${GREEN}   Streamlit 应用地址: https://${GITHUB_USER}-${REPO_NAME}.streamlit.app${NC}"
echo -e "${YELLOW}   如果应用未自动更新，请访问 https://share.streamlit.io 手动 Reboot${NC}"
