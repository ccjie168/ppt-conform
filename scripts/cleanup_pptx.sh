#!/bin/bash
set -e

WORKSPACE_DIR="/workspace"
KEEP_DAYS=1
LOG_FILE="/workspace/.persist/cleanup.log"

mkdir -p "$(dirname "$LOG_FILE")"

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" | tee -a "$LOG_FILE"
}

log "========== 开始清理 PPTX 临时文件 =========="

EXCLUDE_PATTERNS=(
    "*/templates/*"
    "*/.git/*"
    "*/node_modules/*"
    "*/__pycache__/*"
    "*/.pytest_cache/*"
    "*/tests/*"
)

find_cmd="find $WORKSPACE_DIR -name '*.pptx' -type f -mtime +${KEEP_DAYS}"

for pattern in "${EXCLUDE_PATTERNS[@]}"; do
    find_cmd="$find_cmd ! -path '$pattern'"
done

file_count=$(eval "$find_cmd | wc -l")
total_size=$(eval "$find_cmd -exec du -ch {} + 2>/dev/null | tail -1 | cut -f1")

if [ "$file_count" -gt 0 ]; then
    log "发现 $file_count 个超过 ${KEEP_DAYS} 天的 PPTX 文件，总计约 $total_size"
    eval "$find_cmd -delete"
    log "已删除 $file_count 个文件"
else
    log "没有发现超过 ${KEEP_DAYS} 天的 PPTX 文件，无需清理"
fi

log "========== 清理完成 =========="
log ""
