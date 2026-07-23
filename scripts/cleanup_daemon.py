"""
PPTX 临时文件定时清理守护进程
每小时检查一次，删除 /workspace 下超过 24 小时的 .pptx 文件
排除项目模板、git、测试目录等
"""

import os
import time
import logging
from pathlib import Path

WORKSPACE_DIR = Path("/workspace")
KEEP_HOURS = 24
CHECK_INTERVAL = 3600
LOG_FILE = Path("/workspace/.persist/cleanup.log")

EXCLUDE_DIRS = {
    "templates",
    ".git",
    "node_modules",
    "__pycache__",
    ".pytest_cache",
    "tests",
    ".venv",
    "venv",
}

LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("pptx_cleaner")


def should_exclude(file_path: Path) -> bool:
    parts = set(file_path.relative_to(WORKSPACE_DIR).parts)
    return bool(parts & EXCLUDE_DIRS)


def cleanup_old_pptx():
    now = time.time()
    cutoff = now - (KEEP_HOURS * 3600)

    deleted_count = 0
    deleted_size = 0

    for pptx_file in WORKSPACE_DIR.rglob("*.pptx"):
        if not pptx_file.is_file():
            continue
        if should_exclude(pptx_file):
            continue
        try:
            mtime = pptx_file.stat().st_mtime
            if mtime < cutoff:
                size = pptx_file.stat().st_size
                pptx_file.unlink()
                deleted_count += 1
                deleted_size += size
                logger.info(f"已删除: {pptx_file} ({size / 1024:.1f} KB)")
        except Exception as e:
            logger.warning(f"删除失败 {pptx_file}: {e}")

    if deleted_count > 0:
        logger.info(
            f"本次清理: 删除 {deleted_count} 个文件，"
            f"释放 {deleted_size / 1024 / 1024:.2f} MB"
        )
    else:
        logger.info("本次清理: 没有需要删除的文件")


def main():
    logger.info("=" * 50)
    logger.info(f"PPTX 清理守护进程启动")
    logger.info(f"工作目录: {WORKSPACE_DIR}")
    logger.info(f"保留时长: {KEEP_HOURS} 小时")
    logger.info(f"检查间隔: {CHECK_INTERVAL} 秒")
    logger.info(f"排除目录: {', '.join(sorted(EXCLUDE_DIRS))}")
    logger.info("=" * 50)

    while True:
        try:
            cleanup_old_pptx()
        except Exception as e:
            logger.error(f"清理过程出错: {e}", exc_info=True)

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
