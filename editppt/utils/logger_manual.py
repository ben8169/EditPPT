# Saved in logfiles/ 
from datetime import datetime
from pathlib import Path

TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

BASE_DIR = Path(__file__).resolve().parents[1]  # editppt/
LOG_ROOT = BASE_DIR / "logfiles" / TIMESTAMP

def ensure_log_dir():
    LOG_ROOT.mkdir(parents=True, exist_ok=True)
    return LOG_ROOT

def log_path(filename: str) -> Path:
    """
    logfiles/{TIMESTAMP}/filename 경로 반환
    """
    ensure_log_dir()
    return LOG_ROOT / filename




# Enable Logger
import sys
from loguru import logger
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

LOG_ROOT = Path.cwd() / "logfiles" / TIMESTAMP
LOG_ROOT.mkdir(parents=True, exist_ok=True)

def init_logger():
    """
    Initialize loguru logger once.
    Safe to call multiple times (idempotent).
    """
    if getattr(init_logger, "_initialized", False):
        return logger

    logger.remove()

    logger.add(
        sys.stderr,
        level="DEBUG",
        format="<green>{time:HH:mm:ss}</green> | <level>{level}</level> | {message}",
    )
    
    logger.add(
        LOG_ROOT / "error.log",
        level="ERROR",
        encoding="utf-8",
        rotation="10 MB",
    )

    logger.add(
        LOG_ROOT / "app.log",
        level="DEBUG",
        encoding="utf-8",
        rotation="10 MB",
    )

    init_logger._initialized = True
    return logger
