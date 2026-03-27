"""
HRMS Logging Configuration
Конфигурация логирования системы
"""
from loguru import logger
import sys
from pathlib import Path
import settings

# Ensure logs directory exists
settings.ensure_directories()

# Remove default handler
logger.remove()

# Console handler (INFO and above)
# In some environments (like xlwings non-debug), sys.stderr can be None
if sys.stderr:
    logger.add(
        sys.stderr,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> | <level>{message}</level>",
        level="INFO",
        colorize=True
    )

# File handler (all logs)
logger.add(
    settings.LOGS_DIR / "hrms_{time:YYYY-MM-DD}.log",
    rotation=settings.LOG_ROTATION_SIZE,
    retention=f"{settings.LOG_RETENTION_DAYS} days",
    level="INFO",
    format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} | {message}",
    encoding="utf-8"
)

# Debug mode handler (if enabled)
if settings.DEBUG_MODE:
    logger.add(
        settings.LOGS_DIR / "hrms_debug_{time:YYYY-MM-DD}.log",
        rotation=settings.LOG_ROTATION_SIZE,
        retention="7 days",
        level="DEBUG",
        format="{time:YYYY-MM-DD HH:mm:ss.SSS} | {level: <8} | {name}:{function}:{line} | {message}",
        encoding="utf-8"
    )
    logger.info("Debug mode enabled")

logger.info("HRMS logging initialized")

# Export logger
__all__ = ["logger"]
