import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logging(
    *,
    level=logging.INFO,
    log_file="app.log",
    max_bytes=5_000_000,
    backup_count=5,
):
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    # Prevent duplicate handlers if setup_logging() is called twice
    if root_logger.handlers:
        return

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
    )

    # ---- Console handler ----
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)

    # ---- File handler (rotating) ----
    file_handler = RotatingFileHandler(
        log_dir / log_file,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)

    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)
