import logging
import os
import time
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler
from pathlib import Path

from dotenv import load_dotenv
from tqdm.auto import tqdm

load_dotenv()

WORD_SDEF_PATHS = (
    Path("/Applications/Microsoft Word.app/Contents/Resources/Word.sdef"),
    Path("/Applications/Microsoft Word.app/Contents/Resources/English.lproj/Word.sdef"),
)

OPENROUTER_TIMEOUT_SECONDS = 60
SLOW_OPERATION_SECONDS = 5
LOG_PATH = Path(__file__).resolve().parent.parent / "scriptagent.log"

DEFAULT_OPENROUTER_MODELS = (
    "openai/gpt-4o-mini",
    "anthropic/claude-3.5-haiku",
    "google/gemini-2.5-flash",
)


def get_openrouter_models() -> tuple[str, ...]:
    raw = os.environ.get("OPENROUTER_MODELS", "")
    models = tuple(item.strip() for item in raw.split(",") if item.strip())
    return models or DEFAULT_OPENROUTER_MODELS


MODELS = get_openrouter_models()


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("scriptagent")
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    handler = RotatingFileHandler(
        LOG_PATH,
        maxBytes=2_000_000,
        backupCount=3,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    )
    logger.addHandler(handler)
    return logger


LOGGER = configure_logging()


@contextmanager
def timed_operation(label: str):
    started = time.monotonic()
    LOGGER.info("START | %s", label)
    try:
        yield
    except Exception:
        LOGGER.exception("FAILED | %s | %.2fs", label, time.monotonic() - started)
        raise

    elapsed = time.monotonic() - started
    LOGGER.info("DONE | %s | %.2fs", label, elapsed)
    if elapsed >= SLOW_OPERATION_SECONDS:
        tqdm.write(f"  Slow operation: {label} took {elapsed:.1f}s")