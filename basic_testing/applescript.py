import os
import subprocess
import tempfile
import time

from tqdm.auto import tqdm

from .config import LOGGER, SLOW_OPERATION_SECONDS

def run_applescript(script: str, timeout: int = 20) -> tuple[str, str, int]:
    with tempfile.NamedTemporaryFile(mode="w", suffix=".applescript", delete=False) as handle:
        handle.write(script)
        temporary_path = handle.name
    try:
        started = time.monotonic()
        try:
            result = subprocess.run(
                ["osascript", temporary_path],
                capture_output=True,
                text=True,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired:
            elapsed = time.monotonic() - started
            LOGGER.error("AppleScript timed out after %.2fs", elapsed)
            return "", f"AppleScript timed out after {timeout}s.", 124
        elapsed = time.monotonic() - started
        LOGGER.debug(
            "AppleScript finished | code=%s | %.2fs", result.returncode, elapsed
        )
        if elapsed >= SLOW_OPERATION_SECONDS:
            tqdm.write(f"  Slow Word operation took {elapsed:.1f}s")
        return result.stdout.strip(), result.stderr.strip(), result.returncode
    finally:
        os.unlink(temporary_path)
