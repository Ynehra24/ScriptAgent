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


def ensure_app_running(app_name: str = "Microsoft Word") -> None:
    """Try to ensure the named app is running; best-effort, no exception on failure."""
    try:
        # Use `open -a` which will launch the app if not running
        subprocess.run(["open", "-a", app_name], check=False, capture_output=True)
    except Exception:
        LOGGER.exception("Failed to ensure %s running", app_name)


def run_applescript_safe(script: str, timeout: int = 20, retries: int = 2, restart_on_error: bool = True) -> tuple[str, str, int]:
    """Run AppleScript with simple retry and app-restart logic for flaky apps like Word.

    - If the call fails and `restart_on_error` is True the function will attempt to
      open the app and retry up to `retries` times.
    - Returns the same tuple as `run_applescript`.
    """
    last_stdout = ""
    last_stderr = ""
    last_code = 1
    attempt = 0
    while attempt <= retries:
        attempt += 1
        try:
            stdout, stderr, code = run_applescript(script, timeout=timeout)
        except Exception as exc:
            # Unexpected exception from subprocess; capture and treat as failure
            LOGGER.exception("run_applescript threw: %s", exc)
            last_stdout = ""
            last_stderr = str(exc)
            last_code = 1
        else:
            last_stdout, last_stderr, last_code = stdout, stderr, code

        if last_code == 0:
            return last_stdout, last_stderr, last_code

        # If we shouldn't try to restart, break out and return the error
        if not restart_on_error:
            return last_stdout, last_stderr, last_code

        # Try to detect common fatal conditions where relaunching may help
        err_lower = (last_stderr or "").lower()
        if any(token in err_lower for token in ("not running", "is not running", "no such process", "connection refused", "-600", "killed")) or attempt <= retries:
            LOGGER.warning("AppleScript failed (code=%s). Attempting to ensure app is running and retry (%s/%s)", last_code, attempt, retries)
            ensure_app_running()
            time.sleep(0.8 * attempt)
            continue

        # Otherwise return the last result
        return last_stdout, last_stderr, last_code

    return last_stdout, last_stderr, last_code
