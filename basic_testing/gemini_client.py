import json
import os
import time
import urllib.error
import urllib.request

from .config import LOGGER, MODELS, OPENROUTER_TIMEOUT_SECONDS


OPENROUTER_CHAT_URL = "https://openrouter.ai/api/v1/chat/completions"


def get_openrouter_key() -> str:
    key = os.environ.get("OPENROUTER_API_KEY")
    if not key:
        raise RuntimeError("OPENROUTER_API_KEY is not set.")
    return key


def _request_openrouter(
    model: str,
    system_prompt: str,
    prompt: str,
    max_tokens: int,
) -> str:
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0,
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"},
    }

    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        OPENROUTER_CHAT_URL,
        data=body,
        method="POST",
        headers={
            "Authorization": f"Bearer {get_openrouter_key()}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://localhost/uivlm",
            "X-Title": "UIVLM ScriptAgent",
        },
    )

    with urllib.request.urlopen(
        request,
        timeout=OPENROUTER_TIMEOUT_SECONDS,
    ) as response:
        data = json.loads(response.read().decode("utf-8"))

    try:
        return data["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError(f"OpenRouter response did not contain message content: {data}") from exc


def _error_is_retryable(error_text: str) -> bool:
    error_text = error_text.lower()
    return any(
        marker in error_text
        for marker in (
            "400",
            "404",
            "408",
            "409",
            "429",
            "500",
            "502",
            "503",
            "504",
            "rate limit",
            "temporarily",
            "timeout",
            "timed out",
            "unavailable",
            "not found",
            "not supported",
            "no endpoints",
            "provider",
        )
    )


def call_llm_json(system_prompt: str, prompt: str, max_tokens: int = 4096) -> dict:
    models = [MODELS] if isinstance(MODELS, str) else list(MODELS)
    if not models:
        raise RuntimeError("No OpenRouter models configured.")

    model_index = 0
    last_error = None
    messages_for_repair = prompt

    for attempt in range(3):
        raw = None

        for _ in models:
            model = models[model_index % len(models)]
            started = time.monotonic()

            LOGGER.info(
                "OpenRouter request | model=%s | attempt=%s | prompt_chars=%s",
                model,
                attempt + 1,
                len(messages_for_repair),
            )

            try:
                raw = _request_openrouter(
                    model=model,
                    system_prompt=system_prompt,
                    prompt=messages_for_repair,
                    max_tokens=max_tokens,
                )
                LOGGER.info(
                    "OpenRouter response | model=%s | %.2fs | response_chars=%s",
                    model,
                    time.monotonic() - started,
                    len(raw),
                )
                break

            except urllib.error.HTTPError as exc:
                error_body = exc.read().decode("utf-8", errors="replace")
                last_error = f"{exc.code} {error_body}"
                LOGGER.warning(
                    "OpenRouter HTTP failure | model=%s | %.2fs | %s",
                    model,
                    time.monotonic() - started,
                    last_error,
                )
                if _error_is_retryable(last_error):
                    model_index += 1
                    continue
                raise RuntimeError(last_error) from exc

            except Exception as exc:
                last_error = exc
                LOGGER.warning(
                    "OpenRouter failure | model=%s | %.2fs | %s",
                    model,
                    time.monotonic() - started,
                    exc,
                )
                if _error_is_retryable(str(exc)):
                    model_index += 1
                    continue
                raise

        if raw is None:
            if attempt < 2:
                continue
            raise RuntimeError(f"All OpenRouter models failed: {last_error}")

        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            if attempt == 2:
                raise
            messages_for_repair = (
                f"{prompt}\n\nThe previous response was not valid JSON:\n"
                f"{raw}\n\nReturn exactly one valid JSON object only."
            )

    raise RuntimeError("OpenRouter did not return valid JSON.")


def call_gemini(system_prompt: str, prompt: str, max_tokens: int = 4096) -> dict:
    return call_llm_json(system_prompt, prompt, max_tokens)
