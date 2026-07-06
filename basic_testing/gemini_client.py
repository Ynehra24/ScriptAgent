import json
import os
import time

from google import genai
from google.genai import types

from .config import GEMINI_TIMEOUT_MS, LOGGER, MODELS

def get_client():
    key = os.environ.get("GEMINI_KEY")
    if not key:
        raise RuntimeError("GEMINI_KEY is not set.")
    return genai.Client(
        api_key=key,
        http_options=types.HttpOptions(timeout=GEMINI_TIMEOUT_MS),
    )

def call_gemini(system_prompt: str, prompt: str, max_tokens: int = 4096) -> dict:
    client = get_client()
    messages = [{"role": "user", "content": prompt}]
    model_index = 0
    last_error = None

    for attempt in range(3):
        raw = None
        for _ in MODELS:
            model = MODELS[model_index % len(MODELS)]
            started = time.monotonic()
            LOGGER.info(
                "Gemini request | model=%s | attempt=%s | prompt_chars=%s",
                model,
                attempt + 1,
                len(prompt),
            )
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=[
                        types.Content(
                            role="model" if message["role"] == "assistant" else "user",
                            parts=[types.Part.from_text(text=message["content"])],
                        )
                        for message in messages
                    ],
                    config=types.GenerateContentConfig(
                        system_instruction=system_prompt,
                        max_output_tokens=max_tokens,
                        response_mime_type="application/json",
                    ),
                )
                raw = response.text.strip()
                LOGGER.info(
                    "Gemini response | model=%s | %.2fs | response_chars=%s",
                    model,
                    time.monotonic() - started,
                    len(raw),
                )
                break
            except Exception as exc:
                last_error = exc
                LOGGER.warning(
                    "Gemini failure | model=%s | %.2fs | %s",
                    model,
                    time.monotonic() - started,
                    exc,
                )
                error_text = str(exc).lower()
                if any(
                    marker in error_text
                    for marker in (
                        "429",
                        "500",
                        "502",
                        "503",
                        "504",
                        "rate limit",
                        "unavailable",
                        "high demand",
                        "temporarily",
                    )
                ):
                    model_index += 1
                    continue
                raise
        if raw is None:
            if attempt < 2:
                continue
            raise RuntimeError(
                f"All Gemini models were temporarily unavailable: {last_error}"
            )
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            if attempt == 2:
                raise
            messages.extend(
                [
                    {"role": "assistant", "content": raw},
                    {
                        "role": "user",
                        "content": "Return one valid JSON object only.",
                    },
                ]
            )
    raise RuntimeError("Gemini did not return valid JSON.")
