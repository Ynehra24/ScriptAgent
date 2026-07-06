import json

from .gemini_client import call_gemini
from .prompts import CAPABILITY_SELECTOR_PROMPT, FIELD_SELECTOR_PROMPT
from .word_schema import WordSchema

def select_capabilities(schema: WordSchema, task: str) -> tuple[list[str], bool]:
    response = call_gemini(
        CAPABILITY_SELECTOR_PROMPT,
        f"TASK:\n{task}\n\nWORD CLASS CATALOG:\n"
        f"{json.dumps(schema.class_index(), ensure_ascii=False)}",
    )
    classes = [
        name
        for name in response.get("classes", [])
        if name in schema.classes
    ][:10]
    return classes, response.get("inspect_properties", True) is not False

def select_fields(task: str, fields: dict[str, dict]) -> list[str]:
    response = call_gemini(
        FIELD_SELECTOR_PROMPT,
        f"TASK:\n{task}\n\nAVAILABLE FIELDS:\n"
        f"{json.dumps(list(fields.values()), ensure_ascii=False)}",
    )
    selected = [
        path
        for path in response.get("paths", [])
        if path in fields
    ]
    non_wildcard = [path for path in selected if "[*]" not in path]
    wildcard = [path for path in selected if "[*]" in path]
    return (non_wildcard + wildcard)[:8]
