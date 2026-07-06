import json

from .gemini_client import call_gemini
from .prompts import PLANNER_PROMPT
from .property_actions import compact_property_state

def plan_action(
    task: str,
    fields: dict[str, dict],
    property_state: dict,
    content_state: dict,
    history: list[dict],
) -> dict:
    prompt = {
        "task": task,
        "available_property_schema": [
            fields[path] for path in property_state if path in fields
        ],
        "inspected_property_state": compact_property_state(property_state),
        "document_content_state": content_state,
        "history": history[-12:],
    }
    return call_gemini(PLANNER_PROMPT, json.dumps(prompt, ensure_ascii=False))
