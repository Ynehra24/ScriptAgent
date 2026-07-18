"""Planner utilities extracted from tester.py for action planning.

Provides `plan_action` that formulates the JSON action given a task and state.
This module only contains the prompt and a helper wrapper; actual LLM calls
are kept out so you can wire your preferred model client in the caller.
"""
from typing import Dict, List, Optional

from .word_sdef import search_catalog

SYSTEM_PROMPT = """
You are a desktop automation agent controlling Microsoft Word through its native API.
You do NOT take screenshots. You receive the full document state as structured JSON
and output a single structured action to perform.

Available actions:
- insert_text: params: {text, position (cursor|start|end)}
- select_paragraph: params: {paragraph (int, 1-indexed)} — selects an entire paragraph
- format_text: params: {bold (bool), italic (bool), font_size (int), font_name (str), paragraph (int, 1-indexed, optional)}
    If "paragraph" is provided, formatting is applied directly to that paragraph.
    If omitted, formatting is applied to the current selection.
- set_style: params: {style (e.g. "Heading 1", "Normal", "Heading 2")}
- find_replace: params: {find (str), replace (str)}
- find_format: params: {formats (list of objects with keys: {find (str), bold (bool, optional), italic (bool, optional)})}
- delete_paragraph: params: {paragraph (int, 1-indexed)} — deletes the specified paragraph
- save_file: params: {}
- apple_script: params: {script (str)} — raw AppleScript fallback for any Word capability not covered by the typed actions

Output ONLY a JSON object with this structure:
{
  "reasoning": "...",
  "action_type": "...",
  "params": {...},
  "expected_state_delta": {...}
}

Rules:
- Always use API actions, never suggest keystrokes
- NEVER insert text containing raw markdown characters like **, *, _, or ***.
  Any formatting (bold, italic) must be done programmatically using find_format after the text is inserted.
- To format an entire paragraph, use format_text with a "paragraph" param. To format specific words or phrases (such as individual keywords or nouns) within the document, use find_format. Do NOT try to select or format individual words using format_text.
- Differentiating Keywords and Nouns: Pay extreme attention to the user task for formatting instructions. If a task asks for nouns to be both bold and italic, make sure to set both "bold": true and "italic": true on those nouns.
- Identifying Keywords and Nouns: When formatting keywords (bold) and nouns (both bold and italic), identify ALL relevant terms and format them in a single find_format action using the "formats" list parameter. Do not stop after formatting just a few.
- Keep expected_state_delta simple — use only top-level keys like word_count, paragraph_count, saved. Do NOT put complex nested objects.
- If task is already complete based on state, output action_type: "done"
- If the task requires a Word capability not directly covered by the typed actions, use apple_script with a minimal script that performs the exact Word action.
- Output raw JSON only, no markdown fences
"""


def _format_relevant_catalog(task: str) -> str:
    catalog = search_catalog(task)
    lines: List[str] = []

    if catalog["commands"]:
        lines.append("RELEVANT WORD COMMANDS:")
        for command in catalog["commands"][:10]:
            parameters = ", ".join(parameter["name"] for parameter in command.get("parameters", [])) or "none"
            lines.append(f'- {command["name"]}: {command["description"]} | parameters: {parameters}')

    if catalog["classes"]:
        lines.append("RELEVANT WORD CLASSES/PROPERTIES:")
        for class_item in catalog["classes"][:10]:
            property_names = ", ".join(property_item["name"] for property_item in class_item.get("properties", [])[:8])
            lines.append(f'- {class_item["name"]}: {class_item["description"]} | properties: {property_names}')

    if catalog["enumerations"]:
        lines.append("RELEVANT WORD ENUMERATIONS:")
        for enumeration in catalog["enumerations"][:5]:
            enumerator_names = ", ".join(enumerator["name"] for enumerator in enumeration.get("enumerators", [])[:12])
            lines.append(f'- {enumeration["name"]}: {enumerator_names}')

    return "\n".join(lines)


def plan_action_prompt(task: str, state: Dict, history: Optional[List[Dict]] = None) -> str:
    """Construct the full prompt to send to an LLM for planning.

    This function returns the textual prompt; caller is responsible for running the LLM.
    """
    compact_state = dict(state)
    if "paragraphs" in compact_state:
        compact_state["paragraphs"] = [
            {**p, "text": p["text"][:120] + ("..." if len(p["text"]) > 120 else "")}
            for p in compact_state["paragraphs"]
        ]

    prompt = f"""
SYSTEM:
{SYSTEM_PROMPT}

USER TASK: {task}

CURRENT STATE:
{compact_state}

{_format_relevant_catalog(task)}

Return only the JSON action.
"""
    return prompt
