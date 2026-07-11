import json

from .gemini_client import call_llm_json
from .prompts import PLANNER_PROMPT
from .property_actions import compact_property_state
from .execution_runtime import EXECUTOR_CAPABILITIES


CONTEXT_RETRIEVAL_CAPABILITIES = {
    "inspect_object_window": {
        "description": "Retrieve a bounded window of document context for an object type.",
        "arguments": {
            "object": "paragraph|table|section|style|header|footer|comment",
            "start": "1-indexed starting object",
            "limit": "number of objects to inspect",
            "preview_chars": "maximum text returned per object",
        },
    },
    "inspect_object_detail": {
        "description": "Retrieve deeper context for one document object.",
        "arguments": {
            "object": "paragraph|table|section|style|header|footer|comment",
            "index": "1-indexed object index",
            "preview_chars": "maximum text returned",
        },
    },
}


def build_capability_graph(
    schema,
    fields: dict[str, dict],
    selected_classes: list[str] | None = None,
) -> dict:
    selected = {name for name in selected_classes or [] if name in schema.classes}
    selected.update(
        field.get("owner_class")
        for field in fields.values()
        if field.get("owner_class") in schema.classes
    )
    selected.update(
        name
        for name in (
            "document",
            "selection object",
            "paragraph",
            "text range",
            "table",
            "row",
            "cell",
            "font",
            "paragraph format",
            "border options",
        )
        if name in schema.classes
    )

    objects = {}
    for name in sorted(selected):
        info = schema.classes[name]
        properties = [
            prop for prop in schema.inherited_properties(name)
            if prop.get("name")
        ]
        objects[name] = {
            "name": name,
            "plural": info.get("plural"),
            "inherits": info.get("inherits"),
            "description": info.get("description", ""),
            "elements": info.get("elements", []),
            "properties": [
                {
                    "name": prop.get("name"),
                    "type": prop.get("type"),
                    "list": prop.get("list"),
                    "writable": prop.get("writable"),
                    "description": prop.get("description", ""),
                    "enum_values": schema.enum_values(prop.get("type")),
                }
                for prop in properties
            ],
        }

    commands = {
        name: {
            "name": command.get("name"),
            "description": command.get("description", ""),
            "parameters": command.get("parameters", []),
        }
        for name, command in schema.commands.items()
    }

    properties = {
        path: {
            "path": path,
            "owner_class": field.get("owner_class"),
            "word_name": field.get("word_name"),
            "type": field.get("type"),
            "list": field.get("list"),
            "writable": field.get("writable"),
            "enum_values": field.get("enum_values", []),
            "description": field.get("description", ""),
        }
        for path, field in fields.items()
    }

    return {
        "source": "installed Microsoft Word SDEF",
        "objects": objects,
        "commands": commands,
        "properties": properties,
        "enumerations": schema.enumerations,
    }


def build_verification_policy() -> dict:
    return {
        "property_write": "Read the canonical property before and after writing.",
        "command_invocation": "Execute the command, then reread affected object/content state.",
        "context_retrieval": "Append retrieved context to history and replan.",
        "completion": "Return done only when inspected state proves completion or capability is unavailable.",
    }


def build_planning_prompt(
    task: str,
    schema,
    fields: dict[str, dict],
    selected_classes: list[str],
    property_state: dict,
    content_state: dict,
    history: list[dict],
) -> dict:
    return {
        "task": task,
        "planner_contract": {
            "responsibility": [
                "reason about the user task",
                "decide whether more context is needed",
                "decompose into ordered capability invocations",
                "schedule safe execution steps",
            ],
            "must_not": [
                "emit AppleScript",
                "invent objects, commands, properties, or enum values",
                "use action_type planning",
                "use commands that are not in executor_capabilities unless the command is present in the SDEF and can target object 1 of document 1 directly",
                "move execution logic into planning",
            ],
        },
        "output_schema": {
            "context_request": {
                "type": "context_request",
                "retrieval": {
                    "capability": "inspect_object_window|inspect_object_detail",
                    "arguments": {},
                },
                "reasoning": "...",
            },
            "execution_plan": {
                "type": "execution_plan",
                "steps": [
                    {
                        "object": "exact SDEF object name",
                        "command": "one executor_capabilities command",
                        "target": "canonical property path or object reference",
                        "arguments": {},
                        "verification": "property_write|command_invocation|completion",
                    }
                ],
                "reasoning": "...",
            },
        },
        "executor_capabilities": EXECUTOR_CAPABILITIES,
        "context_retrieval_capabilities": CONTEXT_RETRIEVAL_CAPABILITIES,
        "dynamic_sdef_index": schema.focused_context(task, selected_classes),
        "capability_graph": build_capability_graph(schema, fields, selected_classes),
        "verification_policy": build_verification_policy(),
        "inspected_property_state": compact_property_state(property_state),
        "document_content_state": content_state,
        "history": history[-12:],
    }


def plan_action(
    task: str,
    schema,
    fields: dict[str, dict],
    selected_classes: list[str],
    property_state: dict,
    content_state: dict,
    history: list[dict],
) -> dict:
    prompt = build_planning_prompt(
        task=task,
        schema=schema,
        fields=fields,
        selected_classes=selected_classes,
        property_state=property_state,
        content_state=content_state,
        history=history,
    )
    return call_llm_json(PLANNER_PROMPT, json.dumps(prompt, ensure_ascii=False))
