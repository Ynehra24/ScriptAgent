from .applescript import run_applescript
from .content_actions import (
    delete_paragraph,
    execute_context_request,
    format_paragraph,
    insert_table,
    insert_text,
    replace_paragraph,
)
from .property_actions import execute_property_changes


EXECUTOR_CAPABILITIES = {
    "insert_text": {
        "description": "Insert text at document start, document end, or current selection.",
        "arguments": {
            "text": "text to insert",
            "position": "start|end|selection",
            "font_name": "optional font family",
            "font_size": "optional font size",
            "bold": "optional boolean",
            "italic": "optional boolean",
        },
    },
    "replace_paragraph": {
        "description": "Replace one 1-indexed paragraph's text content.",
        "arguments": {"paragraph": "1-indexed paragraph", "text": "replacement text"},
    },
    "delete_paragraph": {
        "description": "Delete one 1-indexed paragraph.",
        "arguments": {"paragraph": "1-indexed paragraph"},
    },
    "format_paragraph": {
        "description": "Apply simple font formatting to one paragraph.",
        "arguments": {
            "paragraph": "1-indexed paragraph",
            "font_name": "optional font family",
            "font_size": "optional font size",
            "bold": "optional boolean",
            "italic": "optional boolean",
        },
    },
    "insert_table": {
        "description": "Insert a table, optionally replacing an existing paragraph, with optional cell shading.",
        "arguments": {
            "rows": "row count",
            "columns": "column count",
            "values": "rows x columns nested list",
            "paragraph": "optional paragraph to replace",
            "position": "start|end|selection when paragraph is omitted",
            "alternating_cell_colors": "optional list of color names or RGB lists, e.g. ['maroon', 'grey']",
            "cell_colors": "optional rows x columns nested list of color names or RGB lists",
        },
    },
    "property_write": {
        "description": "Set one canonical writable property path from the active field schema.",
        "arguments": {"value": "new scalar/list value"},
    },
}


CONTENT_CAPABILITY_HANDLERS = {
    "insert_text": insert_text,
    "replace_paragraph": replace_paragraph,
    "delete_paragraph": delete_paragraph,
    "format_paragraph": format_paragraph,
    "insert_table": insert_table,
}


def execute_planner_output(schema, planner_output: dict, available_fields: dict) -> dict:
    output_type = planner_output.get("type")

    if output_type == "context_request":
        return execute_context_request(planner_output.get("retrieval", {}))

    if output_type == "execution_plan":
        return execute_execution_plan(schema, planner_output, available_fields)

    return {
        "success": False,
        "error": f"Unsupported planner output type: {output_type}",
    }


def execute_execution_plan(schema, plan: dict, available_fields: dict) -> dict:
    outcomes = []
    steps = plan.get("steps", [])

    if not steps:
        return {
            "success": True,
            "steps": [],
            "note": "Planner returned an empty execution plan.",
        }

    for index, step in enumerate(steps, 1):
        result = execute_capability_step(schema, step, available_fields)
        outcomes.append(
            {
                "step": index,
                "capability": step,
                "result": result,
            }
        )

        if not result.get("success"):
            break

    return {
        "success": bool(outcomes) and all(
            item["result"].get("success") for item in outcomes
        ),
        "steps": outcomes,
    }


def execute_capability_step(schema, step: dict, available_fields: dict) -> dict:
    command = step.get("command")

    if command == "property_write":
        return execute_property_write(step, schema, available_fields)

    if command in CONTENT_CAPABILITY_HANDLERS:
        return CONTENT_CAPABILITY_HANDLERS[command](step.get("arguments", {}))

    return execute_sdef_command(schema, step)


def execute_property_write(step: dict, schema, available_fields: dict) -> dict:
    target = step.get("target")
    arguments = step.get("arguments", {})

    if target not in available_fields:
        return {
            "success": False,
            "error": f"Unknown or unavailable property path: {target}",
        }

    if not available_fields[target].get("writable"):
        return {
            "success": False,
            "error": f"Property path is read-only: {target}",
        }

    return execute_property_changes(
        schema,
        [{"path": target, "value": arguments.get("value")}],
        available_fields,
    )


def execute_sdef_command(schema, step: dict) -> dict:
    command = step.get("command")
    target_object = step.get("object")
    arguments = step.get("arguments", {})

    if command not in schema.commands:
        return {
            "success": False,
            "error": f"Command is not present in Word SDEF: {command}",
        }

    if target_object and target_object not in schema.classes:
        return {
            "success": False,
            "error": f"Object is not present in Word SDEF: {target_object}",
        }

    script = compile_sdef_command(command, target_object, arguments)
    stdout, stderr, code = run_applescript(script)

    return {
        "success": code == 0,
        "error": stderr if code != 0 else None,
        "stdout": stdout,
    }


def compile_sdef_command(command: str, target_object: str | None, arguments: dict) -> str:
    # Generic first-pass command compiler. It validates against SDEF before this
    # point, but intentionally does not hardcode individual Word command names.
    rendered_args = []
    for key, value in arguments.items():
        rendered_args.append(f'{key}:{applescript_value(value)}')

    argument_text = ""
    if rendered_args:
        argument_text = " given " + ", ".join(rendered_args)

    if target_object:
        target = f"{target_object} 1 of document 1"
        command_line = f"{command} {target}{argument_text}"
    else:
        command_line = f"{command}{argument_text}"

    return f'tell application "Microsoft Word"\n{command_line}\nend tell'


def applescript_value(value) -> str:
    if value is None:
        return "missing value"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        return "{" + ", ".join(applescript_value(item) for item in value) + "}"
    escaped = str(value).replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'
