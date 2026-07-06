import json

from tqdm.auto import tqdm

from .config import LOGGER, LOG_PATH, timed_operation
from .content_actions import execute_content_action, get_content_state
from .planner import plan_action
from .property_actions import execute_property_changes, read_properties
from .selectors import select_capabilities, select_fields
from .word_schema import WordSchema

def run(task: str, max_steps: int = 20):
    print(f"\n{'=' * 60}\nTASK: {task}\n{'=' * 60}")
    LOGGER.info("TASK START | %s", task)
    setup_progress = tqdm(total=4, desc="Preparing task", unit="stage")
    try:
        setup_progress.set_postfix_str("loading Word schema")
        with timed_operation("Load Word schema"):
            schema = WordSchema()
        setup_progress.update()

        setup_progress.set_postfix_str("selecting capabilities")
        with timed_operation("Select capabilities"):
            selected_classes, inspect_properties = select_capabilities(schema, task)
        setup_progress.update()

        setup_progress.set_postfix_str("building field schema")
        with timed_operation("Build field schema"):
            fields = schema.reachable_field_schema(selected_classes)
        setup_progress.update()

        setup_progress.set_postfix_str("selecting metadata")
        with timed_operation("Select metadata fields"):
            inspected_paths = (
                select_fields(task, fields) if inspect_properties else []
            )
        setup_progress.update()
    except Exception as exc:
        setup_progress.close()
        print(f"Capability discovery failed: {exc}")
        LOGGER.exception("Capability discovery failed")
        return
    setup_progress.close()

    print(f"Relevant Word classes: {', '.join(selected_classes) or 'content commands only'}")
    print(
        f"Available schema fields: {len(fields)}; "
        f"inspecting now: {len(inspected_paths)}"
    )
    print(f"Detailed timing log: {LOG_PATH}")
    history = []

    for step in range(1, max_steps + 1):
        step_progress = tqdm(total=4, desc=f"Step {step}", unit="stage")
        step_progress.set_postfix_str("reading document")
        with timed_operation(f"Step {step}: read document content"):
            content_state = get_content_state()
        step_progress.update()
        if "error" in content_state:
            step_progress.close()
            print(f"State read failed: {content_state['error']}")
            LOGGER.error("State read failed | %s", content_state["error"])
            break

        step_progress.set_postfix_str("inspecting metadata")
        with timed_operation(f"Step {step}: inspect metadata"):
            property_state = read_properties(schema, inspected_paths)
        step_progress.update()

        print(
            f"\nStep {step}: {content_state.get('paragraph_count', 0)} paragraphs, "
            f"{content_state.get('word_count', 0)} words"
        )
        step_progress.set_postfix_str("asking Gemini")
        try:
            with timed_operation(f"Step {step}: Gemini planning"):
                action = plan_action(
                    task, fields, property_state, content_state, history
                )
        except Exception as exc:
            step_progress.close()
            print(f"Planning failed: {exc}")
            LOGGER.exception("Planning failed")
            break
        step_progress.update()

        action_type = action.get("action_type")
        print(f"Action: {action_type} — {action.get('reasoning', '')}")
        if action_type == "done":
            step_progress.update()
            step_progress.close()
            print("Task complete.")
            LOGGER.info("TASK COMPLETE | planner returned done")
            break

        step_progress.set_postfix_str(f"executing {action_type}")
        if action_type == "set_properties":
            with timed_operation(f"Step {step}: execute property changes"):
                result = execute_property_changes(
                    schema, action.get("changes", []), fields
                )
            for outcome in result.get("changes", []):
                if outcome.get("success"):
                    print(
                        f"  Verified {outcome['path']}: "
                        f"{outcome.get('before')} → {outcome.get('after')}"
                    )
                    if outcome["path"] not in inspected_paths:
                        inspected_paths.append(outcome["path"])
                else:
                    print(f"  Failed {outcome.get('path')}: {outcome.get('error')}")
        else:
            before_content = content_state
            with timed_operation(f"Step {step}: execute {action_type}"):
                result = execute_content_action(action)
            with timed_operation(f"Step {step}: verify content action"):
                after_content = get_content_state() if result["success"] else None
            result["before_content"] = {
                "word_count": before_content.get("word_count"),
                "paragraph_count": before_content.get("paragraph_count"),
                "saved": before_content.get("saved"),
            }
            if after_content and "error" not in after_content:
                result["after_content"] = {
                    "word_count": after_content.get("word_count"),
                    "paragraph_count": after_content.get("paragraph_count"),
                    "saved": after_content.get("saved"),
                }
        step_progress.update()
        step_progress.close()

        history_entry = {
            "action": action,
            "success": result.get("success", False),
            "result": result,
        }
        history.append(history_entry)
        print("Execution verified." if result.get("success") else f"Execution failed: {result}")
        LOGGER.info(
            "Step result | step=%s | action=%s | success=%s",
            step,
            action_type,
            result.get("success"),
        )

        with open("trajectory.jsonl", "a", encoding="utf-8") as handle:
            handle.write(
                json.dumps(
                    {
                        "step": step,
                        "task": task,
                        "selected_classes": selected_classes,
                        "action": action,
                        "result": result,
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    else:
        print("Maximum step count reached before the planner returned done.")
