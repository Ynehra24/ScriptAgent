import json

from tqdm.auto import tqdm

from .config import LOGGER, LOG_PATH, timed_operation
from .content_actions import get_content_state
from .execution_runtime import execute_planner_output
from .planner import plan_action
from .property_actions import read_properties
from .selectors import select_capabilities, select_fields
from .word_schema import WordSchema


def _summarize_content_state(content_state: dict) -> dict:
    return {
        "word_count": content_state.get("word_count"),
        "paragraph_count": content_state.get("paragraph_count"),
        "saved": content_state.get("saved"),
        "content_window": content_state.get("content_window"),
    }


def _print_execution_result(result: dict):
    if result.get("success"):
        print("Execution verified.")
        return

    print(f"Execution failed: {result}")


def _append_history(
    history: list[dict],
    action: dict,
    result: dict,
    content_before: dict | None = None,
    content_after: dict | None = None,
):
    entry = {
        "action": action,
        "success": result.get("success", False),
        "result": result,
    }

    if content_before:
        entry["before_content"] = _summarize_content_state(content_before)
    if content_after:
        entry["after_content"] = _summarize_content_state(content_after)

    history.append(entry)


def _write_trajectory(
    step: int,
    task: str,
    selected_classes: list[str],
    action: dict,
    result: dict,
):
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

    print(f"Relevant Word classes: {', '.join(selected_classes) or 'content/context only'}")
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

        step_progress.set_postfix_str("asking planner")
        try:
            with timed_operation(f"Step {step}: capability planning"):
                action = plan_action(
                    task=task,
                    schema=schema,
                    fields=fields,
                    selected_classes=selected_classes,
                    property_state=property_state,
                    content_state=content_state,
                    history=history,
                )
        except Exception as exc:
            step_progress.close()
            print(f"Planning failed: {exc}")
            LOGGER.exception("Planning failed")
            break

        step_progress.update()

        action_type = action.get("type")
        print(f"Planner output: {action_type} — {action.get('reasoning', '')}")

        if action_type == "execution_plan" and not action.get("steps"):
            step_progress.update()
            step_progress.close()
            print("Task complete or no safe executable steps were available.")
            LOGGER.info("TASK COMPLETE | planner returned empty execution plan")
            break

        step_progress.set_postfix_str(f"executing {action_type}")
        before_content = content_state

        with timed_operation(f"Step {step}: execute planner output"):
            result = execute_planner_output(schema, action, fields)

        after_content = None

        if action_type == "context_request":
            step_progress.update()
            step_progress.close()

            _append_history(
                history=history,
                action=action,
                result=result,
                content_before=before_content,
            )
            _write_trajectory(step, task, selected_classes, action, result)

            if result.get("success"):
                print("Context retrieved. Replanning with expanded history.")
                LOGGER.info("Context request satisfied | step=%s", step)
                continue

            _print_execution_result(result)
            LOGGER.info(
                "Step result | step=%s | type=%s | success=%s",
                step,
                action_type,
                result.get("success"),
            )
            continue

        with timed_operation(f"Step {step}: verify planner output"):
            after_content = get_content_state() if result.get("success") else None

        if after_content and "error" not in after_content:
            result["after_content"] = _summarize_content_state(after_content)
        result["before_content"] = _summarize_content_state(before_content)

        step_progress.update()
        step_progress.close()

        _append_history(
            history=history,
            action=action,
            result=result,
            content_before=before_content,
            content_after=after_content,
        )
        _write_trajectory(step, task, selected_classes, action, result)

        _print_execution_result(result)
        LOGGER.info(
            "Step result | step=%s | type=%s | success=%s",
            step,
            action_type,
            result.get("success"),
        )

        if action_type == "execution_plan" and result.get("success"):
            # Let the planner inspect the updated state and decide whether done.
            continue

    else:
        print("Maximum step count reached before the planner returned done.")
