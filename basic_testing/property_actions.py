import json
import time

from tqdm.auto import tqdm

from .config import LOGGER
from .property_path import PropertyPath
from .word_schema import WordSchema

def read_properties(schema: WordSchema, paths: list[str]) -> dict:
    state = {}
    progress = tqdm(
        paths,
        desc="Inspecting Word metadata",
        unit="field",
        leave=False,
        disable=not paths,
    )
    for path in progress:
        progress.set_postfix_str(path.rsplit(".", 1)[-1][:24])
        started = time.monotonic()
        try:
            state[path] = PropertyPath(schema, path).read()
        except Exception as exc:
            state[path] = {"error": str(exc)}
            LOGGER.exception("Property read failed | %s", path)
        finally:
            LOGGER.info(
                "Property read | %s | %.2fs", path, time.monotonic() - started
            )
    return state

def values_match(actual, expected) -> bool:
    if actual == expected:
        return True
    if isinstance(actual, list) and isinstance(expected, list):
        if actual and all(isinstance(item, list) for item in actual):
            return all(values_match(item, expected) for item in actual)
        return False
    if isinstance(actual, list):
        return bool(actual) and all(values_match(item, expected) for item in actual)
    if isinstance(expected, float) and isinstance(actual, (int, float)):
        return abs(actual - expected) < 0.001
    return actual == expected

def execute_property_changes(
    schema: WordSchema, changes: list[dict], available_fields: dict[str, dict]
) -> dict:
    outcomes = []
    progress = tqdm(
        changes,
        desc="Applying Word properties",
        unit="change",
        leave=False,
        disable=not changes,
    )
    for change in progress:
        path = change.get("path")
        progress.set_postfix_str(str(path).rsplit(".", 1)[-1][:24])
        if path not in available_fields:
            outcomes.append(
                {"path": path, "success": False, "error": "Path is not in the active schema."}
            )
            continue
        try:
            property_path = PropertyPath(schema, path)
            before = property_path.read()
            result = property_path.set(change.get("value"))
            if not result["success"]:
                outcomes.append({"path": path, **result, "before": before})
                continue
            after = property_path.read()
            expected = property_path.normalized_expected(change.get("value"))
            verified = values_match(after, expected)
            outcomes.append(
                {
                    "path": path,
                    "success": verified,
                    "verified": verified,
                    "before": before,
                    "after": after,
                    "expected": expected,
                    "error": None if verified else "Post-change value did not match.",
                }
            )
        except Exception as exc:
            outcomes.append({"path": path, "success": False, "error": str(exc)})
    return {
        "success": bool(outcomes) and all(item["success"] for item in outcomes),
        "changes": outcomes,
    }

def compact_property_state(property_state: dict) -> dict:
    compact = {}
    for path, value in property_state.items():
        if not isinstance(value, list) or len(value) <= 20:
            compact[path] = value
            continue
        counts = {}
        originals = {}
        for item in value:
            key = json.dumps(item, sort_keys=True, ensure_ascii=False)
            counts[key] = counts.get(key, 0) + 1
            originals[key] = item
        compact[path] = {
            "item_count": len(value),
            "distinct_values": [
                {"value": originals[key], "count": count}
                for key, count in sorted(
                    counts.items(), key=lambda pair: pair[1], reverse=True
                )[:12]
            ],
        }
    return compact
