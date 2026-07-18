from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from .config import WORD_SDEF_PATHS


def resolve_sdef_path(explicit_path: Optional[str] = None) -> Path:
    if explicit_path:
        candidate = Path(explicit_path).expanduser()
        if candidate.exists():
            return candidate
        raise FileNotFoundError(f"Word sdef not found: {candidate}")

    for candidate in WORD_SDEF_PATHS:
        candidate_path = Path(candidate).expanduser()
        if candidate_path.exists():
            return candidate_path

    raise FileNotFoundError("Could not locate Microsoft Word.sdef in the configured search paths")


@lru_cache(maxsize=4)
def load_sdef_tree(explicit_path: Optional[str] = None) -> ET.ElementTree:
    return ET.parse(resolve_sdef_path(explicit_path))


def _text(value: Optional[str]) -> str:
    return value.strip() if value else ""


def _iter_children(element: ET.Element, tag: str) -> Iterable[ET.Element]:
    return element.findall(f".//{{*}}{tag}")


def list_commands(explicit_path: Optional[str] = None) -> List[Dict[str, Any]]:
    tree = load_sdef_tree(explicit_path)
    commands: List[Dict[str, Any]] = []

    for command in _iter_children(tree.getroot(), "command"):
        commands.append(
            {
                "name": _text(command.get("name")),
                "code": _text(command.get("code")),
                "description": _text(command.get("description")),
                "parameters": [
                    {
                        "name": _text(parameter.get("name")),
                        "code": _text(parameter.get("code")),
                        "type": _text(parameter.get("type")),
                        "optional": parameter.get("optional", "no").lower() == "yes",
                        "description": _text(parameter.get("description")),
                    }
                    for parameter in command.findall("./{*}parameter")
                ],
            }
        )

    return commands


def list_classes(explicit_path: Optional[str] = None) -> List[Dict[str, Any]]:
    tree = load_sdef_tree(explicit_path)
    classes: List[Dict[str, Any]] = []

    for class_element in _iter_children(tree.getroot(), "class"):
        classes.append(
            {
                "name": _text(class_element.get("name")),
                "code": _text(class_element.get("code")),
                "plural": _text(class_element.get("plural")),
                "inherits": _text(class_element.get("inherits")),
                "description": _text(class_element.get("description")),
                "properties": [
                    {
                        "name": _text(property_element.get("name")),
                        "code": _text(property_element.get("code")),
                        "type": _text(property_element.get("type")),
                        "access": _text(property_element.get("access")),
                        "description": _text(property_element.get("description")),
                    }
                    for property_element in class_element.findall("./{*}property")
                ],
                "elements": [
                    _text(element.get("type"))
                    for element in class_element.findall("./{*}element")
                    if _text(element.get("type"))
                ],
            }
        )

    return classes


def list_enumerations(explicit_path: Optional[str] = None) -> List[Dict[str, Any]]:
    tree = load_sdef_tree(explicit_path)
    enumerations: List[Dict[str, Any]] = []

    for enumeration in _iter_children(tree.getroot(), "enumeration"):
        enumerations.append(
            {
                "name": _text(enumeration.get("name")),
                "code": _text(enumeration.get("code")),
                "enumerators": [
                    {
                        "name": _text(enumerator.get("name")),
                        "code": _text(enumerator.get("code")),
                    }
                    for enumerator in enumeration.findall("./{*}enumerator")
                ],
            }
        )

    return enumerations


def summarize_catalog(explicit_path: Optional[str] = None) -> Dict[str, Any]:
    commands = list_commands(explicit_path)
    classes = list_classes(explicit_path)
    enumerations = list_enumerations(explicit_path)
    return {
        "sdef_path": str(resolve_sdef_path(explicit_path)),
        "command_count": len(commands),
        "class_count": len(classes),
        "enumeration_count": len(enumerations),
        "command_names": [command["name"] for command in commands],
        "commands": commands,
        "classes": classes,
        "enumerations": enumerations,
    }


def find_command(name: str, explicit_path: Optional[str] = None) -> Optional[Dict[str, Any]]:
    target = name.strip().lower()
    for command in list_commands(explicit_path):
        if command["name"].lower() == target:
            return command
    return None


def _tokenize(text: str) -> List[str]:
    return [token for token in re.findall(r"[a-z0-9]+", text.lower()) if len(token) > 2]


def _score_text(text: str, tokens: List[str]) -> int:
    lowered = text.lower()
    score = 0
    for token in tokens:
        if token in lowered:
            score += 1
    return score


def search_catalog(query: str, explicit_path: Optional[str] = None, max_results: int = 20) -> Dict[str, List[Dict[str, Any]]]:
    tokens = _tokenize(query)
    if not tokens:
        return {"commands": [], "classes": [], "enumerations": []}

    commands = list_commands(explicit_path)
    classes = list_classes(explicit_path)
    enumerations = list_enumerations(explicit_path)

    scored_commands = []
    for command in commands:
        haystack = " ".join([command["name"], command["description"], " ".join(parameter["name"] for parameter in command["parameters"])])
        score = _score_text(haystack, tokens)
        if score:
            if any(token in command["name"].lower() for token in tokens):
                score += 5
            scored_commands.append((score, command))

    scored_classes = []
    for class_item in classes:
        haystack = " ".join(
            [
                class_item["name"],
                class_item["description"],
                " ".join(property_item["name"] for property_item in class_item["properties"]),
                " ".join(class_item["elements"]),
            ]
        )
        score = _score_text(haystack, tokens)
        if score:
            if any(token in class_item["name"].lower() for token in tokens):
                score += 5
            scored_classes.append((score, class_item))

    scored_enumerations = []
    for enumeration in enumerations:
        haystack = " ".join([enumeration["name"], " ".join(item["name"] for item in enumeration["enumerators"])])
        score = _score_text(haystack, tokens)
        if score:
            if any(token in enumeration["name"].lower() for token in tokens):
                score += 5
            scored_enumerations.append((score, enumeration))

    scored_commands.sort(key=lambda item: (-item[0], item[1]["name"]))
    scored_classes.sort(key=lambda item: (-item[0], item[1]["name"]))
    scored_enumerations.sort(key=lambda item: (-item[0], item[1]["name"]))

    def _dedupe(items: List[Dict[str, Any]], key_name: str = "name") -> List[Dict[str, Any]]:
        seen = set()
        deduped: List[Dict[str, Any]] = []
        for item in items:
            key = item.get(key_name, "")
            if key and key not in seen:
                seen.add(key)
                deduped.append(item)
        return deduped

    return {
        "commands": _dedupe([item[1] for item in scored_commands])[:max_results],
        "classes": _dedupe([item[1] for item in scored_classes])[:max_results],
        "enumerations": _dedupe([item[1] for item in scored_enumerations])[:max_results],
    }