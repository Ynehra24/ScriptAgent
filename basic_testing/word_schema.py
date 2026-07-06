import re
import xml.etree.ElementTree as ET
from collections import deque

from .config import WORD_SDEF_PATHS
from .utils import normalize_name

class WordSchema:
    """Parses Word's installed AppleScript dictionary into a searchable schema."""

    def __init__(self):
        self.path = next((path for path in WORD_SDEF_PATHS if path.exists()), None)
        if not self.path:
            raise FileNotFoundError("Microsoft Word's Word.sdef dictionary was not found.")

        root = ET.parse(self.path).getroot()
        self.classes = {}
        self.enumerations = {}
        self.enumeration_codes = {}
        self.commands = {}

        for node in root.iter("enumeration"):
            name = node.get("name")
            if name:
                enumerators = [
                    item for item in node.findall("enumerator") if item.get("name")
                ]
                self.enumerations[name] = [item.get("name") for item in enumerators]
                self.enumeration_codes[name] = {
                    item.get("name"): int(item.get("code"), 16) & 0xFFFF
                    for item in enumerators
                    if item.get("code", "").startswith("0x")
                }

        for node in root.iter("class"):
            name = node.get("name")
            if not name:
                continue
            properties = []
            for prop in node.findall("property"):
                prop_type = prop.get("type")
                type_node = prop.find("type")
                is_list = False
                if type_node is not None:
                    prop_type = prop_type or type_node.get("type")
                    is_list = type_node.get("list") == "yes"
                properties.append(
                    {
                        "name": prop.get("name"),
                        "code": prop.get("code"),
                        "type": prop_type or "unknown",
                        "list": is_list,
                        "writable": prop.get("access") != "r",
                        "description": prop.get("description", ""),
                    }
                )
            self.classes[name] = {
                "name": name,
                "plural": node.get("plural", f"{name}s"),
                "inherits": node.get("inherits"),
                "description": node.get("description", ""),
                "properties": properties,
                "elements": [
                    child.get("type")
                    for child in node.findall("element")
                    if child.get("type")
                ],
            }

        for node in root.iter("command"):
            name = node.get("name")
            if not name:
                continue
            self.commands[name] = {
                "name": name,
                "description": node.get("description", ""),
                "parameters": [
                    {
                        "name": parameter.get("name"),
                        "type": parameter.get("type", "unknown"),
                        "optional": parameter.get("optional") == "yes",
                    }
                    for parameter in node.findall("parameter")
                ],
            }

    def class_index(self) -> list[dict]:
        return [
            {
                "name": item["name"],
                "description": item["description"][:180],
                "property_count": len(item["properties"]),
                "element_count": len(item["elements"]),
            }
            for item in self.classes.values()
        ]

    def command_index(self) -> list[dict]:
        return [
            {"name": item["name"], "description": item["description"][:160]}
            for item in self.commands.values()
        ]

    def inherited_properties(self, class_name: str) -> list[dict]:
        result = []
        seen = set()
        current = class_name
        while current and current not in seen and current in self.classes:
            seen.add(current)
            result.extend(self.classes[current]["properties"])
            current = self.classes[current]["inherits"]
        unique = {}
        for prop in result:
            unique[prop["name"]] = prop
        return list(unique.values())

    def get_property(self, class_name: str, normalized: str) -> dict | None:
        return next(
            (
                prop
                for prop in self.inherited_properties(class_name)
                if normalize_name(prop["name"]) == normalized
            ),
            None,
        )

    def get_element(self, class_name: str, normalized: str) -> str | None:
        current = class_name
        seen = set()
        while current and current not in seen and current in self.classes:
            seen.add(current)
            for element_type in self.classes[current]["elements"]:
                element_class = self.classes.get(element_type, {})
                candidates = {
                    normalize_name(element_type),
                    normalize_name(element_class.get("plural", f"{element_type}s")),
                }
                if normalized in candidates:
                    return element_type
            current = self.classes[current]["inherits"]
        return None

    def enum_values(self, type_name: str) -> list[str]:
        return self.enumerations.get(type_name, [])

    def reachable_field_schema(
        self, selected_classes: list[str], max_depth: int = 5
    ) -> dict[str, dict]:
        selected = {name for name in selected_classes if name in self.classes}
        routes = self._routes_from_root("document", "document", max_depth)
        selection_routes = self._routes_from_root(
            "selection object", "selection", max_depth
        )
        for class_name, class_routes in selection_routes.items():
            routes.setdefault(class_name, [])
            for route in class_routes:
                if route not in routes[class_name] and len(routes[class_name]) < 6:
                    routes[class_name].append(route)

        fields = {}
        for class_name in selected:
            for route in routes.get(class_name, [])[:3]:
                for prop in self.inherited_properties(class_name):
                    if prop["type"] in self.classes:
                        continue
                    canonical = f"{route}.{normalize_name(prop['name'])}"
                    fields[canonical] = {
                        "path": canonical,
                        "owner_class": class_name,
                        "word_name": prop["name"],
                        "type": prop["type"],
                        "list": prop["list"],
                        "writable": prop["writable"],
                        "enum_values": self.enum_values(prop["type"]),
                        "description": prop["description"],
                    }
        return fields

    def _routes_from_root(
        self, root_class: str, root_path: str, max_depth: int
    ) -> dict[str, list[str]]:
        routes = {}
        queue = deque([(root_class, root_path, 0, (root_class,))])
        while queue:
            class_name, path, depth, ancestry = queue.popleft()
            routes.setdefault(class_name, [])
            if path not in routes[class_name] and len(routes[class_name]) < 3:
                routes[class_name].append(path)
            if depth >= max_depth or class_name not in self.classes:
                continue

            class_info = self.classes[class_name]
            for element_type in class_info["elements"]:
                if element_type in ancestry or element_type not in self.classes:
                    continue
                plural = normalize_name(self.classes[element_type]["plural"])
                queue.append(
                    (
                        element_type,
                        f"{path}.{plural}[*]",
                        depth + 1,
                        ancestry + (element_type,),
                    )
                )

            for prop in self.inherited_properties(class_name):
                target_type = prop["type"]
                if target_type in ancestry or target_type not in self.classes:
                    continue
                queue.append(
                    (
                        target_type,
                        f"{path}.{normalize_name(prop['name'])}",
                        depth + 1,
                        ancestry + (target_type,),
                    )
                )
        return routes
