import re

from .applescript import run_applescript
from .utils import applescript_string, normalize_name
from .word_schema import WordSchema

class PropertyPath:
    """Validates and compiles canonical property paths into AppleScript."""

    SEGMENT_PATTERN = re.compile(r"^([a-z0-9_]+)(?:\[(\*|\d+)\])?$")

    def __init__(self, schema: WordSchema, path: str):
        self.schema = schema
        self.path = path
        self.owner_class = None
        self.property = None
        self.object_expression = None
        self.wildcard = None
        self._resolve()

    def _resolve(self):
        segments = self.path.split(".")
        if len(segments) < 2 or segments[0] not in {"document", "selection"}:
            raise ValueError(f"Invalid property root in path: {self.path}")

        current_class = "document" if segments[0] == "document" else "selection object"
        expression = "document 1" if segments[0] == "document" else "selection"

        for position, raw_segment in enumerate(segments[1:-1], start=1):
            match = self.SEGMENT_PATTERN.match(raw_segment)
            if not match:
                raise ValueError(f"Invalid path segment: {raw_segment}")
            name, index = match.groups()

            element_type = self.schema.get_element(current_class, name)
            if element_type:
                if not index:
                    raise ValueError(f"Collection path requires an index: {raw_segment}")
                if index == "*":
                    if self.wildcard:
                        raise ValueError("Only one wildcard collection is supported per path.")
                    self.wildcard = {
                        "source": expression,
                        "element_type": element_type,
                        "collection": self.schema.classes[element_type]["plural"],
                    }
                    expression = "targetItem"
                else:
                    expression = f"{element_type} {int(index)} of ({expression})"
                current_class = element_type
                continue

            prop = self.schema.get_property(current_class, name)
            if not prop or prop["type"] not in self.schema.classes:
                raise ValueError(
                    f"{raw_segment} is not an object property of {current_class}."
                )
            expression = f"{self._property_token(prop)} of ({expression})"
            current_class = prop["type"]

        final_name = normalize_name(segments[-1])
        prop = self.schema.get_property(current_class, final_name)
        if not prop:
            raise ValueError(f"{segments[-1]} is not a property of {current_class}.")
        if prop["type"] in self.schema.classes:
            raise ValueError(f"{self.path} points to an object, not a scalar property.")

        self.owner_class = current_class
        self.property = prop
        self.object_expression = expression

    @staticmethod
    def _property_token(prop: dict) -> str:
        code = prop.get("code")
        return f"«class {code}»" if code and len(code) == 4 else prop["name"]

    def _value_literal(self, value) -> str:
        prop_type = self.property["type"]
        if value is None:
            return "missing value"
        if self.property["list"]:
            if not isinstance(value, list):
                raise ValueError(f"{self.path} requires a list value.")
            return "{" + ", ".join(self._primitive_literal(item, prop_type) for item in value) + "}"
        return self._primitive_literal(value, prop_type)

    def _primitive_literal(self, value, prop_type: str) -> str:
        enum_values = self.schema.enum_values(prop_type)
        if enum_values:
            requested = normalize_name(str(value))
            match = next(
                (candidate for candidate in enum_values if normalize_name(candidate) == requested),
                None,
            )
            if not match:
                raise ValueError(
                    f"Invalid {prop_type} value {value!r}. Allowed: {enum_values}"
                )
            enum_code = self.schema.enumeration_codes.get(prop_type, {}).get(match)
            return str(enum_code) if enum_code is not None else match
        if prop_type == "boolean":
            if not isinstance(value, bool):
                raise ValueError(f"{self.path} requires true or false.")
            return "true" if value else "false"
        if prop_type == "integer":
            if isinstance(value, bool):
                raise ValueError(f"{self.path} requires an integer.")
            return str(int(value))
        if prop_type in {"real", "double"}:
            if isinstance(value, bool):
                raise ValueError(f"{self.path} requires a number.")
            return str(float(value))
        if prop_type in {"text", "string"}:
            return applescript_string(str(value))
        raise ValueError(
            f"Setting Word type {prop_type!r} is not supported safely yet."
        )

    def read(self):
        value_expression = (
            f"{self._property_token(self.property)} of ({self.object_expression})"
        )
        body = f"set end of outputItems to my scalarText({value_expression})"
        if self.wildcard:
            body = f"""
set targetCount to count {self.wildcard['collection']} of ({self.wildcard['source']})
repeat with targetIndex from 1 to targetCount
    try
        set targetItem to {self.wildcard['element_type']} (targetIndex as integer) of ({self.wildcard['source']})
        {body}
    on error errorMessage
        set end of outputItems to "__ERROR__" & errorMessage
    end try
end repeat
"""
        else:
            body = f"""
try
    {body}
on error errorMessage
    set end of outputItems to "__ERROR__" & errorMessage
end try
"""

        script = f"""
on scalarText(valueObject)
    if valueObject is missing value then return "__MISSING__"
    if class of valueObject is list then
        set oldDelimiters to AppleScript's text item delimiters
        set AppleScript's text item delimiters to ","
        set rendered to valueObject as text
        set AppleScript's text item delimiters to oldDelimiters
        return rendered
    end if
    return valueObject as text
end scalarText

set outputItems to {{}}
tell application "Microsoft Word"
    {body}
end tell
set oldDelimiters to AppleScript's text item delimiters
set AppleScript's text item delimiters to "|||UIVLM_ITEM|||"
set renderedOutput to outputItems as text
set AppleScript's text item delimiters to oldDelimiters
return renderedOutput
"""
        stdout, stderr, code = run_applescript(script)
        if code != 0:
            raise RuntimeError(stderr or "Word property read failed.")
        raw_values = stdout.split("|||UIVLM_ITEM|||") if stdout else []
        values = [self._coerce_read_value(value) for value in raw_values]
        return values if self.wildcard else (values[0] if values else None)

    def set(self, value) -> dict:
        if not self.property["writable"]:
            raise ValueError(f"{self.path} is read-only.")
        literal = self._value_literal(value)
        statement = (
            f"set {self._property_token(self.property)} "
            f"of ({self.object_expression}) to {literal}"
        )
        if self.wildcard:
            statement = f"""
set targetCount to count {self.wildcard['collection']} of ({self.wildcard['source']})
repeat with targetIndex from 1 to targetCount
    set targetItem to {self.wildcard['element_type']} (targetIndex as integer) of ({self.wildcard['source']})
    {statement}
end repeat
"""
        script = f'tell application "Microsoft Word"\n{statement}\nend tell'
        stdout, stderr, code = run_applescript(script)
        return {
            "success": code == 0,
            "error": stderr if code != 0 else None,
            "stdout": stdout,
        }

    def normalized_expected(self, value):
        if self.schema.enum_values(self.property["type"]):
            return normalize_name(str(value))
        if self.property["type"] == "integer":
            return int(value)
        if self.property["type"] in {"real", "double"}:
            return float(value)
        return value

    def _coerce_read_value(self, value: str):
        if value == "__MISSING__":
            return None
        if value.startswith("__ERROR__"):
            return {"error": value[len("__ERROR__") :]}
        if self.property["list"]:
            if not value:
                return []
            return [self._coerce_primitive(item.strip()) for item in value.split(",")]
        return self._coerce_primitive(value)

    def _coerce_primitive(self, value: str):
        prop_type = self.property["type"]
        if self.schema.enum_values(prop_type):
            if re.fullmatch(r"-?\d+", value):
                numeric_value = int(value)
                enum_codes = self.schema.enumeration_codes.get(prop_type, {})
                enum_name = next(
                    (
                        name
                        for name, code in enum_codes.items()
                        if code == numeric_value
                    ),
                    None,
                )
                return normalize_name(enum_name) if enum_name else numeric_value
            return normalize_name(value)
        if prop_type == "boolean":
            return value.lower() == "true"
        if prop_type == "integer":
            return int(float(value))
        if prop_type in {"real", "double"}:
            return float(value)
        return value
