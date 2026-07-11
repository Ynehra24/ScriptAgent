"""
July 9 2026 -> Fixes
1. Major bottleneck found within the limiting of returned metadata. We return the first 300 chars
sure but there is no check to ensure it is the exact 300 needed by the system itself. This needs to
be fixed by replacing the current "always send the first 300 chars" to a system which will run this 
function multiple times if needed to surgically get the piece of metadata it needs.
2. Maybe add temp fields to store local and global attributes pertaining to paragraph and document
levels.
3. Fixed and hardcoded action vocabulary -> think in terms of word objects or use the sdef
"""
import json

from .applescript import run_applescript, run_applescript_safe
from .utils import applescript_string


def _word_bool(value) -> str:
    return "true" if bool(value) else "false"


def _word_text_literal(value: str) -> str:
    parts = []
    text = str(value)
    buffer = []
    for character in text:
        if character == "\t":
            if buffer:
                parts.append(applescript_string("".join(buffer)))
                buffer = []
            parts.append("tab")
        elif character in {"\r", "\n"}:
            if buffer:
                parts.append(applescript_string("".join(buffer)))
                buffer = []
            parts.append("return")
        else:
            buffer.append(character)
    if buffer:
        parts.append(applescript_string("".join(buffer)))
    return " & ".join(parts) if parts else '""'


def _word_rgb_literal(value) -> str:
    named_colors = {
        "maroon": [32768, 0, 0],
        "grey": [32768, 32768, 32768],
        "gray": [32768, 32768, 32768],
        "red": [65535, 0, 0],
        "green": [0, 65535, 0],
        "blue": [0, 0, 65535],
        "white": [65535, 65535, 65535],
        "black": [0, 0, 0],
    }
    if isinstance(value, str):
        value = named_colors.get(value.strip().lower())
    if (
        not isinstance(value, list)
        or len(value) != 3
        or any(isinstance(item, bool) or not isinstance(item, int) for item in value)
    ):
        raise ValueError("Color must be a known name or an RGB list of three integers.")
    return "{" + ", ".join(str(max(0, min(65535, item))) for item in value) + "}"


def _table_shading_script(rows: int, columns: int, arguments: dict) -> str:
    alternating = arguments.get("alternating_cell_colors")
    cell_colors = arguments.get("cell_colors")
    statements = []

    for row_index in range(1, rows + 1):
        for column_index in range(1, columns + 1):
            color = None
            if cell_colors:
                color = cell_colors[row_index - 1][column_index - 1]
            elif alternating:
                color = alternating[(row_index + column_index) % len(alternating)]
            if color is None:
                continue
            color_literal = _word_rgb_literal(color)
            statements.append(
                "set «class 2383» of shading of "
                f"cell {column_index} of row {row_index} of targetTable to {color_literal}"
            )

    return "\n    ".join(statements)


def _table_cell_fill_script(values: list[list]) -> str:
    statements = []
    for row_index, row in enumerate(values, 1):
        for column_index, value in enumerate(row, 1):
            statements.append(
                f"set content of text object of cell {column_index} "
                f"of row {row_index} of targetTable to {_word_text_literal(str(value))}"
            )
    return "\n    ".join(statements)


def get_content_state(
    paragraph_start: int = 1,
    paragraph_limit: int = 100,
    preview_chars: int = 300,
) -> dict:
    paragraph_start = max(1, int(paragraph_start))
    paragraph_limit = max(1, int(paragraph_limit))
    preview_chars = max(0, int(preview_chars))

    script = r'''
on escapeJson(txt)
    if txt is missing value then return ""
    set txt to txt as text
    set oldDelimiters to AppleScript's text item delimiters
    set backslashCharacter to "\\"
    set AppleScript's text item delimiters to backslashCharacter
    set textItems to every text item of txt
    set AppleScript's text item delimiters to backslashCharacter & backslashCharacter
    set txt to textItems as text
    set AppleScript's text item delimiters to quote
    set textItems to every text item of txt
    set AppleScript's text item delimiters to backslashCharacter & quote
    set txt to textItems as text
    set AppleScript's text item delimiters to linefeed
    set textItems to every text item of txt
    set AppleScript's text item delimiters to backslashCharacter & "n"
    set txt to textItems as text
    set AppleScript's text item delimiters to return
    set textItems to every text item of txt
    set AppleScript's text item delimiters to backslashCharacter & "r"
    set txt to textItems as text
    set AppleScript's text item delimiters to tab
    set textItems to every text item of txt
    set AppleScript's text item delimiters to backslashCharacter & "t"
    set txt to textItems as text
    set AppleScript's text item delimiters to oldDelimiters
    return txt
end escapeJson

tell application "Microsoft Word"
    if (count documents) is 0 then error "No Word document is open."
    set selectedText to ""
    try
        set selectedText to content of selection
        if length of selectedText > 300 then set selectedText to text 1 thru 300 of selectedText
    end try
    tell document 1
        set paragraphCount to count paragraphs
        set requestedStart to PARAGRAPH_START
        set requestedLimit to PARAGRAPH_LIMIT
        set requestedPreviewChars to PREVIEW_CHARS
        set requestedEnd to requestedStart + requestedLimit - 1
        if requestedEnd > paragraphCount then set requestedEnd to paragraphCount
        set output to "{" & quote & "name" & quote & ":" & quote & my escapeJson(name) & quote
        set output to output & "," & quote & "saved" & quote & ":" & (saved as text)
        set output to output & "," & quote & "word_count" & quote & ":" & ((count words) as text)
        set output to output & "," & quote & "paragraph_count" & quote & ":" & paragraphCount
        set output to output & "," & quote & "selection_text" & quote & ":" & quote & my escapeJson(selectedText) & quote
        set output to output & "," & quote & "content_window" & quote & ":{" & quote & "start" & quote & ":" & requestedStart & "," & quote & "limit" & quote & ":" & requestedLimit & "," & quote & "preview_chars" & quote & ":" & requestedPreviewChars & "}"
        set output to output & "," & quote & "paragraphs" & quote & ":["
        repeat with paragraphIndex from requestedStart to requestedEnd
            set paragraphText to content of text object of paragraph paragraphIndex
            if requestedPreviewChars is 0 then
                set paragraphText to ""
            else if length of paragraphText > requestedPreviewChars then
                set paragraphText to text 1 thru requestedPreviewChars of paragraphText
            end if
            set styleName to ""
            set fontName to ""
            set fontSize to ""
            set isBold to false
            set isItalic to false
            try
                set styleName to name of style of paragraph paragraphIndex
            end try
            set output to output & "{" & quote & "index" & quote & ":" & paragraphIndex
            set output to output & "," & quote & "text" & quote & ":" & quote & my escapeJson(paragraphText) & quote
            set output to output & "," & quote & "style" & quote & ":" & quote & my escapeJson(styleName) & quote
            set output to output & "," & quote & "font_name" & quote & ":" & quote & my escapeJson(fontName) & quote
            set output to output & "," & quote & "font_size" & quote & ":" & quote & fontSize & quote
            set output to output & "," & quote & "bold" & quote & ":" & (isBold as text)
            set output to output & "," & quote & "italic" & quote & ":" & (isItalic as text) & "}"
            if paragraphIndex < requestedEnd then set output to output & ","
        end repeat
        set output to output & "]}"
        return output
    end tell
end tell
'''
    script = (
        script.replace("PARAGRAPH_START", str(paragraph_start))
        .replace("PARAGRAPH_LIMIT", str(paragraph_limit))
        .replace("PREVIEW_CHARS", str(preview_chars))
    )
    stdout, stderr, code = run_applescript_safe(script, timeout=30, retries=2)
    if code != 0:
        return {"error": stderr or "Unable to read Word document state."}
    try:
        return json.loads(stdout)
    except json.JSONDecodeError as exc:
        return {"error": f"Invalid state returned by Word: {exc}"}

def execute_context_request(retrieval: dict) -> dict:
    capability = retrieval.get("capability")
    arguments = retrieval.get("arguments", {})
    object_type = arguments.get("object", "paragraph")

    if object_type != "paragraph":
        return {
            "success": False,
            "error": f"Context retrieval for {object_type!r} is not implemented yet.",
        }

    if capability == "inspect_object_window":
        start = int(arguments.get("start", 1))
        limit = int(arguments.get("limit", 20))
        preview_chars = int(arguments.get("preview_chars", 600))
        return {
            "success": True,
            "context": get_content_state(start, limit, preview_chars),
        }

    if capability == "inspect_object_detail":
        index = int(arguments.get("index", 1))
        preview_chars = int(arguments.get("preview_chars", 2000))
        return {
            "success": True,
            "context": get_content_state(index, 1, preview_chars),
        }

    return {
        "success": False,
        "error": f"Unsupported context retrieval capability: {capability}",
    }


def insert_text(arguments: dict) -> dict:
    text = arguments.get("text", "")
    position = arguments.get("position", "end")
    font_name = arguments.get("font_name")
    font_size = arguments.get("font_size")
    bold = arguments.get("bold")
    italic = arguments.get("italic")

    if position not in {"start", "end", "selection"}:
        return {"success": False, "error": f"Unsupported insert position: {position}"}

    selection_setup = ""
    if position == "start":
        selection_setup = "select text object of paragraph 1 of document 1"
    elif position == "end":
        selection_setup = """
set insertionPoint to end of content of text object of document 1
set «class 2903» of selection to insertionPoint
set «class 2905» of selection to insertionPoint
"""
    else:
        selection_setup = ""

    formatting = []
    if font_name:
        formatting.append(f"set name of font object of selection to {applescript_string(font_name)}")
    if font_size is not None:
        formatting.append(f"set font size of font object of selection to {float(font_size)}")
    if bold is not None:
        formatting.append(f"set bold of font object of selection to {_word_bool(bold)}")
    if italic is not None:
        formatting.append(f"set italic of font object of selection to {_word_bool(italic)}")

    script = f'''
tell application "Microsoft Word"
    if (count documents) is 0 then error "No Word document is open."
    {selection_setup}
    set content of selection to {_word_text_literal(text)}
    {"\n    ".join(formatting)}
end tell
'''
    stdout, stderr, code = run_applescript_safe(script, retries=2)
    return {"success": code == 0, "error": stderr if code != 0 else None, "stdout": stdout}


def replace_paragraph(arguments: dict) -> dict:
    index = int(arguments.get("paragraph", arguments.get("index", 0)))
    text = arguments.get("text", "")
    if index < 1:
        return {"success": False, "error": "paragraph/index must be 1 or greater."}
    script = f'''
tell application "Microsoft Word"
    if (count documents) is 0 then error "No Word document is open."
    tell document 1
        if {index} > (count paragraphs) then error "Paragraph index is out of range."
        set content of text object of paragraph {index} to {_word_text_literal(text)}
    end tell
end tell
'''
    stdout, stderr, code = run_applescript_safe(script, retries=2)
    return {"success": code == 0, "error": stderr if code != 0 else None, "stdout": stdout}


def delete_paragraph(arguments: dict) -> dict:
    index = int(arguments.get("paragraph", arguments.get("index", 0)))
    if index < 1:
        return {"success": False, "error": "paragraph/index must be 1 or greater."}
    script = f'''
tell application "Microsoft Word"
    if (count documents) is 0 then error "No Word document is open."
    tell document 1
        if {index} > (count paragraphs) then error "Paragraph index is out of range."
        select text object of paragraph {index}
        delete selection
    end tell
end tell
'''
    stdout, stderr, code = run_applescript_safe(script, retries=2)
    if code != 0:
        fallback = replace_paragraph({"paragraph": index, "text": ""})
        if fallback.get("success"):
            fallback["note"] = (
                "Word did not accept deletion of the selected paragraph, "
                "so the paragraph content was cleared instead."
            )
            return fallback
    return {"success": code == 0, "error": stderr if code != 0 else None, "stdout": stdout}


def format_paragraph(arguments: dict) -> dict:
    index = int(arguments.get("paragraph", arguments.get("index", 0)))
    if index < 1:
        return {"success": False, "error": "paragraph/index must be 1 or greater."}

    statements = []
    if arguments.get("font_name"):
        statements.append(f"set name of font object of targetRange to {applescript_string(arguments['font_name'])}")
    if arguments.get("font_size") is not None:
        statements.append(f"set font size of font object of targetRange to {float(arguments['font_size'])}")
    if arguments.get("bold") is not None:
        statements.append(f"set bold of font object of targetRange to {_word_bool(arguments['bold'])}")
    if arguments.get("italic") is not None:
        statements.append(f"set italic of font object of targetRange to {_word_bool(arguments['italic'])}")
    if not statements:
        return {"success": False, "error": "No supported formatting arguments supplied."}

    script = f'''
tell application "Microsoft Word"
    if (count documents) is 0 then error "No Word document is open."
    tell document 1
        if {index} > (count paragraphs) then error "Paragraph index is out of range."
        set targetRange to text object of paragraph {index}
        {"\n        ".join(statements)}
    end tell
end tell
'''
    stdout, stderr, code = run_applescript_safe(script, retries=2)
    return {"success": code == 0, "error": stderr if code != 0 else None, "stdout": stdout}


def insert_table(arguments: dict) -> dict:
    rows = int(arguments.get("rows", 0))
    columns = int(arguments.get("columns", 0))
    values = arguments.get("values")
    paragraph = arguments.get("paragraph")
    if rows < 1 or columns < 1:
        return {"success": False, "error": "rows and columns must be 1 or greater."}
    if values is None:
        values = [["" for _ in range(columns)] for _ in range(rows)]
    if len(values) != rows or any(len(row) != columns for row in values):
        return {"success": False, "error": "values must match rows x columns."}
    if arguments.get("cell_colors"):
        cell_colors = arguments["cell_colors"]
        if len(cell_colors) != rows or any(len(row) != columns for row in cell_colors):
            return {"success": False, "error": "cell_colors must match rows x columns."}
    try:
        shading_script = _table_shading_script(rows, columns, arguments)
    except ValueError as exc:
        return {"success": False, "error": str(exc)}

    fill_script = _table_cell_fill_script(values)
    if paragraph:
        index = int(paragraph)
        range_setup = f"""
set targetRange to text object of paragraph {index} of document 1
set content of targetRange to ""
"""
    else:
        position = arguments.get("position", "end")
        if position == "start":
            range_setup = "set targetRange to text object of paragraph 1 of document 1"
        elif position == "selection":
            range_setup = "set targetRange to selection"
        else:
            range_setup = """
set insertionPoint to end of content of text object of document 1
set «class 2903» of selection to insertionPoint
set «class 2905» of selection to insertionPoint
set targetRange to selection
"""

    script = f'''
tell application "Microsoft Word"
    if (count documents) is 0 then error "No Word document is open."
    {range_setup}
    set targetTable to make new table at document 1 with properties {{text object:targetRange, number of rows:{rows}, number of columns:{columns}}}
    {fill_script}
    {shading_script}
end tell
'''
    stdout, stderr, code = run_applescript_safe(script, retries=2)
    if code == 0:
        return {"success": True, "error": None, "stdout": stdout}

    # Fallback: insert a plain-text table (tab-separated) so the user still gets the data
    try:
        lines = []
        for r in values:
            # convert each cell to string and separate with a tab
            safe_cells = [str(c) for c in r]
            lines.append("\t".join(safe_cells))
        fallback_text = "\n".join(lines)
        # Use the same insertion target as requested
        insert_args = {"text": fallback_text}
        if paragraph:
            insert_args["paragraph"] = paragraph
        else:
            insert_args["position"] = arguments.get("position", "end")

        fallback_result = insert_text(insert_args)
        note = (
            "Fell back to inserting a tab-separated text representation of the table."
        )
        # Merge fallback_result into response
        return {
            "success": bool(fallback_result.get("success")),
            "error": stderr if not fallback_result.get("success") else None,
            "stdout": stdout,
            "fallback": fallback_result,
            "note": note,
        }
    except Exception as exc:
        return {"success": False, "error": f"Insert table failed and fallback failed: {exc}", "stdout": stdout}
