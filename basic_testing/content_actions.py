import json

from .applescript import run_applescript
from .utils import applescript_string

def get_content_state() -> dict:
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
        set output to "{" & quote & "name" & quote & ":" & quote & my escapeJson(name) & quote
        set output to output & "," & quote & "saved" & quote & ":" & (saved as text)
        set output to output & "," & quote & "word_count" & quote & ":" & ((count words) as text)
        set output to output & "," & quote & "paragraph_count" & quote & ":" & paragraphCount
        set output to output & "," & quote & "selection_text" & quote & ":" & quote & my escapeJson(selectedText) & quote
        set output to output & "," & quote & "paragraphs" & quote & ":["
        set maximumParagraphs to paragraphCount
        if maximumParagraphs > 100 then set maximumParagraphs to 100
        repeat with paragraphIndex from 1 to maximumParagraphs
            set paragraphText to content of text object of paragraph paragraphIndex
            if length of paragraphText > 300 then set paragraphText to text 1 thru 300 of paragraphText
            set output to output & "{" & quote & "index" & quote & ":" & paragraphIndex
            set output to output & "," & quote & "text" & quote & ":" & quote & my escapeJson(paragraphText) & quote & "}"
            if paragraphIndex < maximumParagraphs then set output to output & ","
        end repeat
        set output to output & "]}"
        return output
    end tell
end tell
'''
    stdout, stderr, code = run_applescript(script, timeout=30)
    if code != 0:
        return {"error": stderr or "Unable to read Word document state."}
    try:
        return json.loads(stdout)
    except json.JSONDecodeError as exc:
        return {"error": f"Invalid state returned by Word: {exc}"}

def execute_content_action(action: dict) -> dict:
    action_type = action.get("action_type")
    params = action.get("params", {})

    if action_type == "insert_text":
        text = applescript_string(params.get("text", ""))
        position = params.get("position", "cursor")
        lines = ['tell application "Microsoft Word"']
        if position == "start":
            lines.extend(
                [
                    "set «class 2903» of selection to 0",
                    "set «class 2905» of selection to 0",
                ]
            )
        elif position == "end":
            lines.extend(
                [
                    "set insertionPoint to «class 1656» of text object of document 1",
                    "set «class 2903» of selection to insertionPoint",
                    "set «class 2905» of selection to insertionPoint",
                ]
            )
        lines.extend(
            [
                "set insertionStart to «class 2903» of selection",
                f"«event sWRD3023» selection given «class 5418»:{text}",
                "set insertionEnd to «class 2903» of selection",
            ]
        )

        font_values = {
            "font_name": ("«class pnam»", params.get("font_name")),
            "font_size": ("«class ptsz»", params.get("font_size")),
            "bold": ("«class bold»", params.get("bold")),
            "italic": ("«class ital»", params.get("italic")),
            "font_color_rgb": ("«class colr»", params.get("font_color_rgb")),
        }
        requested_font_values = {
            key: value for key, value in font_values.items() if value[1] is not None
        }
        if requested_font_values:
            lines.extend(
                [
                    "set «class 2903» of selection to insertionStart",
                    "set «class 2905» of selection to insertionEnd",
                    "set insertedFont to «class wFnO» of selection",
                ]
            )
            for key, (word_property, value) in requested_font_values.items():
                if key == "font_color_rgb":
                    if (
                        not isinstance(value, list)
                        or len(value) != 3
                        or any(
                            isinstance(component, bool)
                            or not isinstance(component, int)
                            or not 0 <= component <= 65535
                            for component in value
                        )
                    ):
                        return {
                            "success": False,
                            "error": "font_color_rgb must contain three integers from 0 to 65535.",
                        }
                    literal = "{" + ", ".join(str(component) for component in value) + "}"
                elif isinstance(value, bool):
                    literal = "true" if value else "false"
                elif isinstance(value, (int, float)):
                    literal = str(value)
                else:
                    literal = applescript_string(value)
                lines.append(f"set {word_property} of insertedFont to {literal}")
            lines.extend(
                [
                    "set «class 2903» of selection to insertionEnd",
                    "set «class 2905» of selection to insertionEnd",
                ]
            )
        lines.append("end tell")
        script = "\n".join(lines)
    elif action_type == "replace_text":
        find_text = applescript_string(params.get("find", ""))
        replacement = applescript_string(params.get("replace", ""))
        script = f'''
tell application "Microsoft Word"
    set targetRange to text object of document 1
    set targetFind to «class 1717» of targetRange
    «event sWRD1874» targetFind given «class 5632»:{find_text}, «class 5641»:{replacement}, «class 5642»:2
end tell
'''
    elif action_type == "format_matches":
        lines = ['tell application "Microsoft Word"']
        for match in params.get("matches", []):
            find_text = match.get("find")
            if not find_text:
                continue
            lines.extend(
                [
                    "set targetRange to text object of document 1",
                    "set targetFind to «class 1717» of targetRange",
                    "set targetReplacement to «class w125» of targetFind",
                    "«event sWRDwClf» targetReplacement",
                    "set targetFont to «class wFnO» of targetReplacement",
                ]
            )
            for key, word_property in (
                ("bold", "«class bold»"),
                ("italic", "«class ital»"),
                ("font_size", "«class ptsz»"),
                ("font_name", "«class pnam»"),
                ("font_color_rgb", "«class colr»"),
            ):
                if key not in match:
                    continue
                value = match[key]
                if key == "font_color_rgb":
                    if (
                        not isinstance(value, list)
                        or len(value) != 3
                        or any(
                            isinstance(component, bool)
                            or not isinstance(component, int)
                            or not 0 <= component <= 65535
                            for component in value
                        )
                    ):
                        return {
                            "success": False,
                            "error": "font_color_rgb must contain three integers from 0 to 65535.",
                        }
                    literal = "{" + ", ".join(str(component) for component in value) + "}"
                elif isinstance(value, bool):
                    literal = "true" if value else "false"
                elif isinstance(value, (int, float)):
                    literal = str(value)
                else:
                    literal = applescript_string(value)
                lines.append(
                    f"set {word_property} of targetFont to {literal}"
                )
            literal_find = applescript_string(find_text)
            lines.append(
                f"«event sWRD1874» targetFind given "
                f"«class 5632»:{literal_find}, "
                f"«class 5641»:{literal_find}, «class 5642»:2"
            )
        lines.append("end tell")
        script = "\n".join(lines)
    elif action_type == "delete_paragraph":
        paragraph = int(params.get("paragraph", 1))
        script = (
            'tell application "Microsoft Word" to tell document 1 '
            f"to delete text object of paragraph {paragraph}"
        )
    elif action_type == "save_document":
        script = 'tell application "Microsoft Word" to save document 1'
    else:
        return {"success": False, "error": f"Unsupported action: {action_type}"}

    stdout, stderr, code = run_applescript(script)
    return {
        "success": code == 0,
        "error": stderr if code != 0 else None,
        "stdout": stdout,
    }
