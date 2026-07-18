import json
import tempfile
import os
from typing import Dict

from .applescript import run_applescript_safe
from .content_actions import insert_text


def execute_action_mac(action: Dict) -> Dict:
    atype = action.get("action_type")
    params = action.get("params", {})

    try:
        if atype == "apple_script":
            script = params.get("script", "")
            if not script.strip():
                return {"success": False, "error": "Missing AppleScript payload"}
            stdout, stderr, code = run_applescript_safe(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        if atype == "insert_text":
            # delegate to existing insert_text helper to keep behavior consistent
            return insert_text(params)

        # For other actions, we construct minimal scripts and use run_applescript_safe
        if atype == "select_paragraph":
            para_num = int(params.get("paragraph", 1))
            script = f'''
tell application "Microsoft Word"
    tell document 1
        select text object of paragraph {para_num}
    end tell
end tell
'''
            stdout, stderr, code = run_applescript_safe(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        if atype == "format_text":
            # Build script based on params
            para_num = params.get("paragraph")
            parts = []
            if params.get("bold") is not None:
                val = "true" if params["bold"] else "false"
                parts.append(f"set bold of font object to {val}")
            if params.get("italic") is not None:
                val = "true" if params["italic"] else "false"
                parts.append(f"set italic of font object to {val}")
            if params.get("font_size") is not None:
                parts.append(f"set font size of font object to {float(params['font_size'])}")
            if params.get("font_name") is not None:
                parts.append(f'set name of font object to "{params["font_name"]}"')

            if para_num is not None:
                script = f"""
tell application "Microsoft Word"
    tell document 1
        tell paragraph {int(para_num)}
            {"\n            ".join(parts)}
        end tell
    end tell
end tell
"""
            else:
                script = f"""
tell application "Microsoft Word"
    tell selection
        {"\n        ".join(parts)}
    end tell
end tell
"""
            stdout, stderr, code = run_applescript_safe(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        if atype == "save_file":
            script = 'tell application "Microsoft Word" to save document 1'
            stdout, stderr, code = run_applescript_safe(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        return {"success": False, "error": f"Unknown action: {atype}"}

    except Exception as e:
        return {"success": False, "error": str(e)}
