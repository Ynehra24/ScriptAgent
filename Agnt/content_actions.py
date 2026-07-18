from __future__ import annotations

from typing import Dict

from .applescript import run_applescript_safe


def insert_text(params: Dict) -> Dict:
    text = params.get("text", "").replace('"', '\\"')
    position = params.get("position", "cursor")

    if position == "start":
        script = f'''
tell application "Microsoft Word"
    tell document 1
        insert text "{text}" at (create range start 0 end 0)
    end tell
end tell
'''
    elif position == "end":
        script = f'''
tell application "Microsoft Word"
    tell document 1
        insert text "{text}" at end of text object
    end tell
end tell
'''
    else:
        script = f'''
tell application "Microsoft Word"
    tell selection
        type text text "{text}"
    end tell
end tell
'''

    stdout, stderr, code = run_applescript_safe(script)
    return {"success": code == 0, "error": stderr if code != 0 else None, "stdout": stdout if code == 0 else None}