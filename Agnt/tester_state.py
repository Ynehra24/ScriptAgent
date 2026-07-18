import json
import subprocess
import tempfile
import os
from typing import Dict

from .applescript import run_applescript_safe


def get_word_state_mac() -> Dict:
    """Read Word state via AppleScript, outputting structured JSON."""
    applescript = '''
on escape_json(txt)
    if txt is missing value then return ""
    return do shell script "/usr/bin/python3 -c " & quoted form of "import json, sys; print(json.dumps(sys.argv[1])[1:-1])" & " " & quoted form of txt
end escape_json

tell application "Microsoft Word"
    activate
    set sel_start to selection start of selection
    set sel_end to selection end of selection
    try
        set sel_text to content of selection
        if length of sel_text > 200 then
            set sel_text to text 1 thru 200 of sel_text
        end if
    on error
        set sel_text to ""
    end try

    tell document 1
        set doc_name to name
        set doc_saved to saved
        set para_count to count paragraphs
        set word_count to count words

        set output to "name:" & my escape_json(doc_name) & linefeed
        set output to output & "saved:" & (doc_saved as string) & linefeed
        set output to output & "word_count:" & (word_count as string) & linefeed
        set output to output & "paragraph_count:" & (para_count as string) & linefeed
        set output to output & "selection_start:" & (sel_start as string) & linefeed
        set output to output & "selection_end:" & (sel_end as string) & linefeed
        set output to output & "selection_text:" & my escape_json(sel_text) & linefeed
        repeat with i from 1 to para_count
            tell paragraph i
                set para_text to content of text object
            end tell
            set output to output & "paragraph:" & (i as string) & ":" & my escape_json(para_text) & linefeed
        end repeat
        return output
    end tell
end tell
'''

    # Run via temporary file using the shared applescript helper
    with tempfile.NamedTemporaryFile(mode='w', suffix='.applescript', delete=False) as f:
        f.write(applescript)
        tmp = f.name

    try:
        stdout, stderr, code = run_applescript_safe(open(tmp).read(), timeout=30, retries=1)
        if code != 0:
            return {"error": f"AppleScript failed: {stderr}"}
        data = {
            "name": "",
            "saved": False,
            "word_count": 0,
            "paragraph_count": 0,
            "selection_start": 0,
            "selection_end": 0,
            "selection_text": "",
            "paragraphs": [],
        }

        for raw_line in stdout.splitlines():
            if not raw_line:
                continue
            if raw_line.startswith("paragraph:"):
                _, index_text, encoded_text = raw_line.split(":", 2)
                paragraph_index = int(index_text)
                paragraph_text = json.loads(f'"{encoded_text}"')
                while len(data["paragraphs"]) < paragraph_index:
                    data["paragraphs"].append({"text": ""})
                data["paragraphs"][paragraph_index - 1] = {"text": paragraph_text}
                continue

            key, value = raw_line.split(":", 1)
            if key in ("saved",):
                data[key] = value.lower() == "true"
            elif key in ("word_count", "paragraph_count", "selection_start", "selection_end"):
                data[key] = int(value)
            elif key == "selection_text":
                data[key] = json.loads(f'"{value}"')
            elif key == "name":
                data[key] = json.loads(f'"{value}"')

        return data
    except json.JSONDecodeError as e:
        return {"error": f"JSON parse error: {str(e)}\nRaw output: {stdout[:200]}"}
    except Exception as e:
        return {"error": f"Exception: {str(e)}"}
    finally:
        try:
            os.unlink(tmp)
        except Exception:
            pass
