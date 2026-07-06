from pathlib import Path

from .applescript import run_applescript
from .utils import applescript_string

def open_new_document() -> bool:
    stdout, stderr, code = run_applescript(
        'tell application "Microsoft Word" to make new document'
    )
    if code == 0:
        print("\nNew document created.")
    else:
        print(f"\nFailed to create document: {stderr}")
    return code == 0

def open_existing_document(path: str) -> bool:
    expanded = Path(path.strip()).expanduser()
    if not expanded.exists():
        print(f"\nFile not found: {expanded}")
        return False
    script = (
        'tell application "Microsoft Word"\n'
        "activate\n"
        f"open {applescript_string(str(expanded.resolve()))}\n"
        "end tell"
    )
    stdout, stderr, code = run_applescript(script)
    if code == 0:
        print(f"\nOpened: {expanded.resolve()}")
    else:
        print(f"\nFailed to open document: {stderr}")
    return code == 0

def list_open_documents() -> list[str]:
    script = '''
tell application "Microsoft Word"
    if (count documents) is 0 then return ""
    set documentNames to name of every document
    set oldDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to linefeed
    set output to documentNames as text
    set AppleScript's text item delimiters to oldDelimiters
    return output
end tell
'''
    stdout, _, code = run_applescript(script)
    return stdout.splitlines() if code == 0 and stdout else []
