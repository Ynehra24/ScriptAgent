import os
import json
import subprocess
import tempfile
from dotenv import load_dotenv
from groq import Groq

load_dotenv()
client = Groq(api_key=os.environ.get("GROQ_API_KEY"))

# ─── STATE READER ────────────────────────────────────────────────────────────

def get_word_state_mac() -> dict:
    """Read Word state via AppleScript, outputting structured JSON."""
    applescript = '''
on escape_json(txt)
    if txt is missing value then return ""
    set oldDelims to AppleScript's text item delimiters

    set AppleScript's text item delimiters to "\\\\"
    set theItems to every text item of txt
    set AppleScript's text item delimiters to "\\\\\\\\"
    set txt to theItems as string

    set AppleScript's text item delimiters to "\\""
    set theItems to every text item of txt
    set AppleScript's text item delimiters to "\\\\\\""
    set txt to theItems as string

    set AppleScript's text item delimiters to linefeed
    set theItems to every text item of txt
    set AppleScript's text item delimiters to "\\\\n"
    set txt to theItems as string

    set AppleScript's text item delimiters to return
    set theItems to every text item of txt
    set AppleScript's text item delimiters to "\\\\r"
    set txt to theItems as string

    set AppleScript's text item delimiters to tab
    set theItems to every text item of txt
    set AppleScript's text item delimiters to "\\\\t"
    set txt to theItems as string

    set AppleScript's text item delimiters to oldDelims
    return txt
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

        set output to "{" & linefeed
        set output to output & "  \\"name\\": \\"" & my escape_json(doc_name) & "\\"," & linefeed
        set output to output & "  \\"saved\\": " & (doc_saved as string) & "," & linefeed
        set output to output & "  \\"word_count\\": " & (word_count as string) & "," & linefeed
        set output to output & "  \\"paragraph_count\\": " & (para_count as string) & "," & linefeed
        set output to output & "  \\"selection_start\\": " & (sel_start as string) & "," & linefeed
        set output to output & "  \\"selection_end\\": " & (sel_end as string) & "," & linefeed
        set output to output & "  \\"selection_text\\": \\"" & my escape_json(sel_text) & "\\"," & linefeed
        set output to output & "  \\"paragraphs\\": [" & linefeed
        repeat with i from 1 to para_count
            tell paragraph i
                set para_text to content of text object
                set para_style to name local of style
                set para_bold to (bold of font object of text object)
                if para_bold is missing value then
                    set para_bold_str to "null"
                else if para_bold is true then
                    set para_bold_str to "true"
                else
                    set para_bold_str to "false"
                end if
                set para_italic to (italic of font object of text object)
                if para_italic is missing value then
                    set para_italic_str to "null"
                else if para_italic is true then
                    set para_italic_str to "true"
                else
                    set para_italic_str to "false"
                end if
                set para_size to (font size of font object of text object)
                if para_size is missing value then
                    set para_size_str to "null"
                else
                    set para_size_str to para_size as string
                end if
            end tell
            set output to output & "    {" & linefeed
            set output to output & "      \\"text\\": \\"" & my escape_json(para_text) & "\\"," & linefeed
            set output to output & "      \\"style\\": \\"" & my escape_json(para_style) & "\\"," & linefeed
            set output to output & "      \\"bold\\": " & para_bold_str & "," & linefeed
            set output to output & "      \\"italic\\": " & para_italic_str & "," & linefeed
            set output to output & "      \\"font_size\\": " & para_size_str & linefeed
            if i < para_count then
                set output to output & "    }," & linefeed
            else
                set output to output & "    }" & linefeed
            end if
        end repeat
        set output to output & "  ]" & linefeed
        set output to output & "}"
        return output
    end tell
end tell
'''

    with tempfile.NamedTemporaryFile(mode='w', suffix='.applescript', delete=False) as f:
        f.write(applescript)
        tmp = f.name

    try:
        result = subprocess.run(
            ["osascript", tmp],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            return {"error": f"AppleScript failed: {result.stderr.strip()}"}

        data = json.loads(result.stdout.strip())
        return data

    except json.JSONDecodeError as e:
        return {"error": f"JSON parse error: {str(e)}\nRaw output: {result.stdout[:200]}"}
    except Exception as e:
        return {"error": f"Exception: {str(e)}"}
    finally:
        os.unlink(tmp)


# ─── ACTION EXECUTOR ─────────────────────────────────────────────────────────

def execute_action_mac(action: dict) -> dict:
    atype = action["action_type"]
    params = action.get("params", {})

    try:
        if atype == "insert_text":
            text = params.get("text", "").replace('"', '\\"')
            position = params.get("position", "cursor")

            if position == "start":
                # Insert at the very beginning of the document
                script = f'''
                tell application "Microsoft Word"
                    tell document 1
                        insert text "{text}" at (create range start 0 end 0)
                    end tell
                end tell
                '''
            elif position == "end":
                # Insert at the end of the document
                script = f'''
                tell application "Microsoft Word"
                    tell document 1
                        insert text "{text}" at end of text object
                    end tell
                end tell
                '''
            else:
                # Insert at cursor (selection)
                script = f'''
                tell application "Microsoft Word"
                    tell selection
                        type text text "{text}"
                    end tell
                end tell
                '''
            stdout, stderr, code = run_applescript(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        elif atype == "select_paragraph":
            para_num = params.get("paragraph", 1)
            script = f'''
            tell application "Microsoft Word"
                tell document 1
                    select text object of paragraph {para_num}
                end tell
            end tell
            '''
            stdout, stderr, code = run_applescript(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        elif atype == "format_text":
            para_num = params.get("paragraph", None)

            if para_num is not None:
                # Apply formatting directly to the specified paragraph
                parts = []
                if params.get("bold") is not None:
                    val = "true" if params["bold"] else "false"
                    parts.append(f"set bold of font object of text object to {val}")
                if params.get("font_size") is not None:
                    parts.append(f"set font size of font object of text object to {params['font_size']}")
                if params.get("italic") is not None:
                    val = "true" if params["italic"] else "false"
                    parts.append(f"set italic of font object of text object to {val}")
                if params.get("font_name") is not None:
                    parts.append(f'set name of font object of text object to "{params["font_name"]}"')

                script = f'''
                tell application "Microsoft Word"
                    tell document 1
                        tell paragraph {para_num}
                            {chr(10).join(parts)}
                        end tell
                    end tell
                end tell
                '''
            else:
                # Apply formatting to the current selection
                parts = []
                if params.get("bold") is not None:
                    val = "true" if params["bold"] else "false"
                    parts.append(f"set bold of font object to {val}")
                if params.get("font_size") is not None:
                    parts.append(f"set font size of font object to {params['font_size']}")
                if params.get("italic") is not None:
                    val = "true" if params["italic"] else "false"
                    parts.append(f"set italic of font object to {val}")
                if params.get("font_name") is not None:
                    parts.append(f'set name of font object to "{params["font_name"]}"')

                script = f'''
                tell application "Microsoft Word"
                    tell selection
                        {chr(10).join(parts)}
                    end tell
                end tell
                '''
            stdout, stderr, code = run_applescript(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        elif atype == "set_style":
            style = params.get("style", "Normal")
            script = f'''
            tell application "Microsoft Word"
                set style of selection to Word style "{style}" of active document
            end tell
            '''
            stdout, stderr, code = run_applescript(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        elif atype == "find_replace":
            find_text = params.get("find", "").replace('"', '\\"')
            replace_text = params.get("replace", "").replace('"', '\\"')
            script = f'''
            tell application "Microsoft Word"
                tell document 1
                    set myRange to create range start 0 end (end of content of text object)
                    tell find object of myRange
                        execute find find text "{find_text}" replace with "{replace_text}" replace replace all
                    end tell
                end tell
            end tell
            '''
            stdout, stderr, code = run_applescript(script)
        elif atype == "find_format":
            formats = params.get("formats", [])
            if not formats:
                formats = [params]

            script_parts = []
            script_parts.append('tell application "Microsoft Word"')
            script_parts.append('    tell document 1')

            for fmt in formats:
                find_text = fmt.get("find", "").replace('"', '\\"')
                if not find_text:
                    continue
                script_parts.append('        set myRange to create range start 0 end (end of content of text object)')
                script_parts.append('        set myFind to find object of myRange')
                script_parts.append('        clear formatting of myFind')
                script_parts.append('        clear formatting of replacement of myFind')
                if fmt.get("bold") is not None:
                    val = "true" if fmt["bold"] else "false"
                    script_parts.append(f'        set bold of font object of replacement of myFind to {val}')
                if fmt.get("italic") is not None:
                    val = "true" if fmt["italic"] else "false"
                    script_parts.append(f'        set italic of font object of replacement of myFind to {val}')
                if fmt.get("font_size") is not None:
                    script_parts.append(f'        set font size of font object of replacement of myFind to {fmt["font_size"]}')
                if fmt.get("font_name") is not None:
                    script_parts.append(f'        set name of font object of replacement of myFind to "{fmt["font_name"]}"')
                script_parts.append(f'        execute find myFind find text "{find_text}" replace with "{find_text}" replace replace all')

            script_parts.append('    end tell')
            script_parts.append('end tell')

            script = "\n".join(script_parts)
            stdout, stderr, code = run_applescript(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        elif atype == "delete_paragraph":
            para_num = params.get("paragraph", 1)
            script = f'''
            tell application "Microsoft Word"
                tell document 1
                    delete (text object of paragraph {para_num})
                end tell
            end tell
            '''
            stdout, stderr, code = run_applescript(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        elif atype == "save_file":
            script = 'tell application "Microsoft Word" to save document 1'
            stdout, stderr, code = run_applescript(script)
            return {"success": code == 0, "error": stderr if code != 0 else None}

        else:
            return {"success": False, "error": f"Unknown action: {atype}"}

    except Exception as e:
        return {"success": False, "error": str(e)}


def run_applescript(script: str) -> tuple:
    with tempfile.NamedTemporaryFile(mode='w', suffix='.applescript', delete=False) as f:
        f.write(script)
        tmp = f.name
    try:
        result = subprocess.run(
            ["osascript", tmp],
            capture_output=True, text=True, timeout=15
        )
        return result.stdout.strip(), result.stderr.strip(), result.returncode
    finally:
        os.unlink(tmp)


# ─── PLANNER ─────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are a desktop automation agent controlling Microsoft Word through its native API.
You do NOT take screenshots. You receive the full document state as structured JSON
and output a single structured action to perform.

Available actions:
- insert_text: params: {text, position (cursor|start|end)}
- select_paragraph: params: {paragraph (int, 1-indexed)} — selects an entire paragraph
- format_text: params: {bold (bool), italic (bool), font_size (int), font_name (str), paragraph (int, 1-indexed, optional)}
    If "paragraph" is provided, formatting is applied directly to that paragraph.
    If omitted, formatting is applied to the current selection.
- set_style: params: {style (e.g. "Heading 1", "Normal", "Heading 2")}
- find_replace: params: {find (str), replace (str)}
- find_format: params: {formats (list of objects with keys: {find (str), bold (bool, optional), italic (bool, optional)})} — finds all occurrences of each "find" string and applies the specified formatting. Use this to format multiple words/phrases in a single action.
- delete_paragraph: params: {paragraph (int, 1-indexed)} — deletes the specified paragraph
- save_file: params: {}

Output ONLY a JSON object with this structure:
{
  "reasoning": "...",
  "action_type": "...",
  "params": {...},
  "expected_state_delta": {...}
}

Rules:
- Always use API actions, never suggest keystrokes
- NEVER insert text containing raw markdown characters like **, *, _, or ***.
  Any formatting (bold, italic) must be done programmatically using find_format after the text is inserted.
  For example, if you want to write "the **franchise**", insert the text as "the franchise", and then use find_format with find: "franchise", bold: true.
- To format an entire paragraph, use format_text with a "paragraph" param. To format specific words or phrases (such as individual keywords or nouns) within the document, use find_format. Do NOT try to select or format individual words using format_text.
- Differentiating Keywords and Nouns: Pay extreme attention to the user task for formatting instructions. If a task asks for nouns to be both bold and italic, make sure to set both "bold": true and "italic": true on those nouns.
- Identifying Keywords and Nouns: When formatting keywords (bold) and nouns (both bold and italic), identify ALL relevant terms (e.g., "Pokemon", "franchise", "trainers", "Ash Ketchum", "Misty", "Pikachu", "Charizard", "species", "creatures", "regions", "history", "culture", "universe", etc.) and format them in a single find_format action using the "formats" list parameter. Do not stop after formatting just a few.
- Trust find_format success: The paragraph-level state (bold, italic) only shows true if the ENTIRE paragraph has that formatting. Formatting specific words with find_format will not change the paragraph-level bold/italic state to true. Trust that find_format succeeds if execution says OK, and output "done" when you have formatted all target words/phrases once. Do NOT loop formatting the same words.
- Keep expected_state_delta simple — use only top-level keys like word_count, paragraph_count, saved. Do NOT put complex nested objects.
- If task is already complete based on state, output action_type: "done"
- Output raw JSON only, no markdown fences
"""

def plan_action(task: str, state: dict, history: list = None) -> dict:
    # Truncate paragraph text to save tokens
    compact_state = dict(state)
    if "paragraphs" in compact_state:
        compact_state["paragraphs"] = [
            {**p, "text": p["text"][:120] + ("..." if len(p["text"]) > 120 else "")}
            for p in compact_state["paragraphs"]
        ]

    prompt = f"""
TASK: {task}

CURRENT WORD STATE:
{json.dumps(compact_state, indent=2)}
"""
    if history:
        prompt += "\nPREVIOUS ACTIONS (do NOT repeat failed ones):\n"
        for h in history:
            status = "OK" if h.get("success") else f"FAILED: {h.get('error', 'unknown')}"
            prompt += f"- {h['action_type']}({json.dumps(h.get('params', {}))}) -> {status}\n"

    prompt += "\nOutput the next action to take.\n"

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": prompt}
    ]

    models = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "meta-llama/llama-4-scout-17b-16e-instruct"]
    model_idx = 0

    for attempt in range(3):
        response = None
        raw = None
        for _ in range(len(models)):
            model = models[model_idx % len(models)]
            try:
                response = client.chat.completions.create(
                    model=model,
                    max_tokens=2048,
                    messages=messages
                )
                raw = response.choices[0].message.content.strip()
                break
            except Exception as e:
                err_msg = str(e)
                if "rate limit" in err_msg.lower() or "429" in err_msg:
                    print(f"  Rate limit hit for model {model}. Trying alternative model...")
                    model_idx += 1
                else:
                    raise
        else:
            raise Exception("All models hit rate limits or failed.")

        try:
            # Strip markdown fences if model adds them
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
            
            raw_stripped = raw.strip()
            if raw_stripped.endswith("```"):
                raw_stripped = raw_stripped[:-3].strip()
            
            obj, _ = json.JSONDecoder().raw_decode(raw_stripped)
            return obj
        except json.JSONDecodeError as e:
            print(f"  JSON parse error (attempt {attempt + 1}/3) with model {model}: {e}")
            if attempt == 2:
                print(f"  Raw LLM output: {raw[:300]}...")
                raise
            # Retry with a nudge
            messages.append({"role": "assistant", "content": raw})
            messages.append({"role": "user", "content": "That was not valid JSON. Output ONLY a raw JSON object, no markdown, no extra text."})


# ─── VERIFIER ────────────────────────────────────────────────────────────────

def verify(expected_delta: dict, state_before: dict, state_after: dict) -> bool:
    for key, expected in expected_delta.items():
        after_val = state_after.get(key)
        before_val = state_before.get(key, 0)

        if isinstance(expected, str) and expected.startswith("+"):
            delta = int(expected[1:])
            if after_val != before_val + delta:
                print(f"  Verify FAIL: {key} expected {before_val}+{delta}={before_val+delta}, got {after_val}")
                return False
        elif isinstance(expected, str) and expected.startswith("-"):
            delta = int(expected[1:])
            if after_val != before_val - delta:
                print(f"  Verify FAIL: {key} expected {before_val}-{delta}={before_val-delta}, got {after_val}")
                return False
        elif key == "paragraphs" and isinstance(expected, list) and isinstance(after_val, list):
            # Smart paragraph check: verify each expected paragraph's properties are matched by at least one paragraph in state_after.
            for exp_para in expected:
                if isinstance(exp_para, dict):
                    match_found = False
                    for act_para in after_val:
                        if isinstance(act_para, dict):
                            item_match = True
                            for k, v in exp_para.items():
                                if k == "text":
                                    v_clean = v.strip().replace("\r", "").replace("\n", "")
                                    act_clean = act_para.get("text", "").strip().replace("\r", "").replace("\n", "")
                                    if v_clean not in act_clean and act_clean not in v_clean:
                                        item_match = False
                                        break
                                else:
                                    if act_para.get(k) != v:
                                        item_match = False
                                        break
                            if item_match:
                                match_found = True
                                break
                    if not match_found:
                        print(f"  Verify FAIL: Could not find paragraph matching expected criteria: {exp_para}")
                        return False
        else:
            if after_val != expected:
                print(f"  Verify FAIL: {key} expected {expected}, got {after_val}")
                return False
    return True


# ─── MAIN LOOP ───────────────────────────────────────────────────────────────

def run(task: str, max_steps: int = 25):
    print(f"\n{'='*60}")
    print(f"TASK: {task}")
    print(f"{'='*60}\n")

    history = []

    for step in range(max_steps):
        print(f"--- Step {step + 1} ---")

        # Read state
        state_before = get_word_state_mac()
        if "error" in state_before:
            print(f"ERROR: {state_before['error']}")
            break

        print(f"State: {state_before['paragraph_count']} paragraphs, "
              f"{state_before['word_count']} words, "
              f"saved={state_before['saved']}")
        if state_before['paragraphs']:
            print(f"First para: '{state_before['paragraphs'][0]['text'][:50]}...'")

        # Plan
        print("Planning action...")
        try:
            action = plan_action(task, state_before, history)
        except Exception as e:
            print(f"Planning FAILED: {e}")
            break
        print(f"Action: {action['action_type']} | {action.get('reasoning', '')}")
        print(f"Params: {json.dumps(action.get('params', {}))}")

        if action["action_type"] == "done":
            print("\nTask complete (agent says done).")
            break

        # Execute
        result = execute_action_mac(action)
        if not result["success"]:
            print(f"Execution FAILED: {result.get('error')}")
            history.append({**action, "success": False, "error": result.get("error")})
            # Don't break — let the agent try a different approach
            print()
            continue
        print("Execution: OK")

        # Verify
        state_after = get_word_state_mac()
        if "error" in state_after:
            print(f"ERROR reading state after action: {state_after['error']}")
            break

        delta = action.get("expected_state_delta", {})
        ok = None
        if delta:
            ok = verify(delta, state_before, state_after)
            print(f"Verification: {'PASS' if ok else 'FAIL'}")
        else:
            print("Verification: skipped (no delta specified)")

        history.append({**action, "success": True})

        # Save trajectory step
        with open("trajectory.jsonl", "a") as f:
            f.write(json.dumps({
                "step": step + 1,
                "task": task,
                "action": action,
                "result": result,
                "state_before": state_before,
                "state_after": state_after,
                "verified": ok
            }) + "\n")

        print()

    print("Done. Trajectory saved to trajectory.jsonl")


def open_new_document():
    """Create a new blank Word document."""
    script = '''
    tell application "Microsoft Word"
        activate
        make new document
    end tell
    '''
    stdout, stderr, code = run_applescript(script)
    if code == 0:
        print("\n✅ New document created.")
    else:
        print(f"\n❌ Failed to create document: {stderr}")
    return code == 0


def open_existing_document(path: str):
    """Open an existing Word document by file path."""
    path = os.path.expanduser(path.strip())
    if not os.path.exists(path):
        print(f"\n❌ File not found: {path}")
        return False
    abs_path = os.path.abspath(path)
    script = f'''
    tell application "Microsoft Word"
        activate
        open "{abs_path}"
    end tell
    '''
    stdout, stderr, code = run_applescript(script)
    if code == 0:
        print(f"\n✅ Opened: {abs_path}")
    else:
        print(f"\n❌ Failed to open document: {stderr}")
    return code == 0


def list_open_documents():
    """List all currently open Word documents."""
    script = '''
    tell application "Microsoft Word"
        try
            set doc_names to name of every document
            set oldDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to linefeed
            set res to doc_names as string
            set AppleScript's text item delimiters to oldDelims
            return res
        on error
            return ""
        end try
    end tell
    '''
    stdout, stderr, code = run_applescript(script)
    if code == 0 and stdout.strip():
        docs = stdout.strip().split('\n')
        return docs
    return []


def interactive_prompt():
    """Interactive prompt that asks the user what they want to do."""
    print("\n" + "=" * 60)
    print("  📝  ScriptAgent — Microsoft Word Automation")
    print("=" * 60)

    # Check if Word has any open documents
    open_docs = list_open_documents()

    if open_docs:
        print(f"\n📄 Currently open document(s):")
        for i, doc in enumerate(open_docs, 1):
            print(f"   {i}. {doc}")
    else:
        print("\n⚠️  No documents currently open in Word.")

    # Ask what the user wants to do
    print("\n🔧 What would you like to do?")
    print("   1. Edit the current document")
    print("   2. Open an existing file")
    print("   3. Create a new blank document")
    print("   4. Quit")

    while True:
        choice = input("\n👉 Enter choice (1-4): ").strip()
        if choice in ("1", "2", "3", "4"):
            break
        print("   Invalid choice. Please enter 1, 2, 3, or 4.")

    if choice == "4":
        print("\n👋 Goodbye!")
        return

    if choice == "2":
        path = input("\n📂 Enter the file path to open: ").strip()
        if not open_existing_document(path):
            return

    elif choice == "3":
        if not open_new_document():
            return

    elif choice == "1":
        if not open_docs:
            print("\n❌ No document is open. Please open or create one first.")
            return

    # Now ask for the task in a loop
    print("\n" + "-" * 60)
    print("💬 Describe what you want to do with the document.")
    print("   Type 'quit' or 'exit' to stop.")
    print("-" * 60)

    while True:
        task = input("\n🎯 Task: ").strip()
        if not task:
            print("   Please enter a task description.")
            continue
        if task.lower() in ("quit", "exit", "q"):
            print("\n👋 Goodbye!")
            break

        run(task)

        cont = input("\n🔄 Do you want to perform another task? (y/n): ").strip().lower()
        if cont not in ("y", "yes"):
            print("\n👋 Goodbye!")
            break


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        # CLI mode: run the task directly
        task = " ".join(sys.argv[1:])
        run(task)
    else:
        # Interactive mode
        interactive_prompt()