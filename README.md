# ScriptAgent

ScriptAgent is an early work-in-progress project for controlling desktop UI and
applications with vision/language-model style agents. The current runnable
prototype focuses on Microsoft Word automation on macOS through AppleScript and
Gemini-based planning.

This repository is not production-ready yet. It is currently a learning and
experimentation workspace with several holes, rough edges, and partially explored
directions.

## Current Status

The most developed part of the project is `basic_testing`, which was split out
from the original `tester.py` script so the automation loop is easier to inspect
and improve.

Currently implemented:

- Opens an interactive CLI for Microsoft Word automation.
- Can edit the currently open Word document, open an existing file, or create a
  new blank document.
- Reads basic document state from Word:
  - document name
  - saved status
  - word count
  - paragraph count
  - selected text preview
  - first 100 paragraph previews
- Parses Microsoft Word's installed AppleScript dictionary (`Word.sdef`) to build
  a searchable schema of classes, properties, elements, commands, and enums.
- Uses Gemini to:
  - select relevant Word classes for a task
  - select metadata/property fields to inspect
  - plan the next structured action
- Supports these planner actions:
  - `insert_text`
  - `replace_text`
  - `format_matches`
  - `delete_paragraph`
  - `save_document`
  - `set_properties`
  - `done`
- Supports insertion formatting options:
  - font name
  - font size
  - bold
  - italic
  - RGB font color
- Supports property-path based Word metadata/settings edits when the path can be
  resolved safely through the parsed Word schema.
- Verifies property changes by reading values before and after the edit.
- Writes timing/debug logs to `scriptagent.log`.
- Appends task/action/result records to `trajectory.jsonl`.

## Known Gaps

This is still very much a work in progress.

- Only macOS Microsoft Word automation is implemented right now.
- The Word automation depends on AppleScript object codes and Word's local
  scripting dictionary, so behavior may vary across Word versions.
- There is no formal test suite yet.
- There is no dependency file yet (`requirements.txt` / `pyproject.toml`).
- Error recovery is basic; failed planner actions are recorded, but the agent can
  still loop or choose weak follow-up actions.
- Content verification is shallow compared with property verification.
- The project name and larger text-based desktop-control direction are broader than the
  current Word-only prototype.
- Dataset and paper folders are present for research/reference work, but they are
  not wired into the current Word automation loop.
- Generated local artifacts such as logs, trajectories, and Python caches may
  exist in the working tree during development.

## Requirements

Current practical requirements:

- macOS
- Microsoft Word installed
- Python 3.12+ recommended
- A Gemini API key in `.env`:

```bash
GEMINI_KEY=your_key_here
```

Python packages used by the current prototype:

- `python-dotenv`
- `google-genai`
- `tqdm`

Install them manually for now:

```bash
pip install python-dotenv google-genai tqdm
```

## Running

Interactive mode:

```bash
python tester.py
```

Run a task directly:

```bash
python tester.py "Insert a short title at the start of the document"
```

Run the package directly:

```bash
python -m basic_testing
```

## Folder Structure

```text
ScriptAgent/
├── README.md
├── LICENSE
├── tester.py
├── basic_testing/
│   ├── README.md
│   ├── __init__.py
│   ├── __main__.py
│   ├── applescript.py
│   ├── cli.py
│   ├── config.py
│   ├── content_actions.py
│   ├── documents.py
│   ├── gemini_client.py
│   ├── planner.py
│   ├── prompts.py
│   ├── property_actions.py
│   ├── property_path.py
│   ├── runner.py
│   ├── selectors.py
│   ├── utils.py
│   └── word_schema.py
├── datasets/
│   └── ShowUI-desktop-full/
├── papers/
│   ├── OpenCUA.pdf
│   └── S2Agent.pdf
├── scratch/
│   └── query.py
├── scriptagent.log
└── trajectory.jsonl
```

## Basic Testing Module Map

- `tester.py`: compatibility launcher that calls `basic_testing.cli.main`.
- `basic_testing/cli.py`: command-line and interactive entrypoints.
- `basic_testing/runner.py`: main loop that reads state, plans, executes, logs,
  and records trajectory entries.
- `basic_testing/config.py`: environment loading, constants, logging, and timing.
- `basic_testing/applescript.py`: low-level AppleScript execution helper.
- `basic_testing/documents.py`: open/create/list Word documents.
- `basic_testing/word_schema.py`: parses Word's AppleScript dictionary.
- `basic_testing/property_path.py`: validates canonical property paths and
  compiles them into AppleScript expressions.
- `basic_testing/property_actions.py`: reads, sets, verifies, and compacts Word
  properties.
- `basic_testing/content_actions.py`: reads document content and performs text
  actions.
- `basic_testing/gemini_client.py`: Gemini client setup, retries, and JSON
  response handling.
- `basic_testing/prompts.py`: planner and selector prompts.
- `basic_testing/selectors.py`: chooses relevant Word classes and metadata fields.
- `basic_testing/planner.py`: builds planner input and calls Gemini.
- `basic_testing/utils.py`: shared string/name helpers.

## Development Notes

Use this command for a quick syntax check:

```bash
python -m py_compile tester.py basic_testing/*.py
```

Important local artifacts:

- `scriptagent.log`: rotating debug/timing log.
- `trajectory.jsonl`: JSONL history of planner actions and execution results.

Suggested next improvements:

- Add `requirements.txt` or `pyproject.toml`.
- Add a `.gitignore` for logs, caches, `.env`, datasets, and generated artifacts.
- Add unit tests for path parsing, value coercion, and planner action validation.
- Add safer dry-run or confirmation mode for destructive Word actions.
- Improve content-level verification after insert/replace/delete/format actions.
- Separate research assets from runnable prototype code.
