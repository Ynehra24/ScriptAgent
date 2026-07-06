# Basic Testing Module Map

This folder contains the code that used to live in `tester.py`, split by
responsibility so each part is easier to inspect and improve.

- `cli.py` and `__main__.py`: command-line and interactive entrypoints.
- `runner.py`: main task loop that reads state, asks the planner, executes, and logs.
- `config.py`: environment loading, constants, logging, and timing helpers.
- `applescript.py`: low-level AppleScript execution.
- `documents.py`: open/create/list Microsoft Word documents.
- `word_schema.py`: parses Word's AppleScript dictionary.
- `property_path.py`: validates property paths and converts them to AppleScript.
- `property_actions.py`: reads, sets, verifies, and compacts Word properties.
- `content_actions.py`: reads document content and performs text actions.
- `gemini_client.py`: Gemini client setup and JSON response handling.
- `prompts.py`: planner and selector prompts.
- `selectors.py`: chooses relevant Word classes and metadata fields.
- `planner.py`: builds planner input and calls Gemini.
- `utils.py`: small shared helpers.

You can still run the project exactly as before:

```bash
python tester.py
python tester.py "your Word task here"
```

You can also run the package directly:

```bash
python -m basic_testing
```
