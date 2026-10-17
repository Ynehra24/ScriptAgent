"""
CLI entry point for word_automation.

Usage:
  # Create a new document:
  python -m word_automation "Create a Word doc named proposal.docx with title 'Sales Report' and a 3x3 table"

  # Reference and edit an existing document without overwriting:
  python -m word_automation --doc memo.docx "Add a 3rd bullet point summarizing QA test results"
  python -m word_automation "In memo.docx, add a 3rd bullet point summarizing QA test results"

  # Explicitly overwrite an existing document:
  python -m word_automation --doc memo.docx --overwrite "Create a fresh memo with title 'New Memo'"

  # Interactive mode:
  python -m word_automation
"""
from __future__ import annotations

import argparse # Module to make adding arguments to function calling much faster and easier -> handles both positional and optional arguments
import json # Encoding of heirarchial information so that its easier to parse -> Either convert from python format to JSON or from JSON to python format
import sys # Gives access to certain functions/objects which directly work with the interpreter/are maintained by the interpreter
from .config import AutomationConfig # Imported from the config.py file -> Creates a populated config file which has the values for all the API_KEY, URL etc
from .pipeline import WordAutomationPipeline # Just to execute one request end to end


def main():
    parser = argparse.ArgumentParser(
        description="3-Layer Word Automation System (Nemotron + python-docx + Osaurus AppleScript)"
    )
    parser.add_argument(
        "prompt",
        nargs="*",
        help="The automation instruction or prompt to execute",
    )
    parser.add_argument(
        "-d", "--doc", "--document",
        type=str,
        default=None,
        help="Explicitly reference a target/existing Word document path",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Explicitly overwrite/recreate the document if it already exists",
    )
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="Suppress verbose progress logs",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output raw JSON results",
    )
    parser.add_argument(
        "--model",
        type=str,
        default=None,
        help="Override Nemotron model",
    )
    parser.add_argument(
        "--folder", 
        type = str, 
        default = None, 
        help = "Puts the file in a pre-determined folder",
    )

    args = parser.parse_args() # Runs parser and puts everything in a namespace object -> A space of all the different names and what objects they reference

    config = AutomationConfig.from_env()
    if args.model:
        config.nemotron_model = args.model

    pipeline = WordAutomationPipeline(config=config)

    prompt_text = " ".join(args.prompt).strip()

    # Interactive mode if no prompt provided
    if not prompt_text:
        print("=" * 60)
        print("🤖 Word Automation Interactive Shell")
        print("Commands: Type your prompt, or '--doc <file> <prompt>'")
        print("Type 'exit' or 'quit' to exit.")
        print("=" * 60)
        while True:
            try:
                user_input = input("\n📝 Prompt > ").strip()
                if not user_input:
                    continue
                if user_input.lower() in ("exit", "quit", "q"):
                    print("Exiting.")
                    break

                result = pipeline.execute(
                    user_input,
                    target_file=args.doc,
                    overwrite=args.overwrite,
                    verbose=not args.quiet,
                    folder = args.folder
                )
                if args.json:
                    print("\nJSON Result:")
                    print(json.dumps(result, indent=2, default=str))

            except (KeyboardInterrupt, EOFError):
                print("\nExiting.")
                break
        return

    # Single prompt execution
    result = pipeline.execute(
        prompt_text,
        target_file=args.doc,
        overwrite=args.overwrite,
        verbose=not args.quiet,
        folder = args.folder
    )
    if args.json:
        print("\nJSON Result:")
        print(json.dumps(result, indent=2, default=str))

    if result["status"] == "failure":
        sys.exit(1)


if __name__ == "__main__":
    main()
