from .documents import list_open_documents, open_existing_document, open_new_document
from .runner import run

def interactive_prompt():
    print("\n" + "=" * 60)
    print("  ScriptAgent — Microsoft Word Automation")
    print("=" * 60)
    open_documents = list_open_documents()
    if open_documents:
        print("\nCurrently open document(s):")
        for index, document in enumerate(open_documents, 1):
            print(f"   {index}. {document}")
    else:
        print("\nNo documents are currently open in Word.")

    print("\nWhat would you like to do?")
    print("   1. Edit the current document")
    print("   2. Open an existing file")
    print("   3. Create a new blank document")
    print("   4. Quit")

    while True:
        choice = input("\nEnter choice (1-4): ").strip()
        if choice in {"1", "2", "3", "4"}:
            break
        print("Invalid choice.")

    if choice == "4":
        return
    if choice == "2" and not open_existing_document(input("\nFile path: ")):
        return
    if choice == "3" and not open_new_document():
        return
    if choice == "1" and not open_documents:
        print("\nNo document is open.")
        return

    while True:
        task = input("\nTask (or 'quit'): ").strip()
        if task.lower() in {"quit", "exit", "q"}:
            break
        if task:
            run(task)


def main():
    import sys

    if len(sys.argv) > 1:
        run(" ".join(sys.argv[1:]))
    else:
        interactive_prompt()
