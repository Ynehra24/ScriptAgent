"""Simple interactive browser for the sdef sqlite cache.

Provides a minimal TUI using `rich` for listing apps and searching commands.
"""
import json
from pathlib import Path
from typing import Optional

from rich.console import Console
from rich.table import Table

from .sdef_inspector import default_db_path, load_from_sqlite


console = Console()


def list_indexed_apps(db_path: Optional[str] = None):
    db_path = db_path or str(default_db_path())
    if not Path(db_path).exists():
        console.print("No sqlite index found.")
        return []
    try:
        import sqlite3

        conn = sqlite3.connect(db_path)
        cur = conn.execute("SELECT app_path, generated_at FROM sdef_index ORDER BY app_path")
        rows = cur.fetchall()
        conn.close()
    except Exception as exc:
        console.print("Failed to read DB:", exc)
        return []
    return rows


def show_apps(db_path: Optional[str] = None):
    rows = list_indexed_apps(db_path)
    if not rows:
        return
    table = Table(title="Indexed Apps")
    table.add_column("#", style="cyan")
    table.add_column("app_path", style="green")
    table.add_column("generated_at", style="yellow")
    for i, (app_path, gen) in enumerate(rows, 1):
        table.add_row(str(i), app_path, str(gen))
    console.print(table)


def select_app_and_search(db_path: Optional[str] = None):
    rows = list_indexed_apps(db_path)
    if not rows:
        return
    show_apps(db_path)
    choice = console.input("Enter app # to inspect (or 'q' to quit): ")
    if choice.strip().lower() in ("q", "quit", "exit"):
        return
    try:
        idx = int(choice.strip()) - 1
        app_path = rows[idx][0]
    except Exception:
        console.print("Invalid selection")
        return

    data = load_from_sqlite(app_path, db_path=db_path)
    if not data:
        console.print("No index for that app")
        return
    index = data["index"]
    while True:
        q = console.input("Search commands (substring) or 'b' to go back: ")
        if q.strip().lower() in ("b", "back"):
            return
        if not q.strip():
            continue
        results = []
        for k, v in index.get("commands", {}).items():
            if q.lower() in k.lower() or (v.get("description") and q.lower() in v.get("description", "").lower()):
                results.append(v)
        if not results:
            console.print("No matching commands found")
            continue
        table = Table(title=f"Commands matching '{q}'")
        table.add_column("name", style="cyan")
        table.add_column("return_type", style="green")
        table.add_column("params", style="magenta")
        table.add_column("description", style="yellow")
        for r in results:
            params = ", ".join(p.get("name") for p in r.get("parameters", []))
            table.add_row(r.get("name", ""), str(r.get("return_type", "")), params, r.get("description") or "")
        console.print(table)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Browse sdef sqlite cache")
    parser.add_argument("--db-path")
    args = parser.parse_args()
    while True:
        show_apps(args.db_path)
        select_app_and_search(args.db_path)
        cont = console.input("Index another app? (y/N): ")
        if cont.strip().lower() not in ("y", "yes"):
            break
import json
import sqlite3
from pathlib import Path
from typing import Optional

from .sdef_inspector import default_db_path, load_from_sqlite, query_commands

try:
    from rich.console import Console
    from rich.table import Table
except Exception:
    Console = None


def list_indexed_apps(db_path: Optional[str] = None):
    db_path = db_path or str(default_db_path())
    if not Path(db_path).exists():
        print("No index DB found.")
        return []
    conn = sqlite3.connect(db_path)
    cur = conn.execute("SELECT app_path, generated_at FROM sdef_index ORDER BY app_path")
    rows = cur.fetchall()
    conn.close()
    return [r[0] for r in rows]


def pretty_print_commands(cmds):
    if Console:
        cons = Console()
        table = Table()
        table.add_column("Name")
        table.add_column("Return")
        table.add_column("Parameters")
        table.add_column("Description")
        for c in cmds:
            params = ", ".join([p.get("name") + (":" + (p.get("type") or "")) for p in c.get("parameters", [])])
            table.add_row(c.get("name", ""), str(c.get("return_type", "")), params, c.get("description") or "")
        cons.print(table)
    else:
        print(json.dumps(cmds, indent=2))


def browse(db_path: Optional[str] = None):
    apps = list_indexed_apps(db_path)
    if not apps:
        return
    print("Indexed apps:")
    for i, a in enumerate(apps):
        print(f"[{i}] {a}")
    sel = input("Select app index: ")
    try:
        idx = int(sel)
        app = apps[idx]
    except Exception:
        print("invalid selection")
        return
    idx_obj = load_from_sqlite(app, db_path=db_path)
    if not idx_obj:
        print("No index for app")
        return
    while True:
        q = input("query (substring, empty to quit): ")
        if not q:
            break
        results = query_commands(idx_obj["index"], q)
        pretty_print_commands(results)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Browse sdef index DB")
    parser.add_argument("--db-path", help="Path to sqlite DB to read")
    args = parser.parse_args()
    browse(db_path=args.db_path)
