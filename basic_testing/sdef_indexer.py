"""Background indexer for sdef files.

Scans application directories for .app bundles and indexes their sdef into the
SQLite cache used by `sdef_inspector`.
"""
from concurrent.futures import ThreadPoolExecutor, as_completed
import os
from pathlib import Path
from typing import Iterable, List, Optional

from .sdef_inspector import inspect, default_db_path


def find_apps(paths: Iterable[str]) -> List[str]:
    apps = []
    for p in paths:
        root = Path(p).expanduser()
        if not root.exists():
            continue
        for child in root.rglob("*.app"):
            apps.append(str(child))
    return sorted(set(apps))


def index_apps(paths: Iterable[str], db_path: Optional[str] = None, workers: int = 4, refresh: bool = False):
    """Index all apps found under `paths` into the sqlite DB.

    This is best-effort and will skip apps that fail to parse.
    """
    apps = find_apps(paths)
    if not apps:
        return
    db_path = db_path or str(default_db_path())
    with ThreadPoolExecutor(max_workers=workers) as exe:
        futures = {exe.submit(inspect, app, refresh, 10, True, db_path): app for app in apps}
        for fut in as_completed(futures):
            app = futures[fut]
            try:
                idx = fut.result()
                print(f"Indexed {app}: {len(idx.get('commands', {}))} commands, {len(idx.get('classes', {}))} classes")
            except Exception as exc:
                print(f"Failed to index {app}: {exc}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Index sdef for apps under given folders")
    parser.add_argument("paths", nargs="*", default=["/Applications", "~/Applications"], help="Folders to scan for .app bundles")
    parser.add_argument("--db-path", help="SQLite DB path to write to")
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--refresh", action="store_true")
    args = parser.parse_args()
    index_apps(args.paths, db_path=args.db_path, workers=args.workers, refresh=args.refresh)
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Iterable, List, Optional

from .sdef_inspector import inspect


def find_app_bundles(paths: Iterable[str]) -> List[str]:
    apps = []
    for p in paths:
        base = Path(p)
        if not base.exists():
            continue
        if base.is_dir():
            for child in base.iterdir():
                if child.suffix == ".app":
                    apps.append(str(child))
                elif child.is_dir():
                    # shallow search
                    for sub in child.iterdir():
                        if sub.suffix == ".app":
                            apps.append(str(sub))
        elif str(base).endswith(".app"):
            apps.append(str(base))
    return sorted(set(apps))


def index_applications(paths: Iterable[str], workers: int = 6, refresh: bool = False, db_path: Optional[str] = None):
    apps = find_app_bundles(paths)
    print(f"Found {len(apps)} apps to index")
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(inspect, app, refresh, 15, True, db_path): app for app in apps}
        for fut in as_completed(futures):
            app = futures[fut]
            try:
                fut.result()
                print(f"Indexed: {app}")
            except Exception as e:
                print(f"Failed: {app} -> {e}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Index sdef for applications under given paths")
    parser.add_argument("paths", nargs="*", default=["/Applications", "/System/Applications"], help="Directories to scan for .app bundles")
    parser.add_argument("--workers", type=int, default=6)
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument("--db-path", help="SQLite DB path to store index")
    args = parser.parse_args()

    index_applications(args.paths, workers=args.workers, refresh=args.refresh, db_path=args.db_path)
