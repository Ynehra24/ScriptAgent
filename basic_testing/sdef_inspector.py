import hashlib
import json
import os
import sqlite3
import subprocess
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional
import xml.etree.ElementTree as ET

CACHE_DIR = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "scriptagent_sdef"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

def _strip_ns(tag: str) -> str:
    return tag.split('}', 1)[-1] if '}' in tag else tag


def run_sdef(app_path: str, timeout: int = 10) -> str:
    """Run `sdef` on an application bundle and return XML output as string.

    Raises subprocess.CalledProcessError on failure.
    """
    # Prefer full path if given
    proc = subprocess.run(["sdef", app_path], capture_output=True, text=True, timeout=timeout)
    if proc.returncode != 0:
        raise subprocess.CalledProcessError(proc.returncode, proc.args, output=proc.stdout, stderr=proc.stderr)
    return proc.stdout


@dataclass
class ParameterInfo:
    name: str
    type: Optional[str] = None
    description: Optional[str] = None


@dataclass
class CommandInfo:
    name: str
    description: Optional[str] = None
    parameters: List[ParameterInfo] = field(default_factory=list)


@dataclass
class PropertyInfo:
    name: str
    type: Optional[str] = None
    description: Optional[str] = None


@dataclass
class ClassInfo:
    name: str
    description: Optional[str] = None
    properties: List[PropertyInfo] = field(default_factory=list)


def parse_sdef(xml_text: str) -> Dict[str, Any]:
    root = ET.fromstring(xml_text)
    commands: Dict[str, Dict[str, Any]] = {}
    classes: Dict[str, Dict[str, Any]] = {}
    events: Dict[str, Dict[str, Any]] = {}
    enums: Dict[str, Dict[str, Any]] = {}

    for elem in root.iter():
        tag = _strip_ns(elem.tag)
        if tag == "command":
            name = elem.attrib.get("name") or elem.attrib.get("id")
            if not name:
                continue
            desc = None
            params: List[ParameterInfo] = []
            return_type: Optional[str] = elem.attrib.get("result-type") or elem.attrib.get("return-type")
            for child in elem:
                ctag = _strip_ns(child.tag)
                if ctag in ("description", "doc", "documentation"):
                    desc = (child.text or "").strip()
                if ctag == "parameter":
                    pname = child.attrib.get("name") or child.attrib.get("id")
                    ptype = child.attrib.get("type")
                    pdesc = (child.text or "").strip() if child.text else None
                    if pname:
                        params.append(ParameterInfo(name=pname, type=ptype, description=pdesc))
                if ctag in ("result", "returns"):
                    # try to extract result type from child
                    rtype = child.attrib.get("type") or (child.text or None)
                    if rtype:
                        return_type = rtype
            commands[name] = {
                "name": name,
                "description": desc,
                "return_type": return_type,
                "parameters": [asdict(p) for p in params],
            }

        if tag == "class":
            cname = elem.attrib.get("name") or elem.attrib.get("id")
            if not cname:
                continue
            cdesc = None
            props: List[PropertyInfo] = []
            for child in elem:
                ctag = _strip_ns(child.tag)
                if ctag in ("description", "doc", "documentation"):
                    cdesc = (child.text or "").strip()
                if ctag == "property":
                    pname = child.attrib.get("name") or child.attrib.get("id")
                    ptype = child.attrib.get("type")
                    pdesc = (child.text or "").strip() if child.text else None
                    if pname:
                        props.append(PropertyInfo(name=pname, type=ptype, description=pdesc))
            classes[cname] = {
                "name": cname,
                "description": cdesc,
                "properties": [asdict(p) for p in props],
            }

        if tag == "event":
            ename = elem.attrib.get("name") or elem.attrib.get("id")
            if not ename:
                continue
            edesc = None
            eparams: List[ParameterInfo] = []
            for child in elem:
                ctag = _strip_ns(child.tag)
                if ctag in ("description", "doc", "documentation"):
                    edesc = (child.text or "").strip()
                if ctag == "parameter":
                    pname = child.attrib.get("name") or child.attrib.get("id")
                    ptype = child.attrib.get("type")
                    pdesc = (child.text or "").strip() if child.text else None
                    if pname:
                        eparams.append(ParameterInfo(name=pname, type=ptype, description=pdesc))
            events[ename] = {
                "name": ename,
                "description": edesc,
                "parameters": [asdict(p) for p in eparams],
            }

        if tag == "enumeration":
            enum_name = elem.attrib.get("name") or elem.attrib.get("id")
            if not enum_name:
                continue
            items: List[Dict[str, Any]] = []
            for child in elem:
                ctag = _strip_ns(child.tag)
                if ctag in ("item", "enumerator", "constant"):
                    iname = child.attrib.get("name") or child.attrib.get("id")
                    ival = child.attrib.get("value") or child.text
                    if iname:
                        items.append({"name": iname, "value": ival})
            enums[enum_name] = {"name": enum_name, "items": items}

    return {"commands": commands, "classes": classes, "events": events, "enumerations": enums}


def _safe_name(path: str) -> str:
    return path.replace('/', '_').replace(' ', '_')


def _hash_text(text: str) -> str:
    return hashlib.sha1(text.encode('utf-8')).hexdigest()


def cache_path_for(app_path: str) -> Path:
    name = _safe_name(app_path)
    return CACHE_DIR / f"{name}.json"


def default_db_path() -> Path:
    return CACHE_DIR / "sdef_index.db"


def _ensure_db(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS sdef_index (app_path TEXT PRIMARY KEY, xml_hash TEXT, generated_at INTEGER, index_json TEXT)"
    )
    conn.commit()
    return conn


def save_to_sqlite(app_path: str, xml_text: str, index: Dict[str, Any], db_path: Optional[str] = None) -> None:
    db_path = db_path or str(default_db_path())
    conn = _ensure_db(db_path)
    payload = json.dumps(index, separators=(",", ":"))
    conn.execute(
        "INSERT OR REPLACE INTO sdef_index (app_path, xml_hash, generated_at, index_json) VALUES (?, ?, ?, ?)",
        (app_path, _hash_text(xml_text), int(time.time()), payload),
    )
    conn.commit()
    conn.close()


def load_from_sqlite(app_path: str, db_path: Optional[str] = None) -> Optional[Dict[str, Any]]:
    db_path = db_path or str(default_db_path())
    if not Path(db_path).exists():
        return None
    try:
        conn = sqlite3.connect(db_path)
        cur = conn.execute("SELECT xml_hash, index_json FROM sdef_index WHERE app_path = ?", (app_path,))
        row = cur.fetchone()
        conn.close()
        if not row:
            return None
        xml_hash, index_json = row
        return {"xml_hash": xml_hash, "index": json.loads(index_json)}
    except Exception:
        return None


def save_cache(app_path: str, xml_text: str, index: Dict[str, Any]) -> None:
    path = cache_path_for(app_path)
    payload = {
        "app_path": app_path,
        "generated_at": int(time.time()),
        "xml_hash": _hash_text(xml_text),
        "index": index,
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_cache(app_path: str) -> Optional[Dict[str, Any]]:
    path = cache_path_for(app_path)
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        return payload
    except Exception:
        return None


def inspect(app_path: str, refresh: bool = False, timeout: int = 10, use_sqlite: bool = False, db_path: Optional[str] = None) -> Dict[str, Any]:
    """Return an index of commands and classes for `app_path`.

    If a cached index exists and the sdef hasn't changed, return cache.
    Set `refresh=True` to force re-fetching and re-parsing.
    """
    cached = None
    if use_sqlite:
        cached = load_from_sqlite(app_path, db_path=db_path)
    else:
        cached = load_cache(app_path)

    if cached and not refresh:
        try:
            xml_text = run_sdef(app_path, timeout=timeout)
        except Exception:
            # If sdef fails, fall back to cache
            return cached["index"]
        if _hash_text(xml_text) == cached.get("xml_hash"):
            return cached["index"]

    xml_text = run_sdef(app_path, timeout=timeout)
    index = parse_sdef(xml_text)
    try:
        # persist to both cache formats depending on flags
        if use_sqlite:
            save_to_sqlite(app_path, xml_text, index, db_path=db_path)
        else:
            save_cache(app_path, xml_text, index)
    except Exception:
        pass
    return index


def query_commands(index: Dict[str, Any], name_substr: str) -> List[Dict[str, Any]]:
    name_substr = name_substr.lower()
    out = []
    for k, v in index.get("commands", {}).items():
        if name_substr in k.lower() or (v.get("description") and name_substr in v.get("description", "").lower()):
            out.append(v)
    return out


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Inspect sdef for an app and cache results")
    parser.add_argument("app", help="Path to app bundle (e.g. /Applications/Preview.app)")
    parser.add_argument("--refresh", action="store_true", help="Force refresh ignoring cache")
    parser.add_argument("--query", help="Filter commands by substring")
    parser.add_argument("--use-sqlite", action="store_true", help="Use SQLite-backed cache for faster queries")
    parser.add_argument("--db-path", help="Path to sqlite DB file (default in cache dir)")
    args = parser.parse_args()

    try:
        idx = inspect(args.app, refresh=args.refresh, use_sqlite=args.use_sqlite, db_path=args.db_path)
    except subprocess.CalledProcessError as e:
        print("sdef failed:", e.stderr or e.output)
        raise SystemExit(2)

    if args.query:
        results = query_commands(idx, args.query)
        print(json.dumps(results, indent=2))
    else:
        # print a small summary
        print(json.dumps({
            "commands_count": len(idx.get("commands", {})),
            "classes_count": len(idx.get("classes", {})),
        }, indent=2))
