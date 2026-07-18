from dotenv import load_dotenv
import os
import pandas as pd
import sys
import time
load_dotenv()

API_KEY = os.environ.get("OPENROUTER_API_KEY")
if API_KEY:
    print("Successfully loaded key!")

main_model = os.environ.get("OPENROUTER_MODELS")
if main_model:
    print(f"Successfully loaded model {main_model}")

sdef_file = "/Applications/Microsoft Word.app/Contents/Resources/Word.sdef"

# Logging
import logging
LOG_FILE = os.environ.get("SCRIPTAGENT_LOG", "scriptagent.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler()],
)
LOGGER = logging.getLogger("scriptagent")

# Timing threshold for slow operations (seconds)
SLOW_OPERATION_SECONDS = float(os.environ.get("SLOW_OPERATION_SECONDS", "3.0"))

# Word sdef candidate paths
WORD_SDEF_PATHS = [
    sdef_file,
    "/Applications/Microsoft Word.app/Contents/Resources/Word.sdef",
    "/Applications/Microsoft Word.app/Contents/Resources/Word 2019.sdef",
]

# Cache directory for inspector/indexer
from pathlib import Path
CACHE_DIR = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "scriptagent"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

# Default sqlite DB path for sdef inspector/indexer
SDEF_DB_PATH = str(CACHE_DIR / "sdef_index.db")
