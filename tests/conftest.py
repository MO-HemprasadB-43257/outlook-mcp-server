"""Pytest configuration: add project root to path."""
# Author: Hemprasad Badgujar

import sys
from pathlib import Path

root = Path(__file__).resolve().parent.parent
if str(root) not in sys.path:
    sys.path.insert(0, str(root))
