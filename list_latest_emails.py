"""
Fetch and print actual latest emails from Outlook (real data, not dummy).
Requires: Windows, Outlook running, pip install -r requirements.txt
Run from project root: python list_latest_emails.py [count]
"""
# Author: Hemprasad Badgujar

import json
import sys
import platform

if platform.system() != "Windows":
    print("[ERROR] This script requires Windows with Microsoft Outlook installed.")
    sys.exit(1)

# Project root on path
from pathlib import Path

_root = Path(__file__).resolve().parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

try:
    from src.config.config_reader import config
    from src.utils.outlook_client import outlook_client
    from src.utils.email_formatter import format_email_chain, format_email_chain_pretty_text
except ImportError as e:
    print(f"[ERROR] {e}")
    print("Run: pip install -r requirements.txt")
    sys.exit(1)


def _serialize(val):
    """Make value JSON-serializable (e.g. datetime -> isoformat)."""
    if hasattr(val, "isoformat") and callable(getattr(val, "isoformat")):
        return val.isoformat()
    if isinstance(val, dict):
        return {k: _serialize(v) for k, v in val.items()}
    if isinstance(val, (list, tuple)):
        return [_serialize(v) for v in val]
    return val


def main():
    count = 10
    if len(sys.argv) > 1:
        try:
            count = max(1, min(int(sys.argv[1]), 50))
        except ValueError:
            pass

    # Avoid "Extended MAPI login failed" on some Outlook setups (optional feature)
    _orig_get_bool = config.get_bool
    config.get_bool = lambda key, default=True: False if key == "use_extended_mapi_login" else _orig_get_bool(key, default)

    print("Connecting to Outlook and fetching latest", count, "emails...")
    print("(Allow Outlook if a security prompt appears.)\n")

    try:
        emails = outlook_client.get_latest_emails(
            count=count,
            include_personal=True,
            include_shared=True,
        )
    except Exception as e:
        print("[ERROR] Failed to get latest emails:", e)
        print("Ensure Outlook is running and you have granted access.")
        sys.exit(1)

    formatted = format_email_chain(emails, "latest", include_body=True)
    print("=== Latest emails ===\n")
    print(format_email_chain_pretty_text(formatted))
    if "--json" in sys.argv:
        print("\n--- full JSON ---")
        print(json.dumps(_serialize(formatted), indent=2, default=str))


if __name__ == "__main__":
    main()
