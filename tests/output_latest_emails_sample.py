# Author: Hemprasad Badgujar
"""
Print sample "latest emails" formatted output (no Outlook required).
Run from project root: python tests/output_latest_emails_sample.py
Shows the actual list structure that get_latest_emails returns.
"""

import sys
from pathlib import Path

_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from src.utils.email_formatter import format_email_chain

# Sample latest emails (same shape as outlook_client.get_latest_emails)
SAMPLE_LATEST_EMAILS = [
    {
        "subject": "Re: Project Alpha",
        "body": "Please find the attached document. Let me know if you have questions.",
        "sender_name": "Alice Smith",
        "sender_email": "alice@example.com",
        "recipients": ["bob@example.com"],
        "folder_name": "Inbox",
        "mailbox_type": "personal",
        "received_time": "2024-03-10T14:00:00",
        "entry_id": "AAMkAGI2...",
    },
    {
        "subject": "Q1 Report",
        "body": "Summary: we hit 120% of target. Details in the deck.",
        "sender_name": "Bob Jones",
        "sender_email": "bob@example.com",
        "recipients": ["team@example.com"],
        "folder_name": "Inbox",
        "mailbox_type": "personal",
        "received_time": "2024-03-10T12:30:00",
        "entry_id": "AAMkAGI2...",
    },
    {
        "subject": "Meeting tomorrow 10am",
        "body": "Room B. Agenda: roadmap review.",
        "sender_name": "Carol Lee",
        "sender_email": "carol@example.com",
        "recipients": ["alice@example.com", "bob@example.com"],
        "folder_name": "Inbox",
        "mailbox_type": "personal",
        "received_time": "2024-03-10T09:00:00",
        "entry_id": "AAMkAGI2...",
    },
]


def main():
    formatted = format_email_chain(SAMPLE_LATEST_EMAILS, "latest", include_body=True)
    print("=== get_latest_emails output (sample) ===\n")
    print(str(formatted))
    print("\n=== end of output ===")


if __name__ == "__main__":
    main()
    sys.exit(0)
