"""Unit tests for email formatter (no Outlook required)."""
# Author: Hemprasad Badgujar

import sys
from pathlib import Path

# Add project root so "src" can be imported when run directly: python tests/test_email_formatter.py
_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from datetime import datetime

import pytest

from src.utils.email_formatter import (
    format_mailbox_status,
    format_email_chain,
    group_by_conversation,
    get_date_range,
    get_mailbox_distribution,
    get_participants,
    format_single_email,
    get_importance_text,
    parse_iso_time,
)


def test_format_mailbox_status_success():
    """format_mailbox_status structures access result."""
    result = format_mailbox_status({
        "outlook_connected": True,
        "personal_accessible": True,
        "personal_name": "Me",
        "shared_accessible": True,
        "shared_configured": True,
        "shared_name": "Shared",
        "shared_names": ["Shared"],
        "retention_personal_months": 6,
        "retention_shared_months": 12,
        "errors": [],
    })
    assert result["status"] == "success"
    assert result["connection"]["outlook_connected"] is True
    assert result["personal_mailbox"]["accessible"] is True
    assert result["personal_mailbox"]["name"] == "Me"
    assert result["shared_mailbox"]["accessible"] is True
    assert result["shared_mailbox"]["name"] == "Shared"
    assert result["shared_mailbox"].get("names") == ["Shared"]


def test_format_mailbox_status_error():
    """format_mailbox_status includes errors."""
    result = format_mailbox_status({
        "outlook_connected": False,
        "errors": ["Connection failed"],
    })
    assert result["status"] == "success"
    assert result["errors"] == ["Connection failed"]


def test_group_by_conversation():
    """Emails are grouped by normalized subject."""
    emails = [
        {"subject": "Re: Hello", "body": ""},
        {"subject": "Hello", "body": ""},
        {"subject": "Fwd: Hello", "body": ""},
    ]
    grouped = group_by_conversation(emails)
    assert len(grouped) == 1
    assert "hello" in grouped
    assert len(grouped["hello"]) == 3


def test_get_date_range():
    """get_date_range returns first and last dates."""
    d1 = datetime(2024, 1, 1)
    d2 = datetime(2024, 1, 15)
    emails = [{"received_time": d1}, {"received_time": d2}]
    r = get_date_range(emails)
    assert "first" in r and "last" in r
    assert "2024-01-01" in r["first"]
    assert "2024-01-15" in r["last"]


def test_get_date_range_empty():
    """get_date_range handles empty list."""
    r = get_date_range([])
    assert r["first"] is None and r["last"] is None


def test_get_mailbox_distribution():
    """get_mailbox_distribution counts by mailbox type."""
    emails = [
        {"mailbox_type": "personal"},
        {"mailbox_type": "personal"},
        {"mailbox_type": "shared"},
    ]
    d = get_mailbox_distribution(emails)
    assert d["personal"] == 2
    assert d["shared"] == 1


def test_format_email_chain_empty():
    """format_email_chain with no emails returns no_emails_found."""
    result = format_email_chain([], "test")
    assert result["status"] == "no_emails_found"
    assert result["search_subject"] == "test"


def test_format_email_chain_with_emails():
    """format_email_chain structures conversations."""
    emails = [
        {
            "subject": "Re: Hi",
            "sender_name": "A",
            "sender_email": "a@x.com",
            "recipients": [],
            "received_time": datetime(2024, 1, 2),
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "body": "Hello",
            "attachments_count": 0,
            "importance": 1,
            "unread": False,
            "size": 100,
        },
    ]
    result = format_email_chain(emails, "Hi")
    assert result["status"] == "success"
    assert result["summary"]["total_emails"] == 1
    assert len(result["conversations"]) >= 1


def test_format_single_email():
    """format_single_email builds a single email dict."""
    email = {
        "subject": "Test",
        "sender_name": "X",
        "sender_email": "x@y.com",
        "recipients": ["y@z.com"],
        "folder_name": "Inbox",
        "mailbox_type": "personal",
        "body": "Body text",
        "attachments_count": 0,
        "importance": 2,
        "unread": True,
        "size": 500,
        "received_time": datetime(2024, 1, 1),
    }
    out = format_single_email(email)
    assert out["subject"] == "Test"
    assert out["sender_name"] == "X"
    assert out["importance"] == "High"
    assert out["unread"] is True
    assert out.get("entry_id") is None
    assert out.get("body") == "Body text"
    assert out["body_preview"] == "Body text"


def test_format_single_email_include_body_false():
    """format_single_email with include_body=False omits full body."""
    email = {"subject": "S", "body": "Full body content here"}
    out = format_single_email(email, include_body=False)
    assert out["body"] == ""
    assert out["body_preview"] == "Full body content here"


def test_format_single_email_includes_entry_id():
    """format_single_email includes entry_id when present for reply/forward."""
    email = {
        "subject": "Test",
        "sender_name": "X",
        "entry_id": "ABC123OUTLOOKENTRYID",
        "received_time": datetime(2024, 1, 1),
    }
    out = format_single_email(email)
    assert out["entry_id"] == "ABC123OUTLOOKENTRYID"


def test_get_importance_text():
    """get_importance_text maps levels."""
    assert get_importance_text(0) == "Low"
    assert get_importance_text(1) == "Normal"
    assert get_importance_text(2) == "High"
    assert get_importance_text(99) == "Normal"


def test_parse_iso_time():
    """parse_iso_time parses ISO strings."""
    d = parse_iso_time("2024-01-15T10:30:00")
    assert d.year == 2024 and d.month == 1 and d.day == 15
    assert parse_iso_time("") == datetime.min
    assert parse_iso_time(None) == datetime.min


def test_parse_iso_time_accepts_datetime():
    """parse_iso_time returns datetime unchanged when given datetime."""
    dt = datetime(2024, 2, 20, 12, 0, 0)
    assert parse_iso_time(dt) == dt


def test_get_date_range_mixed_string_and_datetime():
    """get_date_range handles both string and datetime received_time."""
    d1 = datetime(2024, 1, 1)
    emails = [
        {"received_time": "2024-01-15T10:00:00"},
        {"received_time": d1},
    ]
    r = get_date_range(emails)
    assert "first" in r and "last" in r
    assert "2024-01-01" in r["first"]
    assert "2024-01-15" in r["last"]


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
