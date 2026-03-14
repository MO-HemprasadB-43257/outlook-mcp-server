"""Unit tests for get_latest_emails MCP tool (schema and call_tool with mocked Outlook)."""
# Author: Hemprasad Badgujar

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import pytest

from src.tools.outlook_tools import get_tools

# Mock Outlook client and platform so we can import outlook_mcp on Linux (no pywin32).
_mock_outlook = MagicMock()
_mock_oc_module = MagicMock()
_mock_oc_module.outlook_client = _mock_outlook
_outlook_mcp_cache = []


def _ensure_outlook_mcp_imported():
    """Import outlook_mcp once with mocks in place."""
    if _outlook_mcp_cache:
        return _outlook_mcp_cache[0]
    with patch("platform.system", return_value="Windows"):
        sys.modules["src.utils.outlook_client"] = _mock_oc_module
        import outlook_mcp as om  # noqa: E402

        _outlook_mcp_cache.append(om)
        return om


# --- Schema tests ---


def test_get_latest_emails_tool_schema():
    """get_latest_emails has no required params; count, include_personal, include_shared, include_body optional."""
    tools = get_tools()
    tool = next(t for t in tools if t.name == "get_latest_emails")
    assert tool.inputSchema["type"] == "object"
    assert tool.inputSchema.get("required", []) == []
    props = tool.inputSchema["properties"]
    assert "count" in props
    assert "include_personal" in props
    assert "include_shared" in props
    assert "include_body" in props


# --- call_tool tests (mocked Outlook) ---


@pytest.fixture(autouse=True)
def reset_mock_outlook():
    """Reset mock between tests."""
    _mock_outlook.reset_mock()
    _mock_outlook.get_latest_emails.return_value = []
    _mock_outlook.get_latest_emails.side_effect = None
    yield


@pytest.mark.asyncio
async def test_get_latest_emails_success_empty():
    """get_latest_emails returns formatted result when client returns empty list."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.get_latest_emails.return_value = []
    result = await om.call_tool(
        "get_latest_emails",
        {"count": 5, "include_personal": True, "include_shared": False},
    )
    assert len(result) == 1
    text = result[0].text
    assert "no_emails_found" in text or "latest" in text or "success" in text.lower()
    _mock_outlook.get_latest_emails.assert_called_once_with(
        count=5, include_personal=True, include_shared=False
    )


@pytest.mark.asyncio
async def test_get_latest_emails_success_with_emails():
    """get_latest_emails returns formatted conversations when client returns emails."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.get_latest_emails.return_value = [
        {
            "subject": "Latest Inbox",
            "body": "Body text",
            "sender_name": "Alice",
            "sender_email": "alice@example.com",
            "recipients": ["bob@example.com"],
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "received_time": "2024-01-15T10:00:00",
            "entry_id": "entry-123",
        }
    ]
    result = await om.call_tool(
        "get_latest_emails",
        {"count": 10, "include_personal": True, "include_shared": True},
    )
    assert len(result) == 1
    text = result[0].text
    assert "success" in text.lower()
    assert "Latest Inbox" in text or "alice" in text.lower()
    _mock_outlook.get_latest_emails.assert_called_once_with(
        count=10, include_personal=True, include_shared=True
    )


@pytest.mark.asyncio
async def test_get_latest_emails_returns_result_structure():
    """get_latest_emails result includes status, summary, conversations, and all_emails_chronological."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.get_latest_emails.return_value = [
        {
            "subject": "Report Q1",
            "body": "Content here",
            "sender_name": "Bob",
            "sender_email": "bob@example.com",
            "recipients": ["team@example.com"],
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "received_time": "2024-03-01T09:00:00",
            "entry_id": "id-456",
        }
    ]
    result = await om.call_tool(
        "get_latest_emails",
        {"count": 5, "include_personal": True, "include_shared": False},
    )
    assert len(result) == 1
    text = result[0].text
    # Result is str(dict) from handler; assert it contains expected structure keys
    assert "status" in text
    assert "success" in text.lower()
    assert "summary" in text
    assert "conversations" in text
    assert "all_emails_chronological" in text
    assert "total_emails" in text or "search_subject" in text
    # Also returns the email content in the result
    assert "Report Q1" in text or "bob" in text.lower()
    assert "entry_id" in text or "id-456" in text


@pytest.mark.asyncio
async def test_get_latest_emails_default_count():
    """get_latest_emails uses default count when not provided."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool("get_latest_emails", {})
    assert len(result) == 1
    _mock_outlook.get_latest_emails.assert_called_once()
    call_kw = _mock_outlook.get_latest_emails.call_args[1]
    assert call_kw["count"] == 10


@pytest.mark.asyncio
async def test_get_latest_emails_include_body_false():
    """get_latest_emails passes include_body to formatter (metadata-only)."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.get_latest_emails.return_value = [
        {
            "subject": "Test",
            "body": "Full body here",
            "sender_name": "S",
            "sender_email": "s@e.com",
            "recipients": [],
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "received_time": "2024-01-01T00:00:00",
            "entry_id": "id1",
        }
    ]
    result = await om.call_tool(
        "get_latest_emails",
        {"count": 1, "include_body": False},
    )
    assert len(result) == 1
    # Formatter with include_body=False yields empty body in formatted email
    text = result[0].text
    assert "success" in text.lower()
    _mock_outlook.get_latest_emails.assert_called_once_with(
        count=1, include_personal=True, include_shared=True
    )


@pytest.mark.asyncio
async def test_get_latest_emails_client_error():
    """get_latest_emails returns error when client raises."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.get_latest_emails.side_effect = RuntimeError("Outlook not connected")
    result = await om.call_tool(
        "get_latest_emails",
        {"count": 5},
    )
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()


@pytest.mark.asyncio
async def test_get_latest_emails_output_lists_all_emails():
    """get_latest_emails output contains the actual list of latest emails (subject, sender, entry_id each)."""
    om = _ensure_outlook_mcp_imported()
    latest_emails = [
        {
            "subject": "Re: Project Alpha",
            "body": "See attached.",
            "sender_name": "Alice",
            "sender_email": "alice@example.com",
            "recipients": ["bob@example.com"],
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "received_time": "2024-03-10T14:00:00",
            "entry_id": "entry-001",
        },
        {
            "subject": "Q1 Report",
            "body": "Summary inside.",
            "sender_name": "Bob",
            "sender_email": "bob@example.com",
            "recipients": ["team@example.com"],
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "received_time": "2024-03-10T12:30:00",
            "entry_id": "entry-002",
        },
        {
            "subject": "Meeting tomorrow",
            "body": "10am room B.",
            "sender_name": "Carol",
            "sender_email": "carol@example.com",
            "recipients": ["alice@example.com"],
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "received_time": "2024-03-10T09:00:00",
            "entry_id": "entry-003",
        },
    ]
    _mock_outlook.get_latest_emails.return_value = latest_emails
    result = await om.call_tool(
        "get_latest_emails",
        {"count": 10, "include_personal": True, "include_body": True},
    )
    assert len(result) == 1
    text = result[0].text
    # Output must contain the actual list: all 3 subjects
    assert "Re: Project Alpha" in text or "Project Alpha" in text
    assert "Q1 Report" in text
    assert "Meeting tomorrow" in text
    # All 3 senders in output
    assert "Alice" in text
    assert "Bob" in text
    assert "Carol" in text
    # All 3 entry_ids in output (for reply/forward)
    assert "entry-001" in text
    assert "entry-002" in text
    assert "entry-003" in text
    # List structure in output
    assert "all_emails_chronological" in text
    assert "total_emails" in text
    assert "3" in text  # total count


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
