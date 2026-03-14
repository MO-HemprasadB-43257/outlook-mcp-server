"""Unit tests for MCP tools (definitions and call_tool with mocked Outlook)."""
# Author: Hemprasad Badgujar

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

# Add project root so "src" can be imported when run directly
_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import pytest

from src.tools.outlook_tools import get_tools

# Mock Outlook client and platform so we can import outlook_mcp on Linux (no pywin32).
# Must be done before first import of outlook_mcp.
_mock_outlook = MagicMock()
_mock_oc_module = MagicMock()
_mock_oc_module.outlook_client = _mock_outlook

# Cache for outlook_mcp module (set once when first async test runs).
_outlook_mcp_cache = []


def _ensure_outlook_mcp_imported():
    """Import outlook_mcp once with mocks in place (for call_tool tests)."""
    if _outlook_mcp_cache:
        return _outlook_mcp_cache[0]
    with patch("platform.system", return_value="Windows"):
        sys.modules["src.utils.outlook_client"] = _mock_oc_module
        import outlook_mcp as om  # noqa: E402

        _outlook_mcp_cache.append(om)
        return om


# --- Tool list / schema tests (no Outlook mocks) ---


def test_get_tools_returns_six_tools():
    """get_tools() returns exactly 6 tools."""
    tools = get_tools()
    assert len(tools) == 6


def test_get_tools_names():
    """All expected tool names are present."""
    tools = get_tools()
    names = [t.name for t in tools]
    expected = [
        "check_mailbox_access",
        "get_email_chain",
        "get_latest_emails",
        "send_email",
        "reply_to_email",
        "forward_email",
    ]
    assert set(names) == set(expected)


def test_get_tools_check_mailbox_access_schema():
    """check_mailbox_access has no required params."""
    tools = get_tools()
    tool = next(t for t in tools if t.name == "check_mailbox_access")
    assert tool.inputSchema["type"] == "object"
    assert tool.inputSchema.get("required", []) == []


def test_get_tools_get_email_chain_schema():
    """get_email_chain requires search_text."""
    tools = get_tools()
    tool = next(t for t in tools if t.name == "get_email_chain")
    assert "search_text" in tool.inputSchema["properties"]
    assert "search_text" in tool.inputSchema["required"]


def test_get_tools_get_latest_emails_schema():
    """get_latest_emails has no required params (count optional)."""
    tools = get_tools()
    tool = next(t for t in tools if t.name == "get_latest_emails")
    assert "count" in tool.inputSchema["properties"]
    assert tool.inputSchema.get("required", []) == []


def test_get_tools_send_email_schema():
    """send_email requires to, subject, body."""
    tools = get_tools()
    tool = next(t for t in tools if t.name == "send_email")
    required = set(tool.inputSchema["required"])
    assert required == {"to", "subject", "body"}


def test_get_tools_reply_to_email_schema():
    """reply_to_email requires entry_id."""
    tools = get_tools()
    tool = next(t for t in tools if t.name == "reply_to_email")
    assert "entry_id" in tool.inputSchema["required"]


def test_get_tools_forward_email_schema():
    """forward_email requires entry_id and to."""
    tools = get_tools()
    tool = next(t for t in tools if t.name == "forward_email")
    required = set(tool.inputSchema["required"])
    assert required == {"entry_id", "to"}


# --- call_tool tests (mocked Outlook, async) ---


@pytest.fixture(autouse=True)
def reset_mock_outlook():
    """Reset mock return values and side_effects between tests."""
    _mock_outlook.reset_mock()
    _mock_outlook.check_access.return_value = None
    _mock_outlook.check_access.side_effect = None
    _mock_outlook.search_emails.return_value = []
    _mock_outlook.search_emails.side_effect = None
    _mock_outlook.get_latest_emails.return_value = []
    _mock_outlook.get_latest_emails.side_effect = None
    _mock_outlook.send_email.return_value = {}
    _mock_outlook.send_email.side_effect = None
    _mock_outlook.reply_to_email.return_value = {}
    _mock_outlook.reply_to_email.side_effect = None
    _mock_outlook.forward_email.return_value = {}
    _mock_outlook.forward_email.side_effect = None
    yield


@pytest.mark.asyncio
async def test_call_tool_unknown_tool():
    """Unknown tool name returns error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool("unknown_tool_xyz", {})
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "unknown" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_check_mailbox_access_success():
    """check_mailbox_access returns formatted status when client succeeds."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.check_access.return_value = {
        "outlook_connected": True,
        "personal_accessible": True,
        "personal_name": "Test User",
        "shared_accessible": False,
        "shared_configured": False,
        "shared_names": [],
        "retention_personal_months": 6,
        "retention_shared_months": 12,
        "errors": [],
    }
    result = await om.call_tool("check_mailbox_access", {})
    assert len(result) == 1
    text = result[0].text
    assert "success" in text.lower()
    assert "Test User" in text


@pytest.mark.asyncio
async def test_call_tool_check_mailbox_access_error():
    """check_mailbox_access returns error when client raises."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.check_access.side_effect = RuntimeError("Outlook not running")
    result = await om.call_tool("check_mailbox_access", {})
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_get_email_chain_success():
    """get_email_chain returns formatted emails when client returns list."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.search_emails.return_value = [
        {
            "subject": "Test",
            "body": "Body",
            "sender_name": "Alice",
            "sender_email": "alice@example.com",
            "recipients": ["bob@example.com"],
            "folder_name": "Inbox",
            "mailbox_type": "personal",
            "received_time": "2024-01-15T10:00:00",
            "entry_id": "abc123",
        }
    ]
    result = await om.call_tool(
        "get_email_chain",
        {"search_text": "test", "include_personal": True, "include_shared": True},
    )
    assert len(result) == 1
    text = result[0].text
    assert "success" in text.lower() or "Test" in text


@pytest.mark.asyncio
async def test_call_tool_get_email_chain_missing_search_text():
    """get_email_chain with empty search_text returns validation error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool("get_email_chain", {"search_text": ""})
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "search_text" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_get_latest_emails_success():
    """get_latest_emails returns formatted result when client returns list."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.get_latest_emails.return_value = []
    result = await om.call_tool(
        "get_latest_emails",
        {"count": 5, "include_personal": True, "include_shared": False},
    )
    assert len(result) == 1
    text = result[0].text
    assert "no_emails_found" in text or "latest" in text or "success" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_send_email_success():
    """send_email returns status when client succeeds."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.send_email.return_value = {
        "status": "sent",
        "to": "a@b.com",
        "subject": "Hi",
    }
    result = await om.call_tool(
        "send_email",
        {"to": "a@b.com", "subject": "Hi", "body": "Hello"},
    )
    assert len(result) == 1
    text = result[0].text
    assert "sent" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_send_email_validation_missing_to():
    """send_email without 'to' returns validation error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool(
        "send_email",
        {"to": "", "subject": "Hi", "body": "Hello"},
    )
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "to" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_send_email_validation_missing_subject():
    """send_email without 'subject' returns validation error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool(
        "send_email",
        {"to": "a@b.com", "subject": "", "body": "Hello"},
    )
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "subject" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_send_email_validation_missing_body():
    """send_email without 'body' returns validation error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool(
        "send_email",
        {"to": "a@b.com", "subject": "Hi", "body": ""},
    )
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "body" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_reply_to_email_validation_missing_entry_id():
    """reply_to_email without entry_id returns validation error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool("reply_to_email", {"entry_id": ""})
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "entry_id" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_reply_to_email_success():
    """reply_to_email returns status when client succeeds."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.reply_to_email.return_value = {
        "status": "sent",
        "action": "reply",
        "entry_id": "abc",
    }
    result = await om.call_tool(
        "reply_to_email",
        {"entry_id": "some-entry-id", "body": "Thanks", "reply_all": False},
    )
    assert len(result) == 1
    text = result[0].text
    assert "sent" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_forward_email_validation_missing_entry_id():
    """forward_email without entry_id returns validation error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool("forward_email", {"entry_id": "", "to": "x@y.com"})
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "entry_id" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_forward_email_validation_missing_to():
    """forward_email without 'to' returns validation error."""
    om = _ensure_outlook_mcp_imported()
    result = await om.call_tool(
        "forward_email",
        {"entry_id": "some-id", "to": ""},
    )
    assert len(result) == 1
    text = result[0].text
    assert "error" in text.lower()
    assert "to" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_forward_email_success():
    """forward_email returns status when client succeeds."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.forward_email.return_value = {
        "status": "sent",
        "action": "forward",
        "to": "b@c.com",
        "entry_id": "abc",
    }
    result = await om.call_tool(
        "forward_email",
        {"entry_id": "some-id", "to": "b@c.com", "body": "FYI"},
    )
    assert len(result) == 1
    text = result[0].text
    assert "sent" in text.lower()


@pytest.mark.asyncio
async def test_call_tool_get_email_chain_include_body_false():
    """get_email_chain with include_body=false passes through."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.search_emails.return_value = []
    result = await om.call_tool(
        "get_email_chain",
        {"search_text": "x", "include_body": False},
    )
    assert len(result) == 1


@pytest.mark.asyncio
async def test_call_tool_coerce_bool_string_false():
    """Boolean params accept string 'false'."""
    om = _ensure_outlook_mcp_imported()
    _mock_outlook.search_emails.return_value = []
    result = await om.call_tool(
        "get_email_chain",
        {"search_text": "x", "include_personal": "false", "include_shared": "true"},
    )
    assert len(result) == 1
    _mock_outlook.search_emails.assert_called_once()
    call_kw = _mock_outlook.search_emails.call_args[1]
    assert call_kw["include_personal"] is False
    assert call_kw["include_shared"] is True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
