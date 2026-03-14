"""MCP tool definitions for Outlook (check_mailbox_access, get_email_chain)."""
# Author: Hemprasad Badgujar

from mcp import types


def get_tools() -> list[types.Tool]:
    """Return list of available MCP tools."""
    return [
        types.Tool(
            name="check_mailbox_access",
            description=(
                "Check connection status and access to personal and shared "
                "mailboxes with retention policy info."
            ),
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
            },
        ),
        types.Tool(
            name="get_email_chain",
            description=(
                "Searches for emails containing the specified text in BOTH "
                "subject and body using exact phrase matching. Retrieves "
                "complete email chains with full bodies. Searches ALL folders "
                "in personal and shared mailboxes. Use specific search terms "
                "(error codes, alert IDs, unique phrases) for best results."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "search_text": {
                        "type": "string",
                        "description": (
                            "Exact text pattern to search for in subject and "
                            "body. The search looks for this exact phrase."
                        ),
                    },
                    "include_personal": {
                        "type": "boolean",
                        "description": "Search personal mailbox (default: true)",
                        "default": True,
                    },
                    "include_shared": {
                        "type": "boolean",
                        "description": "Search shared mailbox (default: true)",
                        "default": True,
                    },
                    "include_body": {
                        "type": "boolean",
                        "description": (
                            "Include full email body (default: true). "
                            "Set false for metadata-only / smaller payload."
                        ),
                        "default": True,
                    },
                },
                "required": ["search_text"],
            },
        ),
        types.Tool(
            name="get_latest_emails",
            description=(
                "Returns the N most recent emails from Inbox (personal and "
                "optionally shared) without any search phrase. Use for 'last "
                "email', 'latest 10 emails'. Each email includes entry_id "
                "for reply/forward."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "count": {
                        "type": "integer",
                        "description": (
                        "Number of most recent emails to return "
                        "(default 10, max from config)."
                    ),
                        "default": 10,
                    },
                    "include_personal": {
                        "type": "boolean",
                        "description": "Include personal mailbox Inbox (default: true)",
                        "default": True,
                    },
                    "include_shared": {
                        "type": "boolean",
                        "description": "Include shared mailbox Inbox (default: true)",
                        "default": True,
                    },
                    "include_body": {
                        "type": "boolean",
                        "description": (
                            "Include full email body (default: true). "
                            "Set false for metadata-only / smaller payload."
                        ),
                        "default": True,
                    },
                },
                "required": [],
            },
        ),
        types.Tool(
            name="send_email",
            description=(
                "Compose and send a new email from the default Outlook "
                "account. Recipients can be comma-separated."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {
                        "type": "string",
                        "description": "Recipient email(s), comma-separated",
                    },
                    "subject": {"type": "string", "description": "Email subject"},
                    "body": {"type": "string", "description": "Email body (plain text)"},
                    "cc": {"type": "string", "description": "CC (optional)"},
                    "bcc": {"type": "string", "description": "BCC (optional)"},
                },
                "required": ["to", "subject", "body"],
            },
        ),
        types.Tool(
            name="reply_to_email",
            description=(
                "Reply to an email. Use entry_id from get_email_chain. "
                "Optional body; reply_all to reply to all."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "Entry ID to reply to (from get_email_chain)",
                    },
                    "body": {"type": "string", "description": "Reply body (optional)"},
                    "reply_all": {
                        "type": "boolean",
                        "description": "Reply to all (default: false)",
                        "default": False,
                    },
                },
                "required": ["entry_id"],
            },
        ),
        types.Tool(
            name="forward_email",
            description=(
                "Forward an email. Use entry_id from get_email_chain. "
                "Required: to. Optional body (e.g. FYI) prepended."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "Entry ID to forward (from get_email_chain)",
                    },
                    "to": {
                        "type": "string",
                        "description": "Recipient email(s), comma-separated",
                    },
                    "body": {"type": "string", "description": "Optional message to prepend (e.g. FYI)"},
                },
                "required": ["entry_id", "to"],
            },
        ),
    ]
