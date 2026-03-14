"""MCP tool definitions for Outlook (check_mailbox_access, get_email_chain)."""

from mcp import types


def get_tools() -> list[types.Tool]:
    """Return list of available MCP tools."""
    return [
        types.Tool(
            name="check_mailbox_access",
            description="Check connection status and access to personal and shared mailboxes with retention policy info",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
            },
        ),
        types.Tool(
            name="get_email_chain",
            description="Searches for emails containing the specified text in BOTH subject and body using exact phrase matching. Retrieves complete email chains with full email bodies for comprehensive analysis. Searches ALL folders in both personal and shared mailboxes. Returns full email content including sender, recipients, timestamps, and complete message bodies. Use specific search terms (error codes, alert identifiers, unique phrases) for best results.",
            inputSchema={
                "type": "object",
                "properties": {
                    "search_text": {
                        "type": "string",
                        "description": "Exact text pattern to search for in email subject and body. The search looks for this exact phrase.",
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
                        "description": "Include full email body in the response (default: true). Set false for metadata-only / smaller payload.",
                        "default": True,
                    },
                },
                "required": ["search_text"],
            },
        ),
        types.Tool(
            name="get_latest_emails",
            description="Returns the N most recent emails from Inbox (personal and optionally shared) without any search phrase. Use this for 'last email', 'latest 10 emails', or when you need the most recent messages regardless of content. Each email includes entry_id for reply/forward.",
            inputSchema={
                "type": "object",
                "properties": {
                    "count": {
                        "type": "integer",
                        "description": "Number of most recent emails to return (default 10, max from config).",
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
                        "description": "Include full email body in the response (default: true). Set false for metadata-only / smaller payload.",
                        "default": True,
                    },
                },
                "required": [],
            },
        ),
        types.Tool(
            name="send_email",
            description="Compose and send a new email from the default Outlook account. Recipients can be comma-separated.",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Recipient email address(es), comma-separated for multiple"},
                    "subject": {"type": "string", "description": "Email subject"},
                    "body": {"type": "string", "description": "Email body (plain text)"},
                    "cc": {"type": "string", "description": "CC recipients, comma-separated (optional)"},
                    "bcc": {"type": "string", "description": "BCC recipients, comma-separated (optional)"},
                },
                "required": ["to", "subject", "body"],
            },
        ),
        types.Tool(
            name="reply_to_email",
            description="Reply to an email. Use entry_id from get_email_chain results. Optional body; reply_all to reply to all recipients.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {"type": "string", "description": "Entry ID of the email to reply to (from get_email_chain)"},
                    "body": {"type": "string", "description": "Reply body to add (optional)"},
                    "reply_all": {"type": "boolean", "description": "If true, reply to all recipients (default: false)", "default": False},
                },
                "required": ["entry_id"],
            },
        ),
        types.Tool(
            name="forward_email",
            description="Forward an email. Use entry_id from get_email_chain results. Required: to. Optional body (e.g. FYI) prepended.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {"type": "string", "description": "Entry ID of the email to forward (from get_email_chain)"},
                    "to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                    "body": {"type": "string", "description": "Optional message to prepend (e.g. FYI)"},
                },
                "required": ["entry_id", "to"],
            },
        ),
    ]
