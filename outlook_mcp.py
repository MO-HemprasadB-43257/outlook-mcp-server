"""Simplified Outlook MCP Server with three main tools."""
# Author: Hemprasad Badgujar

import asyncio
import logging
import platform
import sys
from typing import Any, Optional, Sequence

# Check if running on Windows
if platform.system() != 'Windows':
    print("[ERROR] Outlook MCP Server requires Windows with Microsoft Outlook installed")
    print(f"   Current platform: {platform.system()}")
    print("\n[INFO] To use this server:")
    print("   1. Run on a Windows machine with Outlook installed")
    print("   2. Or use a Windows virtual machine")
    print("   3. Or access a remote Windows desktop")
    sys.exit(1)

from mcp import types
from mcp.server import Server
from mcp.server.stdio import stdio_server

try:
    from src.config.config_reader import config
    from src.tools.outlook_tools import get_tools
    from src.utils.outlook_client import outlook_client
    from src.utils.email_formatter import format_mailbox_status, format_email_chain
except ImportError as e:
    print(f"[ERROR] Import Error: {e}")
    print("\n[INFO] Please install required dependencies:")
    print("   pip install -r requirements.txt")
    print("\nNote: pywin32 is required and only works on Windows")
    sys.exit(1)

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)
# Reduce MCP SDK noise: "Processing request of type X" is INFO; Cursor may show it as [error]. Suppress it.
logging.getLogger("mcp").setLevel(logging.WARNING)


def _coerce_bool(value: Any) -> bool:
    """Coerce MCP argument to bool (handles string 'true'/'false')."""
    if isinstance(value, bool):
        return value
    if value is None:
        return True
    return str(value).lower() not in ("false", "0", "no", "off", "")


# === MCP Server Initialization ===
app = Server("outlook-mcp-server")


@app.list_tools()
async def list_tools() -> list[types.Tool]:
    return get_tools()



# === Tool Handlers ===
async def handle_check_mailbox_access() -> Sequence[types.TextContent]:
    logger.info("Checking mailbox access...")
    try:
        access_result = await asyncio.to_thread(outlook_client.check_access)
        formatted_result = format_mailbox_status(access_result)
        logger.info("Mailbox access check completed")
        return [types.TextContent(type="text", text=str(formatted_result))]
    except Exception as e:
        logger.exception("Error checking mailbox access: %s", e)
        error_response = {
            "status": "error",
            "message": "Could not check mailbox access. Ensure Outlook is running and permission was granted.",
            "troubleshooting": [
                "Make sure Outlook is running",
                "Grant permission when security dialog appears", 
                "Check network connectivity"
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]

async def handle_get_email_chain(
    search_text: str, include_personal: bool, include_shared: bool, include_body: bool = True
) -> Sequence[types.TextContent]:
    logger.info("Searching for emails containing: %s", search_text)
    try:
        emails = await asyncio.to_thread(
            outlook_client.search_emails,
            search_text=search_text,
            include_personal=include_personal,
            include_shared=include_shared,
        )
        formatted_result = format_email_chain(emails, search_text, include_body=include_body)
        logger.info("Found %s emails containing '%s'", len(emails), search_text)
        return [types.TextContent(type="text", text=str(formatted_result))]
    except Exception as e:
        logger.exception("Error searching emails: %s", e)
        error_response = {
            "status": "error",
            "search_text": search_text,
            "message": "Could not search emails. Verify Outlook connection and try again.",
            "troubleshooting": [
                "Verify Outlook connection", 
                "Use specific search terms for best results",
                "Ensure mailboxes are accessible"
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_get_latest_emails(
    count: int, include_personal: bool, include_shared: bool, include_body: bool = True
) -> Sequence[types.TextContent]:
    """Return the N most recent emails from Inbox (no search phrase)."""
    logger.info("Fetching latest %s emails", count)
    try:
        emails = await asyncio.to_thread(
            outlook_client.get_latest_emails,
            count=count,
            include_personal=include_personal,
            include_shared=include_shared,
        )
        formatted_result = format_email_chain(emails, "latest", include_body=include_body)
        logger.info("Returned %s latest emails", len(emails))
        return [types.TextContent(type="text", text=str(formatted_result))]
    except Exception as e:
        logger.exception("Error fetching latest emails: %s", e)
        error_response = {
            "status": "error",
            "message": "Could not fetch latest emails. Verify Outlook connection and try again.",
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_send_email(to: str, subject: str, body: str, cc: Optional[str], bcc: Optional[str]) -> Sequence[types.TextContent]:
    logger.info("Sending email to %s", to)
    try:
        result = await asyncio.to_thread(
            outlook_client.send_email,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
        )
        return [types.TextContent(type="text", text=str(result))]
    except Exception as e:
        logger.exception("Error sending email: %s", e)
        return [types.TextContent(type="text", text=str({
            "status": "error",
            "message": "Could not send email. Verify Outlook is running and try again.",
        }))]


async def handle_reply_to_email(entry_id: str, body: Optional[str], reply_all: bool) -> Sequence[types.TextContent]:
    logger.info("Replying to email %s", entry_id[:50] if entry_id else "")
    try:
        result = await asyncio.to_thread(
            outlook_client.reply_to_email,
            entry_id=entry_id,
            body=body,
            reply_all=reply_all,
        )
        return [types.TextContent(type="text", text=str(result))]
    except Exception as e:
        logger.exception("Error replying to email: %s", e)
        return [types.TextContent(type="text", text=str({
            "status": "error",
            "message": "Could not reply. Verify entry_id from get_email_chain and Outlook connection.",
        }))]


async def handle_forward_email(entry_id: str, to: str, body: Optional[str]) -> Sequence[types.TextContent]:
    logger.info("Forwarding email to %s", to)
    try:
        result = await asyncio.to_thread(
            outlook_client.forward_email,
            entry_id=entry_id,
            to=to,
            body=body,
        )
        return [types.TextContent(type="text", text=str(result))]
    except Exception as e:
        logger.exception("Error forwarding email: %s", e)
        return [types.TextContent(type="text", text=str({
            "status": "error",
            "message": "Could not forward. Verify entry_id from get_email_chain and Outlook connection.",
        }))]


def _safe_str(val: Any) -> str:
    """Coerce value to string for use in .strip() and validation."""
    if val is None:
        return ""
    return str(val)


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> Sequence[types.TextContent]:
    """Dispatch tool calls to appropriate handlers."""
    logger.info("Executing tool: %s", name)
    try:
        if name == "check_mailbox_access":
            return await handle_check_mailbox_access()
        elif name == "get_email_chain":
            search_text = _safe_str(arguments.get("search_text")).strip()
            if not search_text:
                raise ValueError("search_text parameter is required")
            include_personal = _coerce_bool(arguments.get("include_personal", True))
            include_shared = _coerce_bool(arguments.get("include_shared", True))
            include_body = _coerce_bool(arguments.get("include_body", True))
            return await handle_get_email_chain(
                search_text, include_personal, include_shared, include_body
            )
        elif name == "get_latest_emails":
            count_val = arguments.get("count")
            count = int(count_val) if count_val is not None else 10
            count = max(1, min(count, 500))
            include_personal = _coerce_bool(arguments.get("include_personal", True))
            include_shared = _coerce_bool(arguments.get("include_shared", True))
            include_body = _coerce_bool(arguments.get("include_body", True))
            return await handle_get_latest_emails(
                count, include_personal, include_shared, include_body
            )
        elif name == "send_email":
            to = _safe_str(arguments.get("to")).strip()
            subject = _safe_str(arguments.get("subject")).strip()
            body = _safe_str(arguments.get("body")).strip()
            if not to:
                raise ValueError("send_email requires 'to'")
            if not subject:
                raise ValueError("send_email requires 'subject'")
            if not body:
                raise ValueError("send_email requires 'body'")
            cc = arguments.get("cc") if arguments.get("cc") else None
            bcc = arguments.get("bcc") if arguments.get("bcc") else None
            return await handle_send_email(to, subject, body, cc, bcc)
        elif name == "reply_to_email":
            entry_id = _safe_str(arguments.get("entry_id")).strip()
            if not entry_id:
                raise ValueError("reply_to_email requires 'entry_id' from get_email_chain")
            body = arguments.get("body") if arguments.get("body") else None
            reply_all = _coerce_bool(arguments.get("reply_all", False))
            return await handle_reply_to_email(entry_id, body, reply_all)
        elif name == "forward_email":
            entry_id = _safe_str(arguments.get("entry_id")).strip()
            to = _safe_str(arguments.get("to")).strip()
            if not entry_id:
                raise ValueError("forward_email requires 'entry_id' from get_email_chain")
            if not to:
                raise ValueError("forward_email requires 'to'")
            body = arguments.get("body") if arguments.get("body") else None
            return await handle_forward_email(entry_id, to, body)
        else:
            raise ValueError("Unknown tool: " + str(name))
    except ValueError as e:
        logger.warning("Validation error in tool %s: %s", name, e)
        error_response = {
            "status": "error",
            "tool": name,
            "message": str(e),
        }
        return [types.TextContent(type="text", text=str(error_response))]
    except Exception as e:
        logger.exception("Error in tool %s: %s", name, e)
        error_response = {
            "status": "error",
            "tool": name,
            "message": "Tool execution failed. See server logs for details.",
        }
        return [types.TextContent(type="text", text=str(error_response))]


@app.list_resources()
async def list_resources() -> list[types.Resource]:
    """Return list of available resources."""
    return [
        types.Resource(
            uri="outlook-mcp://config",
            name="Current Configuration", 
            description="Show current configuration settings",
            mimeType="text/plain"
        )
    ]


@app.read_resource()
async def read_resource(uri: str) -> str:
    """Read resource content by URI."""
    if uri == "outlook-mcp://config":
        cfg = getattr(config, "config", None)
        if not isinstance(cfg, dict):
            return "Configuration not available.\n"
        lines = ["Configuration (outlook-mcp://config)\n", "=" * 40]
        for key, value in sorted(cfg.items()):
            if key == "shared_mailbox_email" and not value:
                lines.append(f"{key}: <not configured>")
            else:
                lines.append(f"{key}: {value}")
        lines.append("=" * 40)
        return "\n".join(lines)
    raise ValueError("Unknown resource: " + str(uri))


async def main():
    """Main entry point."""
    print("=" * 60)
    print("[STARTING] Outlook MCP Server")
    print("=" * 60)
    
    # Show configuration
    config.show_config()
    
    # Important notes
    print("\n[INFO] Important Notes:")
    print("   * Make sure Microsoft Outlook is running")
    print("   * Grant permission when security dialog appears")  
    print("   * Update config.properties with your shared mailbox details")
    print("   * Server searches ALL folders, not just Inbox")
    
    shared_email = config.get("shared_mailbox_email")
    if not shared_email or "your-shared-mailbox" in str(shared_email).lower() or "example.com" in str(shared_email).lower():
        print("\n[WARNING] Shared mailbox not configured or still using placeholder!")
        print("   Edit src/config/config.properties:")
        print("   - Set shared_mailbox_email to your real shared mailbox (e.g. team@company.com), or")
        print("   - Leave it empty (shared_mailbox_email=) to skip shared mailbox and avoid errors.")
    
    print("\n[TOOLS] Available Tools:")
    print("   1. check_mailbox_access - Test connection and access")
    print("   2. get_email_chain - Search emails by text in subject AND body")
    print("   3. get_latest_emails - Get N most recent Inbox emails (no search phrase)")
    print("   4. send_email - Compose and send a new email")
    print("   5. reply_to_email - Reply (use entry_id from get_email_chain or get_latest_emails)")
    print("   6. forward_email - Forward (use entry_id from get_email_chain or get_latest_emails)")
    
    print("\n[READY] Server ready! Listening for MCP client connections...")
    print("=" * 60)
    
    # Start server
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[INFO] Server stopped by user")
    except Exception as e:
        print("\n[ERROR] Server error:", e)
        logger.error("Server error: %s", e)
