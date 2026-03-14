<p align="center">
  <img src="https://img.shields.io/badge/Outlook-MCP%20Server-0078D4?style=for-the-badge&logo=microsoft-outlook&logoColor=white" alt="Outlook MCP Server" />
</p>

<h1 align="center">📬 Outlook MCP Server</h1>

<p align="center">
  <strong>Bring your Outlook mailboxes into the Model Context Protocol.</strong><br/>
  Search, analyze, and surface email chains for AI assistants—with support for personal and shared mailboxes, parallel search, and near-instant AdvancedSearch.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.8+" />
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?style=flat-square&logo=windows&logoColor=white" alt="Windows" />
  <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License MIT" />
  <img src="https://img.shields.io/badge/MCP-1.0+-purple?style=flat-square" alt="MCP 1.0+" />
</p>

---

## 📑 Table of Contents

- [✨ Features](#-features)
- [📋 Requirements](#-requirements)
- [🚀 Quick Start](#-quick-start)
- [🧪 How to Test](#-how-to-test)
- [⚙️ Configuration](#️-configuration)
- [🔧 Tools](#-tools)
- [🔍 Search & Performance](#-search--performance)
- [🔌 MCP Configuration](#-mcp-configuration)
- [📇 Enable Outlook Search (AdvancedSearch)](#-enable-outlook-search-advancedsearch)
- [🔧 Troubleshooting](#-troubleshooting)
- [📁 Project Structure](#-project-structure)
- [🔒 Security](#-security)
- [🤝 Contributing](#-contributing)
- [👤 Author](#-author)

---

## ✨ Features

| Area | Capability |
|------|-------------|
| **Mailboxes** | Personal + multiple shared mailboxes; configurable list or single |
| **Search** | Exact phrase in subject and body; AdvancedSearch API; optional cross-folder |
| **Performance** | Parallel mailbox search, configurable cache (TTL + size), batch processing |
| **Reliability** | Retry with backoff + jitter, fallback if indexing disabled, health-check tool |
| **Server** | Non-blocking async, config reload, structured errors |

- **Multi-Mailbox** — Access personal inbox and one or more shared mailboxes.
- **AdvancedSearch** — Uses Outlook's index for fast subject + body search.
- **Parallel Search** — Personal and shared mailboxes searched concurrently.
- **Full Chains** — Returns full email bodies and conversation grouping.
- **Configurable** — Cache, workers, batch size, retries, and more in `config.properties`.
- **Health Check** — `check_mailbox_access` doubles as a connectivity health check.

---

## 📋 Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10 or 11 |
| **Outlook** | Desktop app (not Outlook on the web) |
| **Python** | 3.8+ |
| **Packages** | `mcp`, `pywin32` (see `requirements.txt`) |

---

## 🚀 Quick Start

```bash
# 1. Clone and enter project
git clone <repository-url>
cd Outlook-MCP-Server

# 2. Install dependencies
pip install -r requirements.txt

# 3. Configure (edit shared mailbox and options)
#    File: src/config/config.properties
#    Run from project root so config is found.

# 4. Start the server (with Outlook running)
python outlook_mcp.py
```

**Config path:** Loaded from `src/config/config.properties` or from the current working directory. Run from the project root: `python outlook_mcp.py`.

---

## 🧪 How to Test

### 1. Unit tests (no Outlook required)

Runs on any OS (e.g. Linux in CI). Tests config loading and email formatting only. Run from the **project root** (conftest.py adds the root to the path).

```bash
cd Outlook-MCP-Server
pip install -r requirements-ci.txt
python -m pytest tests/test_config_reader.py tests/test_email_formatter.py -v
```

If you run from another directory, set `PYTHONPATH` to the project root (e.g. `set PYTHONPATH=.` on Windows cmd, or `$env:PYTHONPATH="."` on PowerShell).

### 2. Connection test (Windows + Outlook required)

Verifies that the server can connect to Outlook and access mailboxes. **Requires Windows with Microsoft Outlook running.**

```bash
python tests/test_connection.py
```

When prompted, allow Outlook permission if a security dialog appears. You can enter a search term to test email search, or press Enter to use the default.

### 3. Lint and type check (optional)

```bash
pip install -r requirements-dev.txt
python -m flake8 src outlook_mcp.py
python -m mypy src outlook_mcp.py
```

### 4. CI

On push/PR, GitHub Actions runs:

- **Lint:** flake8 and mypy on `src` and `outlook_mcp.py`
- **Tests:** pytest on `test_config_reader.py` and `test_email_formatter.py` (using `requirements-ci.txt`, so no pywin32 on Linux)

Full end-to-end testing (Outlook connection and search) must be done manually on Windows with Outlook installed.

---

## ⚙️ Configuration

All options live in `config.properties`. Reload at runtime with `config.reload()`.

### Mailbox

| Key | Description | Default |
|-----|-------------|---------|
| `shared_mailbox_email` | Single shared mailbox address | — |
| `shared_mailbox_emails` | Multiple shared mailboxes (comma-separated) | — |
| `shared_mailbox_name` | Display name for shared mailbox | Shared Mailbox |

### Search

| Key | Description | Default |
|-----|-------------|---------|
| `max_search_results` | Max emails returned per search | 50 |
| `max_body_chars` | Max characters from body (0 = no limit) | 0 |
| `max_search_body_chars` | Body limit during search (performance) | — |
| `search_all_folders` | Include Sent/Drafts etc. | false |

### Performance & cache

| Key | Description | Default |
|-----|-------------|---------|
| `batch_processing_size` | Batch size for processing results | 50 |
| `max_retry_attempts` | Connection retry attempts | 3 |
| `parallel_search_workers` | Max worker threads for parallel search | 2 |
| `search_cache_ttl_seconds` | Search cache TTL | 3600 |
| `search_cache_max_entries` | Max cache entries | 100 |
| `profile_search` | Log search duration and result count | false |
| `max_recipients_display` | Max recipients per email in output | 10 |

### Other

| Key | Description |
|-----|-------------|
| `connection_timeout_minutes` | Outlook connection timeout |
| `personal_retention_months` | Informational retention (personal) |
| `shared_retention_months` | Informational retention (shared) |
| `use_extended_mapi_login` | Try Extended MAPI to reduce prompts |
| `include_timestamps` | Include timestamps in formatted output |
| `clean_html_content` | Strip HTML from bodies |

---

## 🔧 Tools

### 1. `check_mailbox_access` (health check)

Verifies connection to Outlook and access to configured mailboxes.

| | |
|---|---|
| **Parameters** | None |
| **Returns** | Connection status, personal/shared accessibility, names, retention info, errors |

Use as a **health check** to confirm the server can reach Outlook and shared mailboxes.

---

### 2. `get_email_chain`

Searches for emails containing the given text in **subject and body**, returns full chains.

| | |
|---|---|
| **Parameters** | `search_text` (required), `include_personal`, `include_shared`, `include_body` (default: true — set false for metadata-only / smaller response) |
| **Returns** | Conversations, senders/recipients, timestamps, summary stats, **entry_id** per email. Full **body** included when `include_body` is true; otherwise `body_preview` (500 chars) only. |

**Example request:**

```json
{
  "tool": "get_email_chain",
  "arguments": {
    "search_text": "server error 500",
    "include_personal": true,
    "include_shared": true,
    "include_body": true
  }
}
```

Each email in the response includes **entry_id**. Use it with `reply_to_email` and `forward_email`.

---

### 3. `get_latest_emails`

Returns the **N most recent emails** from Inbox (no search phrase). Use for "last email", "latest 10 emails", or when you need the most recent messages regardless of content.

| | |
|---|---|
| **Parameters** | `count` (optional, default 10), `include_personal`, `include_shared`, `include_body` (default: true — set false for metadata-only / smaller response) |
| **Returns** | Same structure as `get_email_chain`: conversations, **entry_id** per email. Full **body** when `include_body` is true; otherwise `body_preview` only. |

**Example:** `get_latest_emails` with `count: 1` returns your single most recent Inbox email. Use `get_email_chain` when you need to search by a specific phrase.

---

### 4. `send_email`

Compose and send a new email from your default Outlook account. **Sends immediately** (no draft). Send, reply, and forward tools send **real emails**; use with care.

| | |
|---|---|
| **Parameters** | `to` (required), `subject` (required), `body` (required), `cc` (optional), `bcc` (optional). Multiple recipients: comma-separated. |
| **Returns** | `{"status": "sent", "to": "...", "subject": "..."}` or error |

---

### 5. `reply_to_email`

Reply to an email. Use **entry_id** from `get_email_chain` or `get_latest_emails`.

| | |
|---|---|
| **Parameters** | `entry_id` (required), `body` (optional), `reply_all` (optional, default false) |
| **Returns** | Status dict or error |

---

### 6. `forward_email`

Forward an email. Use **entry_id** from `get_email_chain` or `get_latest_emails`.

| | |
|---|---|
| **Parameters** | `entry_id` (required), `to` (required), `body` (optional, e.g. "FYI") |
| **Returns** | Status dict or error |

---

### Other possible features (future)

Ideas for later: list_folders, move_email, mark_read/mark_unread, delete_email, get_calendar_events, create_calendar_event, add_attachment, get_contacts, draft_email (create draft without sending).

---

## 🔍 Search & Performance

- **AdvancedSearch** — Uses Outlook's index; subject and body; case-insensitive phrase match; typically sub-second to a few seconds.
- **Fallback** — If AdvancedSearch fails (e.g. indexing off), falls back to subject-only Restrict and optional iteration.
- **Parallel search** — Personal and each shared mailbox can be searched in parallel (configurable workers).
- **Caching** — Results cached by query and mailbox selection; TTL and max entries configurable.
- **Connection** — Tries existing Outlook instance first; exponential backoff + jitter on retry.
- **Non-blocking** — All Outlook work runs via `asyncio.to_thread()` so the MCP server stays responsive.

**Tips:** Enable Outlook indexing; use specific search terms; set a reasonable `max_search_results`; use `profile_search=true` to tune.

---

## 🔌 MCP Configuration

Add this server to your MCP client so tools (check_mailbox_access, get_email_chain, get_latest_emails, send_email, reply_to_email, forward_email) are available. Use the **full path** to your Python executable and to `outlook_mcp.py`. Outlook must be running on the same Windows machine.

### Claude Desktop

Edit the MCP config file:

- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

Add the `outlook` server under `mcpServers`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\Users\\YourName\\Documents\\GitHub\\Outlook-MCP-Server\\outlook_mcp.py"]
    }
  }
}
```

Use your actual path and double backslashes (`\\`) in JSON. Restart Claude Desktop after saving.

### Cursor

In Cursor: **Settings → Cursor Settings → MCP** (or edit the MCP config file directly). Add:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\Users\\YourName\\Documents\\GitHub\\Outlook-MCP-Server\\outlook_mcp.py"]
    }
  }
}
```

### Generic MCP config (any client)

Any MCP client that supports stdio transport can use this server. Put the server entry in your client's MCP config:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\YourName\\AppData\\Local\\Programs\\Python\\Python311\\python.exe",
      "args": ["C:\\Users\\YourName\\Documents\\GitHub\\Outlook-MCP-Server\\outlook_mcp.py"]
    }
  }
}
```

- **command** — Full path to `python.exe` (e.g. from `where python` on Windows).
- **args** — Array with the full path to `outlook_mcp.py` in this repo.

The server runs over stdio; no extra env vars are required. Configure mailboxes and options in `src/config/config.properties` in the repo folder.

---

## 📇 Enable Outlook search (AdvancedSearch)

If you see "AdvancedSearch unavailable" in logs, the server uses a fallback and still returns results. To use the faster index-based search (AdvancedSearch), enable Outlook indexing:

### 1. Windows Search service

1. Press **Win + R**, type `services.msc`, press Enter.
2. Find **Windows Search** in the list.
3. Right-click → **Properties**.
4. Set **Startup type** to **Automatic**.
5. Click **Start** if the service is not running, then **OK**.

### 2. Outlook indexing

1. Open **Microsoft Outlook**.
2. **File** → **Options**.
3. Go to **Search** (left side).
4. Click **Indexing Options** (or **Indexing**).
5. Ensure **Microsoft Outlook** (and your mail profile) is in the list of indexed locations.
6. If Outlook is missing, click **Modify** and check **Microsoft Outlook** / your mailbox, then **OK**.
7. Click **Indexing Options** again and wait until indexing completes (may take a while for large mailboxes).

### 3. Rebuild index (if search still fails)

1. In **Indexing Options**, click **Advanced**.
2. Under **Troubleshooting**, click **Rebuild**.
3. Confirm and wait for the rebuild to finish (can take 15+ minutes).
4. Restart Outlook and try the MCP server again.

### 4. Restart and test

- Restart Outlook (and the MCP server if it is running).
- Run `python tests/test_connection.py` and run a search. If AdvancedSearch works, you will no longer see "AdvancedSearch unavailable" in the logs.

---

## 🔧 Troubleshooting

| Issue | What to do |
|-------|------------|
| **"Outlook.Application" error** | Use desktop Outlook (not web); start Outlook before the server. |
| **Security dialog** | Normal on first access; click Allow. Optionally enable `use_extended_mapi_login`. |
| **Shared mailbox not accessible** | Check permissions, address in config, and that the mailbox is in your profile. |
| **Shared mailbox: "registry or installation problem"** | Usually means the shared mailbox address is invalid or a placeholder. In `config.properties`: set `shared_mailbox_email` to your **real** shared mailbox address (e.g. `team@yourcompany.com`), or leave it **empty** if you don't use a shared mailbox. The default `your-shared-mailbox@example.com` will always fail. |
| **"AdvancedSearch unavailable (using fallback search)" in logs** | Normal in some Outlook setups. The server automatically uses a fallback (subject/folder filter) and you still get results. For faster index-based search, enable Outlook indexing (File → Options → Search → Indexing Options). |
| **Cursor MCP panel shows [error] and "undefined" for "Processing request of type..."** | Those lines are INFO logs from the MCP SDK (SubscribeRequest, ListToolsRequest, etc.), not from this server. The Outlook MCP server sets the SDK logger to WARNING so they are suppressed. If you still see them, they are harmless; the server is working. |
| **No search results** | Confirm matching emails exist; try broader terms; check indexing. |
| **Slow search** | Narrow search terms; lower `max_search_results`; ensure Outlook isn't syncing heavily. |

**Debug logging:**

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

---

## 📁 Project Structure

```
Outlook-MCP-Server/
├── outlook_mcp.py              # Entrypoint
├── requirements.txt            # mcp, pywin32
├── requirements-ci.txt         # CI (no pywin32)
├── requirements-dev.txt        # flake8, mypy
├── pyproject.toml              # Lint/tool config
├── src/
│   ├── config/
│   │   ├── config_reader.py    # Load, validate, reload
│   │   └── config.properties   # User settings
│   ├── tools/
│   │   └── outlook_tools.py    # MCP tool definitions
│   └── utils/
│       ├── outlook_client.py   # Outlook COM + search + cache
│       └── email_formatter.py  # Format for AI
├── .github/workflows/
│   └── test.yml                # Lint + unit tests
└── tests/
    ├── conftest.py
    ├── test_config_reader.py
    ├── test_email_formatter.py
    └── test_connection.py     # Windows + Outlook
```

---

## 🔒 Security

- **Local only** — Server runs locally and talks to Outlook via COM.
- **No stored credentials** — Uses the current Windows user's Outlook profile.
- **Permission dialogs** — Windows may prompt for Outlook access; allow as needed.
- **Logging** — Does not log email addresses or sensitive config in normal operation.
- **Scope** — Limit to the mailboxes you configure (personal + listed shared).

---

## 🤝 Contributing

1. Fork the repo  
2. Create a feature branch  
3. Make changes and add tests if applicable  
4. Open a pull request  

---

## 📄 License

MIT — see LICENSE.

---

## 🙏 Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io)
- [pywin32](https://github.com/mhammond/pywin32) for the Outlook COM interface

---

## 👤 Author

**Hemprasad Badgujar**

*Outlook MCP Server* — Bring Outlook into the Model Context Protocol.
