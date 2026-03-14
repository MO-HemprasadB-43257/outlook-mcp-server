"""High-performance Outlook client for mailbox access and email search."""
# Author: Hemprasad Badgujar

import logging
import random
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Set

import pythoncom
import win32com.client

from ..config.config_reader import config

logger = logging.getLogger(__name__)


def _shared_mailbox_emails() -> List[str]:
    """Return list of shared mailbox emails from config (multiple or single)."""
    emails = config.get_list("shared_mailbox_emails")
    if emails:
        return [e.strip() for e in emails if e and str(e).strip()]
    single = config.get("shared_mailbox_email")
    if single and str(single).strip():
        return [str(single).strip()]
    return []


def _to_naive_utc(dt: datetime) -> datetime:
    """Convert to naive UTC so we never compare offset-aware with offset-naive."""
    if dt.tzinfo is None:
        return dt
    return dt.astimezone(timezone.utc).replace(tzinfo=None)


def _received_time_for_sort(email: Dict[str, Any]) -> datetime:
    """Normalize received_time to datetime for reliable sorting (handles COM dates)."""
    val = email.get("received_time")
    if val is None:
        return datetime.min
    if isinstance(val, datetime):
        return _to_naive_utc(val)
    try:
        if hasattr(val, "timestamp"):
            return datetime.fromtimestamp(val.timestamp())
        if hasattr(val, "isoformat"):
            dt = datetime.fromisoformat(val.isoformat())
            return _to_naive_utc(dt) if getattr(dt, "tzinfo", None) else dt
    except (ValueError, OSError, TypeError):
        pass
    try:
        dt = datetime.fromisoformat(str(val).replace("Z", "+00:00"))
        return _to_naive_utc(dt) if getattr(dt, "tzinfo", None) else dt
    except (ValueError, TypeError):
        return datetime.min


class OutlookClient:
    """High-performance client for accessing Outlook mailboxes with optimized search.

    Connection and shared-recipient cache are used only on the main thread.
    Worker threads (parallel search) create their own COM objects and do not
    use _shared_recipient_cache, for thread-safety.
    """

    def __init__(self) -> None:
        # === Connection State (main thread only) ===
        self.outlook: Any = None
        self.namespace: Any = None
        self.connected: bool = False
        self._max_retries: int = config.get_int("max_retry_attempts", 3)

        # === Caching (main thread only for shared_recipient; search_cache used after workers join) ===
        self._search_cache: Dict[str, Dict[str, Any]] = {}
        self._folder_cache: Dict[str, Any] = {}
        self._shared_recipient_cache: Any = None  # Main-thread only; workers create own recipient
    
    # === Connection Methods ===
    def connect(self, retry_attempt: int = 0) -> bool:
        """
        Connect to Outlook application with retry logic.
        """
        try:
            logger.info("Connecting to Outlook...")
            start_time = time.time()
            pythoncom.CoInitialize()
            try:
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                logger.info("Connected to existing Outlook instance")
            except Exception:
                logger.info("No existing Outlook instance, launching new one...")
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            if config.get_bool('use_extended_mapi_login', True):
                try:
                    logger.info("Attempting Extended MAPI login to reduce security prompts...")
                    self.namespace.Logon(None, None, False, True)
                    logger.info("Extended MAPI login successful")
                except Exception as logon_error:
                    logger.warning("Extended MAPI login failed: %s", logon_error)
            connection_time = time.time() - start_time
            logger.info("Successfully connected to Outlook in %.2f seconds", connection_time)
            self.connected = True
            return True
        except Exception as e:
            logger.error("Failed to connect to Outlook (attempt %s): %s", retry_attempt + 1, e)
            self.connected = False
            if retry_attempt < self._max_retries - 1:
                base_wait = (2 ** retry_attempt) * 1
                jitter = random.uniform(0, 0.5 * base_wait)
                wait_time = base_wait + jitter
                logger.info("Retrying connection in %.1f seconds...", wait_time)
                time.sleep(wait_time)
                return self.connect(retry_attempt + 1)
            return False
    
    # === Mailbox Access Methods ===
    def check_access(self) -> Dict[str, Any]:
        """
        Check access to personal and shared mailboxes.
        """
        if not self.connected:
            if not self.connect():
                return {"error": "Could not connect to Outlook"}
        shared_emails = _shared_mailbox_emails()
        result = {
            "outlook_connected": True,
            "personal_accessible": False,
            "shared_accessible": False,
            "shared_configured": bool(shared_emails),
            "shared_names": [],
            "retention_personal_months": config.get_int("personal_retention_months", 6),
            "retention_shared_months": config.get_int("shared_retention_months", 12),
            "errors": [],
        }
        try:
            personal_inbox = self.namespace.GetDefaultFolder(6)
            if personal_inbox:
                result["personal_accessible"] = True
                result["personal_name"] = self._get_store_display_name(personal_inbox)
        except Exception as e:
            result["errors"].append(f"Personal mailbox error: {str(e)}")
        if shared_emails:
            for shared_email in shared_emails:
                try:
                    recipient = self.namespace.CreateRecipient(shared_email)
                    recipient.Resolve()
                    if recipient.Resolved:
                        shared_inbox = self.namespace.GetSharedDefaultFolder(recipient, 6)
                        if shared_inbox:
                            result["shared_accessible"] = True
                            result["shared_names"].append(
                                self._get_store_display_name(shared_inbox)
                            )
                except Exception as e:
                    result["errors"].append(f"Shared mailbox {shared_email}: {e}")
            if result.get("shared_names"):
                result["shared_name"] = result["shared_names"][0]
            if result["shared_accessible"] and self._shared_recipient_cache is None:
                try:
                    first = shared_emails[0]
                    self._shared_recipient_cache = self.namespace.CreateRecipient(first)
                    self._shared_recipient_cache.Resolve()
                except Exception:
                    pass
        return result
    
    # === Search Methods ===
    def search_emails(
        self,
        search_text: str,
        include_personal: bool = True,
        include_shared: bool = True,
    ) -> List[Dict[str, Any]]:
        """Search emails in subject and body using exact phrase, parallel."""
        if not self.connected:
            if not self.connect():
                return []
        profile = config.get_bool("profile_search", False)
        search_start = time.time() if profile else None
        max_results = config.get_int("max_search_results", 500)
        cache_ttl = config.get_int("search_cache_ttl_seconds", 3600)
        cache_key = f"{search_text}_{include_personal}_{include_shared}_{max_results}"
        if cache_key in self._search_cache:
            cache_entry = self._search_cache[cache_key]
            if time.time() - cache_entry["timestamp"] < cache_ttl:
                logger.info("Returning cached results for '%s'", search_text)
                return cache_entry["data"]
        all_emails = []
        shared_emails = _shared_mailbox_emails()
        num_tasks = (1 if include_personal else 0) + (len(shared_emails) if include_shared else 0)
        max_workers = max(
            1,
            min(8, config.get_int("parallel_search_workers", 2), max(num_tasks, 1)),
        )
        if num_tasks >= 1 and (include_personal or (include_shared and shared_emails)):
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = []
                if include_personal:
                    futures.append(
                        executor.submit(
                            self._search_mailbox_wrapper,
                            "personal",
                            search_text,
                            max_results,
                            None,
                        )
                    )
                if include_shared:
                    for shared_email in shared_emails:
                        futures.append(
                            executor.submit(
                                self._search_mailbox_wrapper,
                                "shared",
                                search_text,
                                max_results,
                                shared_email,
                            )
                        )
                for future in as_completed(futures):
                    try:
                        emails = future.result()
                        all_emails.extend(emails)
                    except Exception as e:
                        logger.error("Error in parallel search: %s", e)
        else:
            # Sequential search for single mailbox
            if include_personal:
                personal_emails = self._search_mailbox_comprehensive(
                    self.namespace.GetDefaultFolder(6),
                    search_text,
                    'personal',
                    max_results,
                )
                all_emails.extend(personal_emails)
                logger.info("Found %s emails in personal mailbox", len(personal_emails))
            if include_shared and shared_emails:
                for shared_email in shared_emails:
                    try:
                        if (
                            self._shared_recipient_cache is not None
                            and self._shared_recipient_cache.Resolved
                        ):
                            rec = self._shared_recipient_cache
                        else:
                            rec = self.namespace.CreateRecipient(shared_email)
                            rec.Resolve()
                        if rec.Resolved:
                            shared_inbox = self.namespace.GetSharedDefaultFolder(rec, 6)
                            found = self._search_mailbox_comprehensive(
                                shared_inbox,
                                search_text,
                                "shared",
                                max_results - len(all_emails),
                            )
                            all_emails.extend(found)
                            logger.info(
                                "Found %s emails in shared mailbox %s",
                                len(found), shared_email,
                            )
                    except Exception as e:
                        logger.error("Error searching shared mailbox %s: %s", shared_email, e)
        all_emails.sort(key=_received_time_for_sort, reverse=True)
        limited_results = all_emails[:max_results]
        self._search_cache[cache_key] = {
            "data": limited_results,
            "timestamp": time.time(),
        }
        cache_max = config.get_int("search_cache_max_entries", 100)
        while len(self._search_cache) > max(1, cache_max):
            oldest_key = min(
                self._search_cache.keys(),
                key=lambda k: self._search_cache[k].get("timestamp", 0),
            )
            del self._search_cache[oldest_key]
        if profile and search_start is not None:
            elapsed = time.time() - search_start
            logger.info(
                "Search '%s' completed in %.2fs, %s results",
                search_text,
                elapsed,
                len(limited_results),
            )
        return limited_results

    def search_emails_by_subject(
        self,
        subject: str,
        include_personal: bool = True,
        include_shared: bool = True,
    ) -> List[Dict[str, Any]]:
        """Legacy: redirects to search_emails for backward compatibility.

        Args:
            subject (str): Subject to search for.
            include_personal (bool): Search personal mailbox.
            include_shared (bool): Search shared mailbox.
        Returns:
            List[Dict[str, Any]]: List of matching emails.
        """
        return self.search_emails(subject, include_personal, include_shared)

    def get_latest_emails(
        self,
        count: int,
        include_personal: bool = True,
        include_shared: bool = True,
    ) -> List[Dict[str, Any]]:
        """
        Get the N most recent emails from Inbox(es) without a search phrase.
        Use this for "last email", "latest 10 emails", etc.
        """
        if not self.connected:
            if not self.connect():
                return []
        max_cap = config.get_int("max_search_results", 500)
        n = max(1, min(count, max_cap))
        all_emails: List[Dict[str, Any]] = []
        if include_personal:
            try:
                inbox = self.namespace.GetDefaultFolder(6)
                all_emails.extend(self._get_latest_from_inbox(inbox, "personal", n))
            except Exception as e:
                logger.error("Error getting latest from personal inbox: %s", e)
        if include_shared:
            shared_emails = _shared_mailbox_emails()
            for shared_email in shared_emails:
                try:
                    rec = self.namespace.CreateRecipient(shared_email)
                    rec.Resolve()
                    if rec.Resolved:
                        shared_inbox = self.namespace.GetSharedDefaultFolder(rec, 6)
                        all_emails.extend(
                            self._get_latest_from_inbox(shared_inbox, "shared", n)
                        )
                except Exception as e:
                    logger.error("Error getting latest from shared %s: %s", shared_email, e)
        # Sort by normalized received_time (handles COM dates) so newest is first
        all_emails.sort(key=_received_time_for_sort, reverse=True)
        return all_emails[:n]

    def _get_latest_from_inbox(
        self,
        inbox_folder: Any,
        mailbox_type: str,
        max_results: int,
    ) -> List[Dict[str, Any]]:
        """Return up to max_results most recent emails from an Inbox folder.
        Fetches from both start and end of collection, then sorts in Python by
        received_time so we reliably return newest-first (COM Sort order can vary).
        """
        emails: List[Dict[str, Any]] = []
        seen_ids: Set[str] = set()
        try:
            items = inbox_folder.Items
            items.Sort("[ReceivedTime]", True)
            count = getattr(items, "Count", 0) or 0
            if count == 0:
                return []
            # Fetch from both ends: COM Sort order can vary, so cover start and end
            window = min(200, count)
            start_indices = list(range(1, window + 1))
            end_indices = list(range(max(1, count - window + 1), count + 1))
            indices_to_fetch = sorted(set(start_indices) | set(end_indices))
            for i in indices_to_fetch:
                if i < 1 or i > count:
                    continue
                try:
                    item = items.Item(i)
                    entry_id = getattr(item, "EntryID", "") or ""
                    if entry_id and entry_id in seen_ids:
                        continue
                    email_data = self._extract_email_data(
                        item, inbox_folder.Name, mailbox_type
                    )
                    if email_data:
                        if entry_id:
                            seen_ids.add(entry_id)
                        emails.append(email_data)
                except Exception as e:
                    logger.debug("Error getting item %s from inbox: %s", i, e)
            emails.sort(key=_received_time_for_sort, reverse=True)
            return emails[:max_results]
        except Exception as e:
            logger.error("Error listing inbox %s: %s", inbox_folder.Name, e)
        return emails

    def _search_mailbox_wrapper(
        self,
        mailbox_type: str,
        search_text: str,
        max_results: int,
        shared_email: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        Wrapper for parallel mailbox search. Creates thread-local Outlook/namespace
        so COM is not used across threads (thread-safe).
        """
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            if mailbox_type == "personal":
                inbox = namespace.GetDefaultFolder(6)
                return self._search_mailbox_comprehensive(
                    inbox, search_text, "personal", max_results,
                    outlook=outlook, namespace=namespace,
                )
            if mailbox_type == "shared":
                email = shared_email or (config.get("shared_mailbox_email") or "").strip()
                if not email:
                    return []
                recipient = namespace.CreateRecipient(email)
                recipient.Resolve()
                if recipient.Resolved:
                    shared_inbox = namespace.GetSharedDefaultFolder(recipient, 6)
                    return self._search_mailbox_comprehensive(
                        shared_inbox, search_text, "shared", max_results,
                        outlook=outlook, namespace=namespace,
                    )
            return []
        except Exception as e:
            logger.error("Error in mailbox wrapper for %s: %s", mailbox_type, e)
            return []
        finally:
            pythoncom.CoUninitialize()

    def _search_mailbox_comprehensive(
        self,
        inbox_folder: Any,
        search_text: str,
        mailbox_type: str,
        max_results: int,
        outlook: Any = None,
        namespace: Any = None,
    ) -> List[Dict[str, Any]]:
        """
        Optimized search using AdvancedSearch. Use outlook/namespace when called
        from a worker thread (thread-local COM).
        """
        emails = []
        found_ids: Set[str] = set()
        use_outlook = outlook if outlook is not None else self.outlook
        # Coerce to string and escape for query safety
        search_text = str(search_text) if not isinstance(search_text, str) else search_text
        # DASL phrase: escape double-quote only
        search_text_dasl = search_text.replace('"', '""')

        try:
            scope = f"'{inbox_folder.FolderPath}'"
            query = (
                f'urn:schemas:httpmail:subject ci_phrasematch "{search_text_dasl}" OR '
                f'urn:schemas:httpmail:textdescription ci_phrasematch "{search_text_dasl}"'
            )
            logger.info("Performing AdvancedSearch in %s for '%s'", scope, search_text)

            search = use_outlook.AdvancedSearch(
                Scope=scope,
                Filter=query,
                SearchSubFolders=False,  # Don't search subfolders for inbox
                Tag="EmailBodySearch"  # Unique tag for this search
            )

            # Poll for completion with reduced timeout for responsiveness
            start_time = time.time()
            while not search.SearchComplete:
                time.sleep(0.05)
                if time.time() - start_time > 15:  # Reduced timeout to 15 seconds
                    logger.warning("AdvancedSearch timed out after 15 seconds")
                    break

            if search.SearchComplete:
                results = search.Results
                result_count = min(results.Count, max_results)
                batch_size = max(1, config.get_int("batch_processing_size", 50))
                logger.info(
                    "AdvancedSearch completed: found %s matches (taking %s)",
                    results.Count,
                    result_count,
                )
                for batch_start in range(1, result_count + 1, batch_size):
                    batch_end = min(batch_start + batch_size, result_count + 1)
                    for i in range(batch_start, batch_end):
                        try:
                            item = results.Item(i)
                            entry_id = getattr(item, "EntryID", "")
                            if entry_id and entry_id not in found_ids:
                                email_data = self._extract_email_data(
                                    item, inbox_folder.Name, mailbox_type
                                )
                                if email_data:
                                    emails.append(email_data)
                                    found_ids.add(entry_id)
                        except Exception as e:
                            logger.error("Error processing result %s: %s", i, e)
                    if len(emails) >= max_results:
                        break
            else:
                logger.warning("AdvancedSearch did not complete successfully")

        except Exception as e:
            logger.info(
                "AdvancedSearch unavailable (using fallback search): %s",
                (str(e)[:80] + "...") if len(str(e)) > 80 else str(e),
            )
            logger.debug("AdvancedSearch full error: %s", e)

            search_text_escaped = search_text.replace("'", "''").replace('"', '""')
            # Escape LIKE wildcards so user input cannot broaden the match
            search_text_escaped = search_text_escaped.replace("%", "[%]").replace("_", "[_]")
            items = inbox_folder.Items
            items.Sort("[ReceivedTime]", True)

            try:
                subject_filter = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_text_escaped}%'"
                filtered_items = items.Restrict(subject_filter)
                
                for item in filtered_items:
                    if len(emails) >= max_results:
                        break
                    
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in found_ids:
                        email_data = self._extract_email_data(item, inbox_folder.Name, mailbox_type)
                        if email_data:
                            emails.append(email_data)
                            found_ids.add(entry_id)
            except Exception as fallback_error:
                logger.debug("Fallback subject filter failed: %s", fallback_error)
        
        if len(emails) < max_results and config.get_bool("search_all_folders", False):
            try:
                additional_emails = self._search_other_folders(
                    inbox_folder.Parent,
                    search_text,
                    mailbox_type,
                    max_results - len(emails),
                    found_ids,
                    outlook=use_outlook,
                )
                emails.extend(additional_emails)
                logger.info("Additional emails found in other folders: %s", len(additional_emails))
            except Exception as e:
                logger.error("Error searching other folders: %s", e)
        
        return emails

    def _search_other_folders(
        self,
        store: Any,
        search_text: str,
        mailbox_type: str,
        max_results: int,
        found_ids: Set[str],
        outlook: Any = None,
    ) -> List[Dict[str, Any]]:
        """Search other folders using AdvancedSearch. Pass outlook when called from worker thread."""
        emails = []
        key_folders = ["Sent Items", "Drafts"]
        use_outlook = outlook if outlook is not None else self.outlook

        for folder_name in key_folders:
            if len(emails) >= max_results:
                break
            try:
                folder = self._get_folder_by_name(store, folder_name)
                if folder:
                    scope = f"'{folder.FolderPath}'"
                    search_text_escaped = search_text.replace('"', '""')
                    query = (
                        f'urn:schemas:httpmail:subject ci_phrasematch "{search_text_escaped}" OR '
                        f'urn:schemas:httpmail:textdescription ci_phrasematch "{search_text_escaped}"'
                    )
                    logger.info("AdvancedSearch in %s for '%s'", folder_name, search_text)
                    search = use_outlook.AdvancedSearch(
                        Scope=scope, 
                        Filter=query, 
                        SearchSubFolders=False, 
                        Tag=f"OtherFolderSearch_{folder_name}"
                    )
                    
                    # Poll with shorter timeout for secondary folders
                    start_time = time.time()
                    while not search.SearchComplete:
                        time.sleep(0.1)
                        if time.time() - start_time > 10:  # Shorter timeout for secondary folders
                            break
                    
                    if search.SearchComplete:
                        results = search.Results
                        result_count = min(results.Count, max_results - len(emails))
                        
                        for i in range(1, result_count + 1):
                            item = results.Item(i)
                            entry_id = getattr(item, 'EntryID', '')
                            if entry_id and entry_id not in found_ids:
                                email_data = self._extract_email_data(item, folder_name, mailbox_type)
                                if email_data:
                                    emails.append(email_data)
                                    found_ids.add(entry_id)
            except Exception as e:
                logger.debug("Error searching %s: %s", folder_name, e)
        
        return emails
    
    def _extract_email_data(
        self, item: Any, folder_name: str, mailbox_type: str
    ) -> Optional[Dict[str, Any]]:
        """Extract email data with optimized body and recipient handling."""
        try:
            # Get the full email body
            body = getattr(item, 'Body', '')
            # Limit body size: max_search_body_chars during search phase, else max_body_chars (0 = no limit)
            max_search_body = config.get_int('max_search_body_chars', 0)
            max_final_body = config.get_int('max_body_chars', 0)
            max_body_chars = max_search_body or max_final_body
            if max_body_chars > 0 and len(body) > max_body_chars:
                body = body[:max_body_chars] + " [truncated]"
            
            # Clean HTML if configured
            if config.get_bool('clean_html_content', True) and body:
                body = self._clean_html(body)
            
            # Get recipients list with limit for performance
            recipients = []
            max_recipients = min(config.get_int('max_recipients_display', 10), 20)  # Limit to 20 for memory safety
            try:
                recipient_count = 0
                for recipient in item.Recipients:
                    if recipient_count >= max_recipients:
                        recipients.append(f"... and {item.Recipients.Count - recipient_count} more")
                        break
                    recipients.append(getattr(recipient, 'Name', getattr(recipient, 'Address', '')))
                    recipient_count += 1
            except Exception as recv_err:
                logger.debug("Error reading recipients: %s", recv_err)

            rt = getattr(item, "ReceivedTime", None)
            if rt is None:
                rt = datetime.now()
            elif isinstance(rt, datetime):
                rt = _to_naive_utc(rt)

            email_data = {
                'subject': getattr(item, 'Subject', 'No Subject'),
                'sender_name': getattr(item, 'SenderName', 'Unknown'),
                'sender_email': getattr(item, 'SenderEmailAddress', ''),
                'recipients': recipients,
                'received_time': rt,
                'folder_name': folder_name,
                'mailbox_type': mailbox_type,
                'importance': getattr(item, 'Importance', 1),
                'body': body,  # Full body for summarization
                'size': getattr(item, 'Size', 0),
                'attachments_count': getattr(item.Attachments, 'Count', 0) if hasattr(item, 'Attachments') else 0,
                'unread': getattr(item, 'Unread', False),
                'entry_id': getattr(item, 'EntryID', '')
            }
            
            # Release COM reference to free memory
            item = None
            
            return email_data
        except Exception as e:
            logger.error("Error extracting email data: %s", e)
            return None

    def send_email(
        self,
        to: str,
        subject: str,
        body: str,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Compose and send a new email from the default account."""
        if not self.connected:
            if not self.connect():
                return {"status": "error", "message": "Could not connect to Outlook"}
        try:
            mail = self.outlook.CreateItem(0)  # olMailItem
            mail.To = to.strip()
            mail.Subject = subject
            mail.Body = body
            if cc:
                mail.CC = cc.strip()
            if bcc:
                mail.BCC = bcc.strip()
            mail.Send()
            return {"status": "sent", "to": to, "subject": subject}
        except Exception as e:
            logger.exception("Send email failed: %s", e)
            return {"status": "error", "message": "Failed to send email", "detail": str(e)[:200]}

    def reply_to_email(
        self,
        entry_id: str,
        body: Optional[str] = None,
        reply_all: bool = False,
    ) -> Dict[str, Any]:
        """Reply to an email by entry_id (from get_email_chain). Optionally add body; reply_all to reply to all."""
        if not entry_id or not str(entry_id).strip():
            return {"status": "error", "message": "entry_id is required"}
        if not self.connected:
            if not self.connect():
                return {"status": "error", "message": "Could not connect to Outlook"}
        try:
            item = self.namespace.GetItemFromID(entry_id.strip())
            reply = item.ReplyAll() if reply_all else item.Reply()
            if body:
                reply.Body = body + "\r\n" + (reply.Body or "")
            reply.Send()
            return {"status": "sent", "action": "reply_all" if reply_all else "reply", "entry_id": entry_id[:50]}
        except Exception as e:
            logger.exception("Reply to email failed: %s", e)
            return {"status": "error", "message": "Failed to reply", "detail": str(e)[:200]}

    def forward_email(
        self,
        entry_id: str,
        to: str,
        body: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Forward an email by entry_id. Requires to; optional body (e.g. FYI) prepended."""
        if not entry_id or not str(entry_id).strip():
            return {"status": "error", "message": "entry_id is required"}
        if not to or not str(to).strip():
            return {"status": "error", "message": "to is required"}
        if not self.connected:
            if not self.connect():
                return {"status": "error", "message": "Could not connect to Outlook"}
        try:
            item = self.namespace.GetItemFromID(entry_id.strip())
            fwd = item.Forward()
            fwd.To = to.strip()
            if body:
                fwd.Body = body + "\r\n" + (fwd.Body or "")
            fwd.Send()
            return {"status": "sent", "action": "forward", "to": to, "entry_id": entry_id[:50]}
        except Exception as e:
            logger.exception("Forward email failed: %s", e)
            return {"status": "error", "message": "Failed to forward", "detail": str(e)[:200]}

    def _get_store_display_name(self, folder: Any) -> str:
        """Safely get store display name from a folder."""
        try:
            if hasattr(folder, 'Parent'):
                parent = folder.Parent
                if hasattr(parent, 'DisplayName'):
                    return parent.DisplayName
                elif hasattr(parent, 'Name'):
                    return parent.Name
            return "Mailbox"
        except Exception as e:
            logger.debug("Could not get store display name: %s", e)
            return "Mailbox"

    def _get_folder_by_name(self, store: Any, name: str) -> Optional[Any]:
        """Get folder by name from cache or store."""
        cache_key = f"{id(store)}_{name}"
        
        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]
        
        try:
            for folder in store.GetRootFolder().Folders:
                if folder.Name.lower() == name.lower():
                    self._folder_cache[cache_key] = folder
                    return folder
        except Exception as e:
            logger.debug("Could not get folder by name %s: %s", name, e)

        return None
    
    def _clean_html(self, text: str) -> str:
        """Clean HTML from email body."""
        # Remove HTML tags
        text = re.sub(r'<[^>]+>', '', text)
        
        # Decode common HTML entities
        html_entities = {
            '&amp;': '&',
            '&lt;': '<',
            '&gt;': '>',
            '&quot;': '"',
            '&#39;': "'",
            '&nbsp;': ' '
        }
        
        for entity, char in html_entities.items():
            text = text.replace(entity, char)
        
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text


# Global client instance
outlook_client = OutlookClient()
