"""Simple email formatting for AI-readable responses."""
# Author: Hemprasad Badgujar

import json
from collections import defaultdict
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

from ..config.config_reader import config


def _serialize_for_json(val: Any) -> Any:
    """Make value JSON-serializable (datetime -> isoformat, etc.)."""
    if hasattr(val, "isoformat") and callable(getattr(val, "isoformat")):
        return val.isoformat()
    if isinstance(val, dict):
        return {k: _serialize_for_json(v) for k, v in val.items()}
    if isinstance(val, (list, tuple)):
        return [_serialize_for_json(v) for v in val]
    return val


def format_mailbox_status(access_result: Dict[str, Any]) -> Dict[str, Any]:
    """
    Format mailbox access status for AI consumption.
    Args:
        access_result (Dict[str, Any]): Mailbox access result.
    Returns:
        Dict[str, Any]: Formatted mailbox status.
    """
    return {
        "status": "success",
        "connection": {
            "outlook_connected": access_result.get("outlook_connected", False),
            "timestamp": datetime.now().isoformat()
        },
        "personal_mailbox": {
            "accessible": access_result.get("personal_accessible", False),
            "name": access_result.get("personal_name", "Personal Mailbox"),
            "retention_months": access_result.get("retention_personal_months", 6)
        },
        "shared_mailbox": {
            "configured": access_result.get("shared_configured", False),
            "accessible": access_result.get("shared_accessible", False),
            "name": access_result.get("shared_name", "Shared Mailbox"),
            "names": access_result.get("shared_names", []),
            "email": config.get("shared_mailbox_email", "Not configured"),
            "retention_months": access_result.get("retention_shared_months", 12),
        },
        "errors": access_result.get("errors", []),
        "notes": {
            "security_dialog": (
                "You may need to grant permission when Outlook security "
                "dialog appears."
            ),
            "retention_info": (
                "Search scope is limited by retention policy settings."
            ),
        }
    }


def format_email_chain(
    emails: List[Dict[str, Any]], search_subject: str, include_body: bool = True
) -> Dict[str, Any]:
    """
    Format email chain results for AI analysis.
    Args:
        emails: List of email dicts.
        search_subject: Subject or phrase searched.
        include_body: If True include full body; if False only body_preview.
    Returns:
        Formatted email chain analysis.
    """
    if not emails:
        return {
            "status": "no_emails_found",
            "search_subject": search_subject,
            "message": f"No emails found for subject: '{search_subject}'"
        }
    conversations = group_by_conversation(emails)
    stats = {
        "total_emails": len(emails),
        "conversations": len(conversations),
        "date_range": get_date_range(emails),
        "mailbox_distribution": get_mailbox_distribution(emails),
        "participants": get_participants(emails)
    }

    def fmt(e: Dict[str, Any]) -> Dict[str, Any]:
        return format_single_email(e, include_body=include_body)

    # Format conversations chronologically
    formatted_conversations = []
    for conv_id, conv_emails in conversations.items():
        conv_emails.sort(key=lambda x: x.get('received_time', datetime.min))
        formatted_conv = {
            "conversation_id": conv_id,
            "email_count": len(conv_emails),
            "date_range": get_date_range(conv_emails),
            "participants": get_participants(conv_emails),
            "emails": [fmt(email) for email in conv_emails]
        }
        formatted_conversations.append(formatted_conv)

    def _max_time(c):
        return max(
            parse_iso_time(e["received_time"])
            for e in c["emails"]
            if e.get("received_time")
        )
    formatted_conversations.sort(key=_max_time, reverse=True)

    sorted_emails = sorted(
        emails,
        key=lambda x: _ensure_datetime(x.get("received_time")) or datetime.min,
        reverse=True,
    )
    return {
        "status": "success",
        "search_subject": search_subject,
        "summary": stats,
        "conversations": formatted_conversations,
        "all_emails_chronological": [fmt(email) for email in sorted_emails],
    }


def format_email_chain_to_json(formatted: Dict[str, Any]) -> str:
    """Return email chain as JSON for MCP tool (structured, parseable)."""
    return json.dumps(_serialize_for_json(formatted), indent=2, default=str)


def format_email_chain_pretty_text(
    formatted: Dict[str, Any], width: int = 72
) -> str:
    """Return email chain as readable plain text (e.g. list_latest_emails)."""
    lines: List[str] = []
    status = formatted.get("status", "")
    if status == "no_emails_found":
        return formatted.get("message", "No emails found.")
    summary = formatted.get("summary", {})
    total = summary.get("total_emails", 0)
    conv_count = summary.get("conversations", 0)
    date_range = summary.get("date_range", {})
    lines.append(f"  Total: {total} email(s)  |  Conversations: {conv_count}")
    if date_range:
        lines.append(f"  Date range: {date_range.get('first', '')} → {date_range.get('last', '')}")
    lines.append("")
    emails = formatted.get("all_emails_chronological", [])
    for i, email in enumerate(emails, 1):
        subject = email.get("subject", "No Subject")
        sender = email.get("sender_name", "Unknown")
        sender_email = email.get("sender_email", "")
        from_line = f"{sender} <{sender_email}>" if sender_email else sender
        received = email.get("received_time") or ""
        body = (email.get("body") or email.get("body_preview", "")) or ""
        body_preview = (body[:200] + "…") if len(body) > 200 else body
        lines.append("─" * width)
        lines.append(f"  [{i}]  {subject}")
        lines.append("─" * width)
        lines.append(f"  From:   {from_line}")
        lines.append(f"  Date:   {received}")
        if body_preview:
            lines.append(f"  Body:   {body_preview.strip()[:500]}")
        lines.append("")
    if emails:
        lines.append("─" * width)
    return "\n".join(lines)


# --- Reserved for future alert tool: format_alert_analysis and helpers below ---


def format_alert_analysis(alerts: List[Dict[str, Any]], search_pattern: str) -> Dict[str, Any]:
    """Format alert analysis results for AI consumption. Reserved for future alert tool."""
    
    if not alerts:
        return {
            "status": "no_alerts_found",
            "search_pattern": search_pattern,
            "message": f"No alerts found for pattern: '{search_pattern}'"
        }
    
    # Analyze alert patterns based on importance levels
    urgent_alerts = []
    normal_alerts = []
    
    analyze_importance = config.get_bool('analyze_importance_levels', True)
    
    for alert in alerts:
        # Check if alert is marked as high importance by sender
        is_urgent = alert.get('importance', 1) > 1
        
        # Additional urgency indicators if enabled
        if analyze_importance:
            subject = alert.get('subject', '').lower()
            # Simple urgency detection based on common urgent phrases
            urgent_phrases = ['urgent', 'critical', 'emergency', 'asap', 'immediate']
            is_urgent = is_urgent or any(phrase in subject for phrase in urgent_phrases)
        
        if is_urgent:
            urgent_alerts.append(alert)
        else:
            normal_alerts.append(alert)
    
    # Calculate alert frequency by day
    alert_frequency = calculate_daily_frequency(alerts)
    
    # Get recent alerts (last 10)
    recent_alerts = sorted(alerts, key=lambda x: x.get('received_time', datetime.min), reverse=True)[:10]
    
    # Summary statistics
    stats = {
        "total_alerts": len(alerts),
        "urgent_alerts": len(urgent_alerts),
        "normal_alerts": len(normal_alerts),
        "date_range": get_date_range(alerts),
        "mailbox_distribution": get_mailbox_distribution(alerts),
        "daily_frequency": alert_frequency,
        "response_indicators": analyze_responses(alerts)
    }
    
    return {
        "status": "success",
        "search_pattern": search_pattern,
        "summary": stats,
        "urgent_alerts": [format_single_email(alert) for alert in urgent_alerts[:5]],  # Top 5 urgent
        "recent_alerts": [format_single_email(alert) for alert in recent_alerts],
        "timeline": create_alert_timeline(alerts),
        "recommendations": generate_alert_recommendations(stats, urgent_alerts)
    }


def format_single_email(email: Dict[str, Any], include_body: bool = True) -> Dict[str, Any]:
    """Format a single email for AI consumption. Set include_body=False to omit full body (smaller response)."""
    raw_body = email.get("body", "") or ""
    formatted = {
        "subject": email.get("subject", "No Subject"),
        "sender_name": email.get("sender_name", "Unknown"),
        "sender_email": email.get("sender_email", ""),
        "recipients": email.get("recipients", []),
        "folder": email.get("folder_name", "Unknown"),
        "mailbox": email.get("mailbox_type", "unknown"),
        "body_preview": raw_body[:500] if raw_body else "",
        "attachments": email.get("attachments_count", 0),
        "importance": get_importance_text(email.get("importance", 1)),
        "unread": email.get("unread", False),
        "size_kb": round(email.get("size", 0) / 1024, 1),
        "entry_id": email.get("entry_id"),
    }
    if include_body:
        formatted["body"] = raw_body
    else:
        formatted["body"] = ""
    
    if config.get_bool("include_timestamps", True):
        received_time = email.get("received_time")
        if received_time is None:
            formatted["received_time"] = None
        elif isinstance(received_time, str):
            formatted["received_time"] = received_time
        elif hasattr(received_time, "isoformat") and callable(getattr(received_time, "isoformat")):
            formatted["received_time"] = received_time.isoformat()
        else:
            try:
                formatted["received_time"] = str(received_time)
            except Exception:
                formatted["received_time"] = None

    return formatted


def group_by_conversation(emails: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """Group emails by conversation based on subject similarity."""
    conversations = defaultdict(list)
    
    for email in emails:
        subject = email.get('subject', '').strip()
        
        # Clean subject for grouping (remove Re:, Fwd:, etc.)
        clean_subject = subject
        prefixes = ['re:', 'fwd:', 'fw:', 'reply:', 'forward:']
        for prefix in prefixes:
            if clean_subject.lower().startswith(prefix):
                clean_subject = clean_subject[len(prefix):].strip()
        
        # Use cleaned subject as conversation key
        conv_key = clean_subject.lower()
        conversations[conv_key].append(email)
    
    return dict(conversations)


def get_date_range(emails: List[Dict[str, Any]]) -> Dict[str, Optional[str]]:
    """Get date range of emails. Handles received_time as string or datetime."""
    if not emails:
        return {"first": None, "last": None}
    dates = [_ensure_datetime(e.get("received_time")) for e in emails]
    dates = [d for d in dates if d is not None]
    if not dates:
        return {"first": None, "last": None}
    return {
        "first": min(dates).isoformat(),
        "last": max(dates).isoformat(),
    }


def get_mailbox_distribution(emails: List[Dict[str, Any]]) -> Dict[str, int]:
    """Get distribution of emails across mailboxes."""
    distribution = {"personal": 0, "shared": 0, "unknown": 0}
    
    for email in emails:
        mailbox_type = email.get('mailbox_type', 'unknown')
        if mailbox_type in distribution:
            distribution[mailbox_type] += 1
        else:
            distribution['unknown'] += 1
    
    return distribution


def get_participants(emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Get list of email participants with counts."""
    participant_counts = defaultdict(int)
    participant_emails = {}
    
    for email in emails:
        sender = email.get('sender_name', 'Unknown')
        sender_email = email.get('sender_email', '')
        
        participant_counts[sender] += 1
        if sender_email:
            participant_emails[sender] = sender_email
        
        # Count recipients
        for recipient in email.get('recipients', []):
            participant_counts[recipient] += 1
    
    # Sort by participation count
    participants = []
    for name, count in sorted(participant_counts.items(), key=lambda x: x[1], reverse=True):
        participants.append({
            "name": name,
            "email": participant_emails.get(name, ''),
            "participation_count": count
        })
    
    return participants[:10]  # Top 10 participants


def calculate_daily_frequency(alerts: List[Dict[str, Any]]) -> float:
    """Calculate average alerts per day. Handles received_time as string or datetime."""
    if not alerts:
        return 0.0
    dates = []
    for alert in alerts:
        dt = _ensure_datetime(alert.get("received_time"))
        if dt is not None:
            dates.append(dt.date())
    if not dates:
        return 0.0
    date_range_days = (max(dates) - min(dates)).days + 1
    return round(len(alerts) / max(date_range_days, 1), 2)


def analyze_responses(alerts: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Analyze response patterns in alerts."""
    replies = sum(1 for alert in alerts if alert.get('subject', '').lower().startswith(('re:', 'reply:')))
    total = len(alerts)
    
    return {
        "replies_found": replies,
        "total_alerts": total,
        "response_rate_percent": round((replies / total) * 100, 1) if total > 0 else 0
    }


def create_alert_timeline(alerts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Create chronological timeline of alerts. Handles received_time as string or datetime."""
    timeline = []

    def _sort_key(a: Dict[str, Any]) -> datetime:
        dt = _ensure_datetime(a.get("received_time"))
        return dt if dt is not None else datetime.min

    sorted_alerts = sorted(alerts, key=_sort_key)

    for alert in sorted_alerts:
        dt = _ensure_datetime(alert.get("received_time"))
        timeline_entry = {
            "timestamp": dt.isoformat() if dt is not None else None,
            "subject": (alert.get("subject") or "No Subject")[:100],
            "sender": alert.get("sender_name", "Unknown"),
            "mailbox": alert.get("mailbox_type", "unknown"),
            "folder": alert.get("folder_name", "Unknown"),
            "importance": get_importance_text(alert.get("importance", 1)),
        }
        timeline.append(timeline_entry)

    return timeline


def generate_alert_recommendations(stats: Dict[str, Any], urgent_alerts: List[Dict[str, Any]]) -> List[str]:
    """Generate actionable recommendations based on alert analysis."""
    recommendations = []
    
    total_alerts = stats.get('total_alerts', 0)
    urgent_count = stats.get('urgent_alerts', 0)
    daily_frequency = stats.get('daily_frequency', 0)
    
    # High frequency alerts
    if daily_frequency > 5:
        recommendations.append(f"High alert frequency detected ({daily_frequency} alerts/day) - investigate root causes")
    
    # High urgency rate
    if urgent_count > 0 and total_alerts > 0:
        urgent_rate = (urgent_count / total_alerts) * 100
        if urgent_rate > 30:
            recommendations.append(f"High urgency rate ({urgent_rate:.1f}%) - review alert thresholds")
    
    # Response rate analysis
    response_rate = stats.get('response_indicators', {}).get('response_rate_percent', 0)
    if response_rate < 50:
        recommendations.append(f"Low response rate ({response_rate}%) - review alert response procedures")
    
    # Recent urgent alerts
    if urgent_count > 0:
        recommendations.append(f"Review {urgent_count} urgent alerts for immediate action")
    
    # Mailbox distribution
    mailbox_dist = stats.get('mailbox_distribution', {})
    if mailbox_dist.get('personal', 0) > 0 and mailbox_dist.get('shared', 0) == 0:
        recommendations.append("Alerts found only in personal mailbox - verify shared mailbox routing")
    
    if not recommendations:
        recommendations.append("No immediate issues detected - continue monitoring")
    
    return recommendations


def get_importance_text(importance: int) -> str:
    """Convert importance number to text."""
    importance_map = {0: "Low", 1: "Normal", 2: "High"}
    return importance_map.get(importance, "Normal")


def _to_naive_utc(dt: datetime) -> datetime:
    """Convert to naive UTC to avoid comparing offset-aware with offset-naive."""
    if getattr(dt, "tzinfo", None) is None:
        return dt
    return dt.astimezone(timezone.utc).replace(tzinfo=None)


def parse_iso_time(iso_string: Optional[Any]) -> datetime:
    """Parse ISO timestamp string or datetime to datetime. Returns datetime.min on invalid input."""
    if iso_string is None:
        return datetime.min
    if hasattr(iso_string, "isoformat") and hasattr(iso_string, "date"):
        return _to_naive_utc(iso_string)  # type: ignore[arg-type]
    if not isinstance(iso_string, str):
        return datetime.min
    try:
        dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
        return _to_naive_utc(dt) if getattr(dt, "tzinfo", None) else dt
    except (ValueError, AttributeError):
        return datetime.min


def _ensure_datetime(val: Any) -> Optional[datetime]:
    """Normalize value to datetime for date range/sort. Returns None if invalid."""
    if val is None:
        return None
    if hasattr(val, "isoformat") and hasattr(val, "date"):
        if isinstance(val, datetime):
            return _to_naive_utc(val)
        return val  # type: ignore[return-value]
    if isinstance(val, str):
        dt = parse_iso_time(val)
        return dt if dt != datetime.min else None
    return None
