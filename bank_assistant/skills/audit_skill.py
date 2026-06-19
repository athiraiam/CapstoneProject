"""
skills/audit_skill.py
----------------------
Reusable skill: logging every query, result, and session event to audit log files.
Used by AuditLoggerAgent and hooks.
"""

import os
import datetime

LOGS_DIR = os.path.join(os.path.dirname(__file__), "..", "logs")


def _log_file(name: str) -> str:
    os.makedirs(LOGS_DIR, exist_ok=True)
    return os.path.join(LOGS_DIR, name)


def log_query(question: str, matched_policy: str | None, found: bool, session_id: str = "unknown") -> None:
    """Append one query record to query_audit.log."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status = "MATCHED" if found else "FALLBACK"
    policy = matched_policy or "—"
    line = f"[{ts}] SESSION={session_id} | STATUS={status} | POLICY={policy} | Q={question}\n"
    with open(_log_file("query_audit.log"), "a", encoding="utf-8") as f:
        f.write(line)


def log_unresolved(question: str, session_id: str = "unknown") -> None:
    """Append to unresolved_queries.log — used by escalation hook."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] SESSION={session_id} | UNRESOLVED | Q={question}\n"
    with open(_log_file("unresolved_queries.log"), "a", encoding="utf-8") as f:
        f.write(line)


def log_session_event(event: str, session_id: str = "unknown", detail: str = "") -> None:
    """Log session lifecycle events (start, end, clear)."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] SESSION={session_id} | EVENT={event} | {detail}\n"
    with open(_log_file("session_events.log"), "a", encoding="utf-8") as f:
        f.write(line)


def detect_compliance_risk(question: str) -> bool:
    """
    Flag if the question may indicate a compliance issue.
    Returns True if any risk keyword is found.
    """
    risk_keywords = [
        "fraud", "money laundering", "suspicious", "bypass", "override",
        "without kyc", "skip verification", "fake document", "bribe",
        "corrupt", "illegal", "sanction", "blacklist", "terrorist",
    ]
    q_lower = question.lower()
    return any(kw in q_lower for kw in risk_keywords)


def read_audit_log(log_name: str = "query_audit.log", last_n: int = 20) -> list[str]:
    """Read the last N lines from an audit log. Used by the UI audit panel."""
    path = _log_file(log_name)
    if not os.path.exists(path):
        return ["No log entries yet."]
    with open(path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    return [l.rstrip() for l in lines[-last_n:]]
