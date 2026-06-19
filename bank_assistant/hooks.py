"""
hooks.py
---------
All lifecycle hooks for the Bank Employee Assistant.
Hooks are called at specific points in the agent pipeline by the Orchestrator.

Available hooks:
  - pre_query_hook       : runs before any question is processed
  - post_answer_hook     : runs after every answer is returned
  - session_start_hook   : runs once when the app session starts
  - session_end_hook     : runs when the employee clears chat
  - escalation_hook      : runs when a query goes to FallbackAgent / EscalationAgent
  - compliance_alert_hook: runs when AuditLoggerAgent detects a risk keyword
"""

import datetime
from skills.audit_skill import log_session_event, log_unresolved


def pre_query_hook(question: str, session_id: str) -> dict:
    """
    Runs BEFORE the question reaches any agent.

    Responsibilities:
    - Validate the question is not empty
    - Log that a query was received
    - Return a sanitized question

    Returns: { "valid": bool, "question": str, "reason": str }
    """
    question = question.strip()

    if not question:
        return {"valid": False, "question": "", "reason": "Empty question"}

    if len(question) < 3:
        return {"valid": False, "question": question, "reason": "Question too short"}

    # Log incoming query
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_session_event("PRE_QUERY", session_id, f"Q={question[:80]}")

    return {"valid": True, "question": question, "reason": "OK"}


def post_answer_hook(question: str, result: dict, session_id: str, response_ms: float) -> None:
    """
    Runs AFTER the agent pipeline returns an answer.

    Responsibilities:
    - Log the matched policy and response time
    - Track answer found vs fallback rate
    """
    policy = result.get("policy_name") or result.get("policy_names") or "None"
    found  = result.get("found", False)
    status = "ANSWERED" if found else "FALLBACK"

    log_session_event(
        "POST_ANSWER", session_id,
        f"STATUS={status} | POLICY={policy} | TIME={response_ms:.0f}ms | Q={question[:60]}"
    )


def session_start_hook(session_id: str) -> None:
    """
    Runs ONCE when a new browser session starts.

    Responsibilities:
    - Log session start with timestamp
    - Record session ID for traceability
    """
    log_session_event("SESSION_START", session_id, "App loaded — knowledge base ready")


def session_end_hook(session_id: str, total_queries: int) -> None:
    """
    Runs when the employee clicks 'Clear Chat History'.

    Responsibilities:
    - Log session summary (total queries handled)
    - Mark session as closed
    """
    log_session_event(
        "SESSION_END", session_id,
        f"Total queries this session: {total_queries}"
    )


def escalation_hook(question: str, session_id: str) -> None:
    """
    Runs whenever a query is escalated (FallbackAgent / EscalationAgent).

    Responsibilities:
    - Log to the unresolved_queries.log for compliance review
    """
    log_unresolved(question, session_id)
    log_session_event("ESCALATION", session_id, f"Escalated: {question[:80]}")


def compliance_alert_hook(question: str, session_id: str) -> dict:
    """
    Runs when AuditLoggerAgent detects compliance risk keywords.

    Responsibilities:
    - Return a warning message to display in the UI
    - Log the flagged query

    Returns: { "alert": str }  — warning text to show in the chat UI
    """
    log_session_event("COMPLIANCE_ALERT", session_id, f"RISK FLAGGED: {question[:80]}")
    return {
        "alert": (
            "⚠️ **Compliance Alert:** Your query contains sensitive keywords. "
            "This question has been flagged and logged for review by the Compliance team. "
            "Please contact **compliance@bank.internal** immediately if this is urgent."
        )
    }
