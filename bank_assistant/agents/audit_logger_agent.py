"""
agents/audit_logger_agent.py
-----------------------------
SubAgent 7 — AuditLoggerAgent
Background agent that logs every query, result, and session event.
Runs after every answer — invisible to the employee.
"""

from skills.audit_skill import log_query, log_unresolved, log_session_event, detect_compliance_risk


class AuditLoggerAgent:
    """
    Background subagent for audit logging and compliance risk detection.
    """

    name = "AuditLoggerAgent"

    def __init__(self, session_id: str = "unknown"):
        self.session_id = session_id

    def log_answer(self, question: str, matched_policy: str | None, found: bool) -> dict:
        """
        Log a query+answer event.
        If the answer was not found, also log to unresolved log.
        If a compliance risk is detected, return a warning flag.

        Returns:
            {
              "logged"           : bool
              "compliance_risk"  : bool  — True if risk keywords detected
              "session_id"       : str
            }
        """
        # Always log to main audit log
        log_query(question, matched_policy, found, self.session_id)

        # Log unresolved separately
        if not found:
            log_unresolved(question, self.session_id)

        # Detect compliance risk
        risk = detect_compliance_risk(question)

        return {
            "logged":          True,
            "compliance_risk": risk,
            "session_id":      self.session_id,
        }

    def log_event(self, event: str, detail: str = "") -> None:
        """Log a session lifecycle event (start, end, clear)."""
        log_session_event(event, self.session_id, detail)
