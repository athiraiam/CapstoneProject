"""
agents/orchestrator.py
-----------------------
OrchestratorAgent — the top-level agent that coordinates all subagents.

Pipeline for every employee question:
  1. pre_query_hook          → validate & sanitize input
  2. QueryClassifierAgent    → detect intent
  3. Route to correct agent:
       policy_query  → PolicySearchAgent
                        if no match → FallbackAgent (self-healing)
                        if compliance risk → compliance_alert_hook
       summary       → SummaryAgent
       escalation    → EscalationAgent
       out_of_scope  → return fallback
  4. AuditLoggerAgent        → log query + result (background)
  5. post_answer_hook        → log response time
  6. escalation_hook         → if escalated, log to unresolved log
  7. Return final result to app.py
"""

import time
import uuid

from agents.query_classifier_agent import (
    QueryClassifierAgent,
    INTENT_POLICY_QUERY, INTENT_SUMMARY, INTENT_ESCALATION, INTENT_OUT_OF_SCOPE,
)
from agents.policy_search_agent  import PolicySearchAgent
from agents.multi_policy_agent   import MultiPolicyAgent
from agents.summary_agent        import SummaryAgent
from agents.escalation_agent     import EscalationAgent
from agents.fallback_agent       import FallbackAgent
from agents.audit_logger_agent   import AuditLoggerAgent
from skills.search_skill         import cross_reference
from hooks import (
    pre_query_hook, post_answer_hook,
    session_start_hook, session_end_hook,
    escalation_hook, compliance_alert_hook,
)


class OrchestratorAgent:
    """
    Top-level agent that owns the full question → answer pipeline.
    Instantiated once per Streamlit session.
    """

    name = "OrchestratorAgent"

    def __init__(self, policies: dict, session_id: str | None = None):
        self.policies   = policies
        self.session_id = session_id or str(uuid.uuid4())[:8]

        # Instantiate all subagents
        self.classifier  = QueryClassifierAgent()
        self.searcher    = PolicySearchAgent(policies)
        self.multi       = MultiPolicyAgent(policies)
        self.summarizer  = SummaryAgent(policies)
        self.escalator   = EscalationAgent(policies)
        self.fallback    = FallbackAgent(policies)
        self.auditor     = AuditLoggerAgent(self.session_id)

        # Fire session start hook
        session_start_hook(self.session_id)

    # ── Main entry point ──────────────────────────────────────────────────────
    def handle(self, question: str) -> dict:
        """
        Process one employee question through the full pipeline.

        Returns:
            {
              "answer"          : str
              "policy_name"     : str | None
              "policy_names"    : list | None   (MultiPolicyAgent)
              "found"           : bool
              "intent"          : str
              "escalated"       : bool
              "email_draft"     : str | None
              "compliance_alert": str | None
              "agent_used"      : str           (name of the agent that answered)
              "session_id"      : str
            }
        """
        start_time = time.time()

        # ── Step 1: Pre-query hook ────────────────────────────────────────────
        pre = pre_query_hook(question, self.session_id)
        if not pre["valid"]:
            return self._make_result(
                answer      = "Please type a valid question.",
                intent      = INTENT_OUT_OF_SCOPE,
                agent_used  = "OrchestratorAgent",
                found       = False,
            )

        question = pre["question"]

        # ── Step 2: Classify intent ───────────────────────────────────────────
        classification = self.classifier.run(question)
        intent         = classification["intent"]

        # ── Step 3: Route to correct subagent ────────────────────────────────
        result        = None
        compliance_alert_msg = None

        if intent == INTENT_SUMMARY:
            raw    = self.summarizer.run(question)
            result = self._make_result(
                answer      = raw["answer"],
                policy_name = raw.get("policy_name"),
                found       = raw["found"],
                intent      = intent,
                agent_used  = self.summarizer.name,
            )

        elif intent == INTENT_ESCALATION:
            raw    = self.escalator.run(question)
            result = self._make_result(
                answer      = raw["answer"],
                policy_name = raw.get("policy_name"),
                found       = raw["found"],
                intent      = intent,
                agent_used  = self.escalator.name,
                escalated   = True,
                email_draft = raw.get("email_draft"),
            )

        elif intent == INTENT_OUT_OF_SCOPE:
            result = self._make_result(
                answer     = "This question is outside the scope of bank policy documents.\n\nPlease contact **compliance@bank.internal** if you need help.",
                intent     = intent,
                agent_used = "OrchestratorAgent",
                found      = False,
            )

        else:
            # INTENT_POLICY_QUERY — check if multi-policy question
            cross = cross_reference(question, self.policies)

            if len(cross) >= 2 and cross[0]["score"] >= 2 and cross[1]["score"] >= 1:
                # Route to MultiPolicyAgent
                raw    = self.multi.run(question)
                if raw["found"]:
                    result = self._make_result(
                        answer       = raw["answer"],
                        policy_names = raw.get("policy_names"),
                        policy_name  = raw["policy_names"][0] if raw.get("policy_names") else None,
                        found        = True,
                        intent       = intent,
                        agent_used   = self.multi.name,
                    )
                else:
                    result = None   # fall through to fallback below

            if result is None:
                # Route to PolicySearchAgent
                raw = self.searcher.run(question)

                if raw["found"]:
                    result = self._make_result(
                        answer      = raw["answer"],
                        policy_name = raw.get("policy_name"),
                        found       = True,
                        intent      = intent,
                        agent_used  = self.searcher.name,
                    )
                else:
                    # Self-healing: FallbackAgent takes over
                    fb = self.fallback.run(question)
                    escalation_hook(question, self.session_id)
                    result = self._make_result(
                        answer     = fb["answer"],
                        found      = False,
                        intent     = intent,
                        agent_used = self.fallback.name,
                        escalated  = fb.get("needs_escalation", False),
                    )

        # ── Step 4: Audit logging (background) ───────────────────────────────
        audit = self.auditor.log_answer(
            question,
            result.get("policy_name"),
            result.get("found", False),
        )

        # ── Step 5: Compliance alert hook ─────────────────────────────────────
        if audit.get("compliance_risk"):
            alert = compliance_alert_hook(question, self.session_id)
            result["compliance_alert"] = alert["alert"]

        # ── Step 6: Escalation hook ───────────────────────────────────────────
        if not result.get("found") and not result.get("escalated"):
            escalation_hook(question, self.session_id)

        # ── Step 7: Post-answer hook ──────────────────────────────────────────
        elapsed_ms = (time.time() - start_time) * 1000
        post_answer_hook(question, result, self.session_id, elapsed_ms)

        result["session_id"] = self.session_id
        return result

    def on_session_end(self, total_queries: int) -> None:
        """Called when employee clears chat — fires session_end_hook."""
        session_end_hook(self.session_id, total_queries)

    # ── Helper ────────────────────────────────────────────────────────────────
    @staticmethod
    def _make_result(
        answer:           str,
        intent:           str           = "policy_query",
        agent_used:       str           = "OrchestratorAgent",
        found:            bool          = True,
        policy_name:      str | None    = None,
        policy_names:     list | None   = None,
        escalated:        bool          = False,
        email_draft:      str | None    = None,
        compliance_alert: str | None    = None,
        session_id:       str           = "",
    ) -> dict:
        return {
            "answer":           answer,
            "policy_name":      policy_name,
            "policy_names":     policy_names or [],
            "found":            found,
            "intent":           intent,
            "escalated":        escalated,
            "email_draft":      email_draft,
            "compliance_alert": compliance_alert,
            "agent_used":       agent_used,
            "session_id":       session_id,
        }
