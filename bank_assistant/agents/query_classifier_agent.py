"""
agents/query_classifier_agent.py
----------------------------------
SubAgent 1 — QueryClassifierAgent
Classifies every incoming question into one of four intents BEFORE search.
This tells the Orchestrator which downstream agent to invoke.
"""

from skills.search_skill import extract_keywords

# Intent labels
INTENT_POLICY_QUERY  = "policy_query"    # Normal policy search
INTENT_SUMMARY       = "summary"         # Employee wants a full policy summary
INTENT_ESCALATION    = "escalation"      # Wants to escalate / contact team
INTENT_OUT_OF_SCOPE  = "out_of_scope"    # Nothing to do with banking policies

# Keywords that signal each intent
SUMMARY_TRIGGERS = [
    "summarize", "summary", "overview", "brief", "explain the policy",
    "what does the policy say", "give me an overview", "tell me about the policy",
]

ESCALATION_TRIGGERS = [
    "escalate", "contact", "speak to", "human", "manager", "compliance team",
    "operations team", "raise a complaint", "who do i call", "email address",
    "get help", "urgent", "not resolved",
]


class QueryClassifierAgent:
    """
    SubAgent that classifies employee questions into intents.
    Runs first in the orchestrator pipeline.
    """

    name = "QueryClassifierAgent"

    def run(self, question: str) -> dict:
        """
        Classify the question.

        Returns:
            {
              "intent"    : str    — one of the INTENT_* constants
              "keywords"  : list   — extracted meaningful words
              "question"  : str    — original question passed through
            }
        """
        q_lower = question.lower().strip()

        if not q_lower:
            return {"intent": INTENT_OUT_OF_SCOPE, "keywords": [], "question": question}

        # Check summary intent
        if any(trigger in q_lower for trigger in SUMMARY_TRIGGERS):
            return {
                "intent":   INTENT_SUMMARY,
                "keywords": extract_keywords(question),
                "question": question,
            }

        # Check escalation intent
        if any(trigger in q_lower for trigger in ESCALATION_TRIGGERS):
            return {
                "intent":   INTENT_ESCALATION,
                "keywords": extract_keywords(question),
                "question": question,
            }

        # Default: treat as a policy query
        return {
            "intent":   INTENT_POLICY_QUERY,
            "keywords": extract_keywords(question),
            "question": question,
        }
