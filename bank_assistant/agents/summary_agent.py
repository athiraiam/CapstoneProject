"""
agents/summary_agent.py
------------------------
SubAgent 4 — SummaryAgent
Generates a concise 5-bullet summary of any policy document on demand.
Triggered when the QueryClassifier detects a 'summary' intent.
Example: "Give me a summary of the loan policy"
"""

from skills.search_skill import search_policy
from skills.format_skill import summarize_policy, POLICY_DISPLAY_NAMES

# Map common name fragments to policy keys
POLICY_NAME_HINTS = {
    "loan":      "loan_policy",
    "kyc":       "kyc_policy",
    "complaint": "customer_complaint_policy",
    "credit":    "credit_card_policy",
    "card":      "credit_card_policy",
    "account":   "account_opening_policy",
}


class SummaryAgent:
    """
    SubAgent that summarizes a named policy into 5 bullet points.
    """

    name = "SummaryAgent"

    def __init__(self, policies: dict):
        self.policies = policies

    def _detect_policy_key(self, question: str) -> str | None:
        """Detect which policy the employee is asking about from the question text."""
        q_lower = question.lower()
        for hint, key in POLICY_NAME_HINTS.items():
            if hint in q_lower:
                return key

        # Fall back to keyword search if no explicit name found
        result = search_policy(question, self.policies)
        return result.get("matched_key")

    def run(self, question: str) -> dict:
        """
        Produce a summary of the best-matched policy.

        Returns:
            {
              "answer"      : str   — 5-bullet summary markdown
              "policy_name" : str   — display name of the summarized policy
              "found"       : bool
            }
        """
        policy_key = self._detect_policy_key(question)

        if not policy_key or policy_key not in self.policies:
            return {
                "answer":      "I could not identify which policy you'd like summarized. Please name the policy explicitly (e.g. 'Summarize the loan policy').",
                "policy_name": None,
                "found":       False,
            }

        policy_text  = self.policies[policy_key]
        summary      = summarize_policy(policy_text, policy_key)
        display_name = POLICY_DISPLAY_NAMES.get(policy_key, policy_key)

        return {
            "answer":      summary,
            "policy_name": display_name,
            "found":       True,
        }
