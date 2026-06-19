"""
agents/policy_search_agent.py
------------------------------
SubAgent 2 — PolicySearchAgent
Searches the knowledge base and returns the best matching policy answer.
Core search logic lives here — extracted from utils/agent.py.
"""

from skills.search_skill import search_policy
from skills.format_skill import format_answer, POLICY_DISPLAY_NAMES

FALLBACK_MESSAGE = (
    "I could not find a relevant policy for your question.\n\n"
    "Please contact the **compliance or operations team** for assistance.\n\n"
    "📧 compliance@bank.internal | 📧 operations@bank.internal"
)


class PolicySearchAgent:
    """
    SubAgent that performs keyword-based policy search and returns a structured result.
    """

    name = "PolicySearchAgent"

    def __init__(self, policies: dict):
        self.policies = policies

    def run(self, question: str) -> dict:
        """
        Search all policies and return the best answer.

        Returns:
            {
              "answer"       : str   — formatted answer text
              "policy_key"   : str   — internal key e.g. 'loan_policy'
              "policy_name"  : str   — display name e.g. 'Loan Policy (LP-001)'
              "found"        : bool
              "score"        : int   — confidence score (keyword matches)
            }
        """
        result = search_policy(question, self.policies)

        if result["matched_key"] is None:
            return {
                "answer":      FALLBACK_MESSAGE,
                "policy_key":  None,
                "policy_name": None,
                "found":       False,
                "score":       0,
            }

        answer = format_answer(result["snippet"], result["matched_key"])
        display_name = POLICY_DISPLAY_NAMES.get(result["matched_key"], result["matched_key"])

        return {
            "answer":      answer,
            "policy_key":  result["matched_key"],
            "policy_name": display_name,
            "found":       True,
            "score":       result["score"],
        }
