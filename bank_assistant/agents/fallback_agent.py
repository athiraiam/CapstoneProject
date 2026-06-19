"""
agents/fallback_agent.py
-------------------------
SubAgent 6 — FallbackAgent
Self-healing agent: when PolicySearchAgent returns no match,
FallbackAgent tries to suggest the closest policy and
automatically triggers EscalationAgent for contact routing.
"""

from skills.search_skill import cross_reference
from skills.format_skill import POLICY_DISPLAY_NAMES


class FallbackAgent:
    """
    Self-healing subagent — activated automatically when no policy matches.
    Tries to recover by finding partial matches and suggests next steps.
    """

    name = "FallbackAgent"

    def __init__(self, policies: dict):
        self.policies = policies

    def run(self, question: str) -> dict:
        """
        Attempt recovery from a failed policy search.

        1. Try cross-reference for any partial match
        2. If partial match found → suggest it
        3. If truly nothing → instruct escalation

        Returns:
            {
              "answer"          : str
              "suggestion"      : str | None  — suggested policy name if partial match found
              "needs_escalation": bool
              "found"           : bool (always False — this agent is the fallback)
            }
        """
        # Try partial cross-reference (might find something at low score)
        partials = cross_reference(question, self.policies)

        if partials:
            # Found a low-confidence partial match — suggest it
            best = partials[0]
            display = POLICY_DISPLAY_NAMES.get(best["matched_key"], best["matched_key"])
            answer = (
                f"I couldn't find an exact match for your question.\n\n"
                f"The closest policy I found is the **{display}**. "
                f"It may contain relevant information.\n\n"
                f"If this doesn't help, please contact the compliance or operations team."
            )
            return {
                "answer":           answer,
                "suggestion":       display,
                "needs_escalation": True,
                "found":            False,
            }

        # Truly nothing found
        return {
            "answer": (
                "I could not find any relevant information in the policy knowledge base.\n\n"
                "Please contact the **Compliance or Operations team** directly:\n"
                "📧 compliance@bank.internal | 📧 operations@bank.internal"
            ),
            "suggestion":       None,
            "needs_escalation": True,
            "found":            False,
        }
