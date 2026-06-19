"""
agents/multi_policy_agent.py
-----------------------------
SubAgent 3 — MultiPolicyAgent
Handles questions that span more than one policy domain.
Example: "What KYC documents do I need when applying for a loan?"
         → touches both KYC Policy AND Loan Policy
"""

from skills.search_skill import cross_reference
from skills.format_skill import POLICY_DISPLAY_NAMES


class MultiPolicyAgent:
    """
    SubAgent that checks and merges answers from multiple policies.
    Activated by the Orchestrator when cross_reference returns 2+ matches.
    """

    name = "MultiPolicyAgent"

    def __init__(self, policies: dict):
        self.policies = policies

    def run(self, question: str) -> dict:
        """
        Find all relevant policies and merge their snippets.

        Returns:
            {
              "answer"       : str   — merged answer from all matched policies
              "policy_names" : list  — all matched display names
              "found"        : bool
              "multi"        : bool  — True if more than 1 policy matched
            }
        """
        matches = cross_reference(question, self.policies)

        if not matches:
            return {
                "answer":       "No relevant policies found.",
                "policy_names": [],
                "found":        False,
                "multi":        False,
            }

        if len(matches) == 1:
            # Only one policy — no need for merging
            m = matches[0]
            display = POLICY_DISPLAY_NAMES.get(m["matched_key"], m["matched_key"])
            return {
                "answer":       m["snippet"],
                "policy_names": [display],
                "found":        True,
                "multi":        False,
            }

        # Multiple policies matched — merge top 2
        top_two = matches[:2]
        sections = []
        display_names = []

        for m in top_two:
            display = POLICY_DISPLAY_NAMES.get(m["matched_key"], m["matched_key"])
            display_names.append(display)
            sections.append(f"**From {display}:**\n{m['snippet']}")

        merged_answer = "\n\n---\n\n".join(sections)

        return {
            "answer":       merged_answer,
            "policy_names": display_names,
            "found":        True,
            "multi":        True,
        }
