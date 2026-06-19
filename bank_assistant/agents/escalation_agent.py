"""
agents/escalation_agent.py
---------------------------
SubAgent 5 — EscalationAgent
Handles all escalation requests and fallback cases.
Drafts a pre-filled email and shows the correct contact details.
Also handles self-healing: if PolicySearchAgent fails, Orchestrator
calls this agent automatically.
"""

from skills.format_skill import generate_escalation_email, POLICY_DISPLAY_NAMES
from skills.search_skill import search_policy

# Contact routing by policy domain
ESCALATION_CONTACTS = {
    "loan_policy":               ("Loan Operations Manager",    "loan.operations@bank.internal"),
    "kyc_policy":                ("Compliance Team",            "compliance@bank.internal"),
    "customer_complaint_policy": ("Customer Experience Team",   "customerexperience@bank.internal"),
    "credit_card_policy":        ("Credit Card Department",     "creditcards@bank.internal"),
    "account_opening_policy":    ("Account Operations Team",    "accountoperations@bank.internal"),
}

DEFAULT_CONTACT = ("Compliance / Operations Team",
                   "compliance@bank.internal | operations@bank.internal")


class EscalationAgent:
    """
    SubAgent that generates escalation guidance and email drafts.
    """

    name = "EscalationAgent"

    def __init__(self, policies: dict):
        self.policies = policies

    def run(self, question: str, matched_policy_key: str | None = None) -> dict:
        """
        Generate an escalation response with the correct contact and a draft email.

        Returns:
            {
              "answer"       : str  — escalation message shown in chat
              "policy_name"  : str | None
              "found"        : bool (always True — escalation IS a valid response)
              "escalated"    : bool
              "email_draft"  : str  — pre-filled email body
            }
        """
        # If no policy key given, try to infer from the question
        if not matched_policy_key:
            result = search_policy(question, self.policies)
            matched_policy_key = result.get("matched_key")

        # Pick the right contact team
        if matched_policy_key and matched_policy_key in ESCALATION_CONTACTS:
            team_name, email = ESCALATION_CONTACTS[matched_policy_key]
        else:
            team_name, email = DEFAULT_CONTACT

        display_name = POLICY_DISPLAY_NAMES.get(matched_policy_key, None) if matched_policy_key else None
        email_draft  = generate_escalation_email(question, matched_policy_key)

        answer = (
            f"This query requires human review.\n\n"
            f"**Please contact:** {team_name}\n"
            f"**Email:** {email}\n\n"
            f"A draft escalation email has been prepared for you below."
        )

        return {
            "answer":      answer,
            "policy_name": display_name,
            "found":       True,
            "escalated":   True,
            "email_draft": email_draft,
        }
