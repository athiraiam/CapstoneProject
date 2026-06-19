"""
skills/format_skill.py
-----------------------
Reusable skill: formatting and summarising policy text.
Used by SummaryAgent and OrchestratorAgent.
"""


POLICY_DISPLAY_NAMES = {
    "loan_policy":               "Loan Policy (LP-001)",
    "kyc_policy":                "KYC Policy (KYC-002)",
    "customer_complaint_policy": "Customer Complaint Policy (CCP-003)",
    "credit_card_policy":        "Credit Card Policy (CCP-004)",
    "account_opening_policy":    "Account Opening Policy (AOP-005)",
}


def format_answer(raw_snippet: str, policy_key: str) -> str:
    """
    Take a raw text snippet and the policy key.
    Returns a cleanly formatted markdown string ready for display.
    """
    display_name = POLICY_DISPLAY_NAMES.get(policy_key, policy_key)
    lines = [l for l in raw_snippet.strip().split("\n") if l.strip()]
    formatted = "\n".join(f"{line}" for line in lines)
    return formatted


def summarize_policy(policy_text: str, policy_key: str) -> str:
    """
    Produce a 5-bullet summary of a policy document.
    Picks the most informative lines (non-empty, not just dashes or headers).
    Used by SummaryAgent on 'summarize' intent.
    """
    display_name = POLICY_DISPLAY_NAMES.get(policy_key, policy_key)
    lines = [l.strip() for l in policy_text.split("\n") if l.strip()]

    # Filter out pure header lines and separator lines
    info_lines = [
        l for l in lines
        if not l.startswith("─") and not l.startswith("=")
        and len(l) > 20 and not l.startswith("POLICY")
    ]

    # Pick 5 evenly spaced lines for a balanced summary
    if len(info_lines) >= 5:
        step = len(info_lines) // 5
        bullets = [info_lines[i * step] for i in range(5)]
    else:
        bullets = info_lines[:5]

    bullet_text = "\n".join(f"- {b}" for b in bullets)
    return f"**Summary of {display_name}:**\n\n{bullet_text}"


def generate_escalation_email(question: str, matched_policy: str | None) -> str:
    """
    Draft a pre-filled escalation email body for the compliance/operations team.
    Used by EscalationAgent and FallbackAgent.
    """
    policy_note = (
        f"Closest matching policy: {POLICY_DISPLAY_NAMES.get(matched_policy, 'None identified')}"
        if matched_policy
        else "No matching policy could be identified."
    )

    return (
        f"**Draft Escalation Email**\n\n"
        f"**To:** compliance@bank.internal / operations@bank.internal\n"
        f"**Subject:** Policy Query Escalation — Requires Clarification\n\n"
        f"Dear Team,\n\n"
        f"An employee has submitted the following policy query that could not be "
        f"fully resolved by the internal knowledge base:\n\n"
        f"> *{question}*\n\n"
        f"{policy_note}\n\n"
        f"Please review and respond to the employee at your earliest convenience.\n\n"
        f"Regards,\n"
        f"NexaBank Internal Assistant"
    )
