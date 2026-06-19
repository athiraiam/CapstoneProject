"""
agent.py
--------
The core AI agent logic.
Takes an employee's question, searches the knowledge base,
and returns a structured answer with the source policy name.
"""

from utils.knowledge_base import search_policies, extract_answer

# Fallback message when no policy matches the question
FALLBACK_MESSAGE = (
    "I'm sorry, I could not find a relevant policy for your question.\n\n"
    "Please contact the **compliance or operations team** for further assistance.\n\n"
    "📧 compliance@bank.internal | 📧 operations@bank.internal"
)

# Friendly display names for each policy file
POLICY_DISPLAY_NAMES = {
    "loan_policy":               "Loan Policy (LP-001)",
    "kyc_policy":                "KYC Policy (KYC-002)",
    "customer_complaint_policy": "Customer Complaint Policy (CCP-003)",
    "credit_card_policy":        "Credit Card Policy (CCP-004)",
    "account_opening_policy":    "Account Opening Policy (AOP-005)",
}


def get_answer(question: str, policies: dict[str, str]) -> dict:
    """
    Main agent function.

    Args:
        question : The employee's question string.
        policies : Dictionary of { policy_name: full_text } loaded at startup.

    Returns a dict:
        {
            "answer"      : str   – the answer text to display,
            "policy_name" : str   – human-readable source policy name,
            "found"       : bool  – True if a matching policy was found,
        }
    """
    if not question.strip():
        return {
            "answer": "Please type a question to get started.",
            "policy_name": None,
            "found": False,
        }

    # Step 1: Find the matching policy text
    matched_policy_text = search_policies(question, policies)

    if matched_policy_text is None:
        return {
            "answer": FALLBACK_MESSAGE,
            "policy_name": None,
            "found": False,
        }

    # Step 2: Find which policy key matched (to show the display name)
    from utils.knowledge_base import KEYWORD_MAP
    question_lower = question.lower()
    matched_key = None
    best_score = 0
    for policy_name, keywords in KEYWORD_MAP.items():
        score = sum(1 for kw in keywords if kw in question_lower)
        if score > best_score:
            best_score = score
            matched_key = policy_name

    # Step 3: Extract a focused snippet instead of the full policy
    snippet = extract_answer(question, matched_policy_text, context_lines=20)

    display_name = POLICY_DISPLAY_NAMES.get(matched_key, matched_key)

    answer = f"{snippet}"

    return {
        "answer": answer,
        "policy_name": display_name,
        "found": True,
    }
