"""
knowledge_base.py
-----------------
Loads all policy text files from the /policies folder and provides
a search function that matches employee questions to the right policy.
"""

import os


# ── 1. Load all policy files into a dictionary ──────────────────────────────
def load_policies(policies_dir: str) -> dict[str, str]:
    """
    Reads every .txt file in policies_dir and returns:
        { "loan_policy": "full text ...", "kyc_policy": "full text ...", ... }
    """
    policies = {}
    for filename in os.listdir(policies_dir):
        if filename.endswith(".txt"):
            policy_name = filename.replace(".txt", "")          # e.g. "loan_policy"
            file_path = os.path.join(policies_dir, filename)
            with open(file_path, "r", encoding="utf-8") as f:
                policies[policy_name] = f.read()
    return policies


# ── 2. Keyword map – links common words to the correct policy ────────────────
KEYWORD_MAP = {
    "loan_policy": [
        "loan", "borrow", "lending", "personal loan", "home loan", "auto loan",
        "business loan", "interest rate", "tenure", "emi", "disbursement",
        "credit score", "prepayment", "collateral", "repayment", "mortgage"
    ],
    "kyc_policy": [
        "kyc", "know your customer", "identity", "verification", "id proof",
        "address proof", "document", "passport", "driver license", "national id",
        "aml", "anti money laundering", "due diligence", "onboarding",
        "identity verification", "id verification"
    ],
    "customer_complaint_policy": [
        "complaint", "complain", "grievance", "dispute", "unhappy", "dissatisfied",
        "escalate", "escalation", "feedback", "issue", "problem", "resolution",
        "refund", "compensation", "ombudsman", "resolve", "service complaint"
    ],
    "credit_card_policy": [
        "credit card", "card", "visa", "mastercard", "billing", "statement",
        "credit limit", "reward", "points", "cashback", "annual fee",
        "lost card", "stolen card", "block card", "minimum payment", "apr",
        "cash advance", "platinum card", "gold card", "classic card", "chargeback"
    ],
    "account_opening_policy": [
        "account", "open account", "savings account", "current account",
        "checking account", "fixed deposit", "joint account", "business account",
        "minimum balance", "debit card", "online banking", "dormant",
        "account closure", "welcome kit", "cheque book", "initial deposit"
    ],
}


# ── 3. Search function ───────────────────────────────────────────────────────
def search_policies(question: str, policies: dict[str, str]) -> str:
    """
    Matches the employee's question against the keyword map.
    Returns the relevant policy text, or a fallback message.

    Steps:
      1. Lowercase the question.
      2. Check each policy's keyword list for any match.
      3. Return the first matching policy's content.
      4. If no match, return the fallback message.
    """
    question_lower = question.lower()

    # Score each policy by how many keywords match
    scores = {}
    for policy_name, keywords in KEYWORD_MAP.items():
        score = sum(1 for kw in keywords if kw in question_lower)
        if score > 0:
            scores[policy_name] = score

    if not scores:
        # No policy matched → standard fallback
        return None

    # Pick the policy with the highest keyword match score
    best_match = max(scores, key=scores.get)
    return policies.get(best_match, None)


# ── 4. Extract a focused answer snippet from the policy text ─────────────────
def extract_answer(question: str, policy_text: str, context_lines: int = 15) -> str:
    """
    Instead of dumping the entire policy, tries to find the most relevant
    paragraph/section based on the question keywords, and returns that snippet.
    If nothing specific is found, returns the first `context_lines` lines.
    """
    question_words = set(question.lower().split())
    lines = policy_text.split("\n")

    best_line_idx = None
    best_score = 0

    for i, line in enumerate(lines):
        line_lower = line.lower()
        score = sum(1 for word in question_words if len(word) > 3 and word in line_lower)
        if score > best_score:
            best_score = score
            best_line_idx = i

    if best_line_idx is not None and best_score > 0:
        # Return a window of lines around the best matching line
        start = max(0, best_line_idx - 2)
        end = min(len(lines), best_line_idx + context_lines)
        snippet = "\n".join(lines[start:end])
        return snippet

    # Fallback: return the first `context_lines` lines of the policy
    return "\n".join(lines[:context_lines])
