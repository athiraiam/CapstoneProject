"""
skills/search_skill.py
----------------------
Reusable skill: keyword search + snippet extraction from the policy knowledge base.
Used by PolicySearchAgent and MultiPolicyAgent.
"""

from utils.knowledge_base import search_policies, extract_answer, KEYWORD_MAP


def search_policy(question: str, policies: dict) -> dict:
    """
    Search all policies and return the best matching result.

    Returns:
        {
          "matched_key"  : str | None  — internal policy file key
          "policy_text"  : str | None  — full policy text
          "snippet"      : str | None  — focused answer paragraph
          "score"        : int         — keyword match score
        }
    """
    question_lower = question.lower()

    scores = {}
    for policy_name, keywords in KEYWORD_MAP.items():
        score = sum(1 for kw in keywords if kw in question_lower)
        if score > 0:
            scores[policy_name] = score

    if not scores:
        return {"matched_key": None, "policy_text": None, "snippet": None, "score": 0}

    best_key = max(scores, key=scores.get)
    policy_text = policies.get(best_key, "")
    snippet = extract_answer(question, policy_text, context_lines=20)

    return {
        "matched_key": best_key,
        "policy_text": policy_text,
        "snippet":     snippet,
        "score":       scores[best_key],
    }


def extract_keywords(text: str) -> list[str]:
    """
    Pull meaningful keywords (length > 3) from a question string.
    Used by QueryClassifierAgent to understand intent.
    """
    stopwords = {"what", "when", "where", "which", "have", "does", "will",
                 "that", "this", "with", "from", "about", "there", "their",
                 "they", "them", "then", "than", "into", "your", "some"}
    words = text.lower().replace("?", "").replace(",", "").split()
    return [w for w in words if len(w) > 3 and w not in stopwords]


def cross_reference(question: str, policies: dict) -> list[dict]:
    """
    Check if the question spans multiple policies.
    Returns a list of all policies with score > 0, sorted by score descending.
    Used by MultiPolicyAgent.
    """
    question_lower = question.lower()
    results = []

    for policy_name, keywords in KEYWORD_MAP.items():
        score = sum(1 for kw in keywords if kw in question_lower)
        if score > 0:
            results.append({
                "matched_key": policy_name,
                "score":       score,
                "snippet":     extract_answer(question, policies.get(policy_name, ""), context_lines=12),
            })

    results.sort(key=lambda x: x["score"], reverse=True)
    return results
