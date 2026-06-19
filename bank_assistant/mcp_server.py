"""
mcp_server.py
-------------
MCP (Model Context Protocol) server for the Internal Bank Employee Assistant.

Exposes the policy knowledge base as MCP tools so any MCP-compatible client
(Claude Desktop, Claude Code, custom agents) can query bank policies directly.

Tools exposed:
  - search_policy        : keyword search → best matching policy snippet
  - list_policies        : list all available policy names and codes
  - get_policy_full      : return the complete text of one policy
  - cross_reference      : find which policies match a multi-topic question
  - check_compliance_risk: flag whether a question contains risk keywords

Run standalone:
    python mcp_server.py

Or register it in claude_mcp_config.json for Claude Desktop / Claude Code.
"""

import json
import sys
import os

# Allow imports from the bank_assistant package
sys.path.insert(0, os.path.dirname(__file__))

from utils.knowledge_base import load_policies, extract_answer, KEYWORD_MAP
from skills.search_skill import search_policy, cross_reference
from skills.audit_skill import detect_compliance_risk

# ── Load policies once at startup ─────────────────────────────────────────────
POLICIES_DIR = os.path.join(os.path.dirname(__file__), "policies")
POLICIES = load_policies(POLICIES_DIR)

POLICY_DISPLAY_NAMES = {
    "loan_policy":              ("Loan Policy",             "LP-001"),
    "kyc_policy":               ("KYC Policy",              "KYC-002"),
    "customer_complaint_policy":("Customer Complaint Policy","CCP-003"),
    "credit_card_policy":       ("Credit Card Policy",      "CCP-004"),
    "account_opening_policy":   ("Account Opening Policy",  "AOP-005"),
}

# ── MCP Tool Definitions ──────────────────────────────────────────────────────
TOOLS = [
    {
        "name": "search_policy",
        "description": (
            "Search the NexaBank internal policy knowledge base for an answer "
            "to an employee question. Returns the most relevant policy snippet "
            "and the source policy name."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "question": {
                    "type": "string",
                    "description": "The employee's policy question in natural language."
                }
            },
            "required": ["question"],
        },
    },
    {
        "name": "list_policies",
        "description": "List all available bank policy documents with their codes and topics.",
        "input_schema": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "get_policy_full",
        "description": "Return the complete text of a specific bank policy document.",
        "input_schema": {
            "type": "object",
            "properties": {
                "policy_code": {
                    "type": "string",
                    "description": (
                        "The policy code. One of: LP-001, KYC-002, CCP-003, CCP-004, AOP-005."
                    )
                }
            },
            "required": ["policy_code"],
        },
    },
    {
        "name": "cross_reference",
        "description": (
            "Check whether a question spans multiple bank policies and return "
            "relevant snippets from each matching policy. Useful for complex "
            "questions that touch more than one domain (e.g. loan + KYC)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "question": {
                    "type": "string",
                    "description": "A question that may span multiple policy domains."
                }
            },
            "required": ["question"],
        },
    },
    {
        "name": "check_compliance_risk",
        "description": (
            "Check whether an employee's question contains compliance risk keywords "
            "(fraud, money laundering, bypass KYC, etc.). Returns a risk flag and "
            "the matched keywords."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "question": {
                    "type": "string",
                    "description": "The question to screen for compliance risk."
                }
            },
            "required": ["question"],
        },
    },
]

# ── Tool Handlers ─────────────────────────────────────────────────────────────

def handle_search_policy(question: str) -> dict:
    result = search_policy(question, POLICIES)
    if not result["matched_key"]:
        return {
            "found": False,
            "answer": (
                "No matching policy found. "
                "Please contact compliance@bank.internal for assistance."
            ),
            "policy_name": None,
            "policy_code": None,
            "score": 0,
        }
    key = result["matched_key"]
    display_name, code = POLICY_DISPLAY_NAMES.get(key, (key, "—"))
    return {
        "found": True,
        "answer": result["snippet"],
        "policy_name": display_name,
        "policy_code": code,
        "score": result["score"],
    }


def handle_list_policies() -> dict:
    policies_list = []
    for key, (name, code) in POLICY_DISPLAY_NAMES.items():
        keywords_sample = KEYWORD_MAP.get(key, [])[:5]
        policies_list.append({
            "code": code,
            "name": name,
            "sample_topics": keywords_sample,
        })
    return {"policies": policies_list, "total": len(policies_list)}


CODE_TO_KEY = {
    "LP-001":  "loan_policy",
    "KYC-002": "kyc_policy",
    "CCP-003": "customer_complaint_policy",
    "CCP-004": "credit_card_policy",
    "AOP-005": "account_opening_policy",
}

def handle_get_policy_full(policy_code: str) -> dict:
    key = CODE_TO_KEY.get(policy_code.upper())
    if not key:
        return {
            "found": False,
            "error": f"Unknown policy code: {policy_code}. Valid codes: {list(CODE_TO_KEY.keys())}",
        }
    text = POLICIES.get(key, "")
    display_name, code = POLICY_DISPLAY_NAMES.get(key, (key, "—"))
    return {
        "found": True,
        "policy_name": display_name,
        "policy_code": code,
        "text": text,
        "char_count": len(text),
    }


def handle_cross_reference(question: str) -> dict:
    results = cross_reference(question, POLICIES)
    if not results:
        return {"found": False, "matches": [], "total_matches": 0}
    matches = []
    for r in results:
        key = r["matched_key"]
        display_name, code = POLICY_DISPLAY_NAMES.get(key, (key, "—"))
        matches.append({
            "policy_name": display_name,
            "policy_code": code,
            "score": r["score"],
            "snippet": r["snippet"],
        })
    return {"found": True, "matches": matches, "total_matches": len(matches)}


def handle_check_compliance_risk(question: str) -> dict:
    risk_keywords = [
        "fraud", "money laundering", "suspicious", "bypass", "override",
        "without kyc", "skip verification", "fake document", "bribe",
        "corrupt", "illegal", "sanction", "blacklist", "terrorist",
    ]
    q_lower = question.lower()
    matched = [kw for kw in risk_keywords if kw in q_lower]
    is_risk = len(matched) > 0
    return {
        "is_compliance_risk": is_risk,
        "matched_keywords": matched,
        "alert": (
            "⚠️ Compliance risk detected. This query has been flagged. "
            "Contact compliance@bank.internal immediately."
        ) if is_risk else None,
    }


# ── MCP JSON-RPC Dispatcher ───────────────────────────────────────────────────

def dispatch(tool_name: str, tool_input: dict) -> dict:
    if tool_name == "search_policy":
        return handle_search_policy(tool_input["question"])
    elif tool_name == "list_policies":
        return handle_list_policies()
    elif tool_name == "get_policy_full":
        return handle_get_policy_full(tool_input["policy_code"])
    elif tool_name == "cross_reference":
        return handle_cross_reference(tool_input["question"])
    elif tool_name == "check_compliance_risk":
        return handle_check_compliance_risk(tool_input["question"])
    else:
        return {"error": f"Unknown tool: {tool_name}"}


def handle_request(request: dict) -> dict:
    method  = request.get("method", "")
    req_id  = request.get("id")
    params  = request.get("params", {})

    if method == "initialize":
        return {
            "jsonrpc": "2.0", "id": req_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {}},
                "serverInfo": {
                    "name":    "nexabank-policy-assistant",
                    "version": "1.0.0",
                },
            },
        }

    elif method == "tools/list":
        return {
            "jsonrpc": "2.0", "id": req_id,
            "result": {"tools": TOOLS},
        }

    elif method == "tools/call":
        tool_name  = params.get("name", "")
        tool_input = params.get("arguments", {})
        result     = dispatch(tool_name, tool_input)
        return {
            "jsonrpc": "2.0", "id": req_id,
            "result": {
                "content": [{"type": "text", "text": json.dumps(result, indent=2)}],
                "isError": "error" in result,
            },
        }

    elif method == "notifications/initialized":
        return None  # no response needed for notifications

    else:
        return {
            "jsonrpc": "2.0", "id": req_id,
            "error": {"code": -32601, "message": f"Method not found: {method}"},
        }


# ── stdio transport (MCP standard) ───────────────────────────────────────────
def run_stdio():
    """Run the MCP server over stdin/stdout (the standard MCP transport)."""
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request  = json.loads(line)
            response = handle_request(request)
            if response is not None:
                print(json.dumps(response), flush=True)
        except json.JSONDecodeError as e:
            error_response = {
                "jsonrpc": "2.0", "id": None,
                "error": {"code": -32700, "message": f"Parse error: {e}"},
            }
            print(json.dumps(error_response), flush=True)


if __name__ == "__main__":
    run_stdio()
