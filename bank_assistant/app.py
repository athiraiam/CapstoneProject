"""
app.py
------
Streamlit chatbot UI for the Internal Bank Employee Assistant.
Run with:  streamlit run app.py
"""

import os
import base64
import uuid
import streamlit as st
from utils.knowledge_base import load_policies
from agents.orchestrator import OrchestratorAgent


def svg_to_b64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def img(b64: str, h: int = 24) -> str:
    return f'<img src="data:image/svg+xml;base64,{b64}" height="{h}" style="vertical-align:middle;" />'


ASSETS = os.path.join(os.path.dirname(__file__), "assets")

# Load all SVG assets
LOGO_B64        = svg_to_b64(os.path.join(ASSETS, "bank_logo.svg"))
SHIELD_B64      = svg_to_b64(os.path.join(ASSETS, "shield_badge.svg"))
BOT_AVATAR_B64  = svg_to_b64(os.path.join(ASSETS, "bot_avatar.svg"))
USER_AVATAR_B64 = svg_to_b64(os.path.join(ASSETS, "user_avatar.svg"))
ICON_LOAN_B64   = svg_to_b64(os.path.join(ASSETS, "icon_loan.svg"))
ICON_KYC_B64    = svg_to_b64(os.path.join(ASSETS, "icon_kyc.svg"))
ICON_COMP_B64   = svg_to_b64(os.path.join(ASSETS, "icon_complaint.svg"))
ICON_CARD_B64   = svg_to_b64(os.path.join(ASSETS, "icon_card.svg"))
ICON_ACC_B64    = svg_to_b64(os.path.join(ASSETS, "icon_account.svg"))
ICON_DOC_B64    = svg_to_b64(os.path.join(ASSETS, "icon_policy_doc.svg"))
ICON_WARN_B64   = svg_to_b64(os.path.join(ASSETS, "icon_fallback.svg"))
ICON_TEL_B64    = svg_to_b64(os.path.join(ASSETS, "icon_contact.svg"))
ICON_Q_B64      = svg_to_b64(os.path.join(ASSETS, "icon_questions.svg"))

BOT_AVATAR_URI  = f"data:image/svg+xml;base64,{BOT_AVATAR_B64}"
USER_AVATAR_URI = f"data:image/svg+xml;base64,{USER_AVATAR_B64}"

# ── Page configuration ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="NexaBank — Employee Assistant",
    page_icon=f"data:image/svg+xml;base64,{SHIELD_B64}",
    layout="wide",
)

# ── Load policies ─────────────────────────────────────────────────────────────
@st.cache_resource
def init_knowledge_base():
    policies_dir = os.path.join(os.path.dirname(__file__), "policies")
    return load_policies(policies_dir)

policies = init_knowledge_base()

# ── Session ID (unique per browser session) ───────────────────────────────────
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())[:8]

# ── Orchestrator — one instance per session ───────────────────────────────────
@st.cache_resource
def get_orchestrator(_policies: dict, _session_id: str) -> OrchestratorAgent:
    return OrchestratorAgent(_policies, _session_id)

orchestrator = get_orchestrator(policies, st.session_state.session_id)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Header ── */
    .header-bar {
        background: linear-gradient(90deg, #001f4d 0%, #0055b3 60%, #0077e6 100%);
        border-radius: 14px;
        margin-bottom: 22px;
        overflow: hidden;
        box-shadow: 0 4px 18px rgba(0,51,102,0.18);
    }
    .header-inner {
        display: flex; align-items: center;
        padding: 18px 28px; gap: 20px;
    }
    .header-divider {
        width: 3px; height: 52px;
        background: rgba(255,255,255,0.25);
        border-radius: 2px; flex-shrink: 0;
    }
    .header-text h1 {
        color: white; margin: 0;
        font-size: 1.55rem; font-weight: 800; letter-spacing: -0.3px;
    }
    .header-text p {
        color: #b3d1ff; margin: 3px 0 0 0; font-size: 0.88rem;
    }
    .header-badge {
        display: flex; align-items: center; gap: 6px;
        background: rgba(255,255,255,0.10);
        border: 1px solid rgba(255,255,255,0.20);
        border-radius: 20px; padding: 5px 14px;
        margin-left: auto;
        color: #FFC107; font-size: 0.78rem; font-weight: 700;
        white-space: nowrap;
    }
    .header-pill {
        display: flex; align-items: center; gap: 5px;
        background: rgba(255,255,255,0.08);
        border-radius: 20px; padding: 3px 10px;
        color: #99ccff; font-size: 0.72rem;
        margin-top: 6px; width: fit-content;
    }

    /* ── Sidebar ── */
    .sidebar-logo-wrap {
        background: linear-gradient(150deg, #001f4d 0%, #004499 100%);
        padding: 18px 14px 14px 14px;
        border-radius: 12px; margin-bottom: 18px;
        text-align: center;
        box-shadow: inset 0 -2px 8px rgba(0,0,0,0.2);
    }
    .sidebar-tagline {
        color: #7aaad9; font-size: 0.68rem;
        letter-spacing: 2px; margin-top: 7px;
        text-transform: uppercase;
    }
    .sidebar-divider-label {
        display: flex; align-items: center; gap: 8px;
        color: #003366; font-weight: 700;
        font-size: 0.85rem; margin: 4px 0 8px 0;
    }
    .policy-card {
        display: flex; align-items: center; gap: 10px;
        background: #f4f7fb;
        border: 1px solid #dde6f0;
        padding: 10px 12px;
        border-radius: 9px; margin-bottom: 8px;
        transition: background 0.2s;
    }
    .policy-card:hover { background: #e6eef8; }
    .policy-card-text strong { font-size: 0.87rem; color: #002a6e; }
    .policy-card-text small  { font-size: 0.75rem; color: #5577aa; display: block; }

    .sample-q-header {
        display: flex; align-items: center; gap: 8px;
        color: #003366; font-weight: 700;
        font-size: 0.85rem; margin: 4px 0 8px 0;
    }

    .sidebar-footer {
        background: linear-gradient(135deg, #001f4d 0%, #003d99 100%);
        border-radius: 10px; padding: 12px;
        text-align: center; margin-top: 8px;
    }
    .sidebar-footer-text {
        color: #6699cc; font-size: 0.68rem;
        margin-top: 6px; letter-spacing: 0.5px;
    }

    /* ── Policy source badge ── */
    .policy-badge {
        display: flex; align-items: center; gap: 8px;
        background: linear-gradient(90deg, #e8f2ff 0%, #f0f6ff 100%);
        border-left: 4px solid #0066CC;
        padding: 7px 14px; border-radius: 6px;
        font-size: 0.84rem; color: #002966;
        font-weight: 600; margin-bottom: 10px;
    }

    /* ── Fallback / warning box ── */
    .fallback-box {
        display: flex; align-items: flex-start; gap: 10px;
        background: #fff8e6;
        border-left: 4px solid #FFC107;
        padding: 12px 16px; border-radius: 6px;
        color: #7a5000;
    }
    .fallback-box .fb-text { flex: 1; }

    /* ── Welcome panel ── */
    .welcome-panel {
        background: linear-gradient(135deg, #f0f6ff 0%, #e8f2ff 100%);
        border: 1px solid #c0d8f8;
        border-radius: 12px; padding: 18px 22px;
        margin-bottom: 6px;
    }
    .welcome-panel h4 { color: #002966; margin: 0 0 8px 0; font-size: 1rem; }
    .topic-grid {
        display: grid; grid-template-columns: 1fr 1fr;
        gap: 7px; margin-top: 8px;
    }
    .topic-chip {
        display: flex; align-items: center; gap: 7px;
        background: white; border: 1px solid #cde0ff;
        border-radius: 20px; padding: 5px 10px;
        font-size: 0.8rem; color: #003380;
    }

    /* ── Chat ── */
    .stChatMessage { border-radius: 12px; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="header-bar">
  <div class="header-inner">
    <img src="data:image/svg+xml;base64,{LOGO_B64}" height="54" alt="NexaBank"/>
    <div class="header-divider"></div>
    <div class="header-text">
      <h1>Internal Employee Assistant</h1>
      <p>Instant answers from the official NexaBank policy knowledge base</p>
      <div class="header-pill">
        {img(ICON_DOC_B64, 14)}&nbsp;5 Policies Loaded &nbsp;·&nbsp;
        {img(SHIELD_B64, 14)}&nbsp;Offline &amp; Secure
      </div>
    </div>
    <div class="header-badge">
      {img(SHIELD_B64, 18)}&nbsp;Internal Use Only
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:

    # Logo block
    st.markdown(f"""
    <div class="sidebar-logo-wrap">
      <img src="data:image/svg+xml;base64,{LOGO_B64}" width="175" alt="NexaBank"/>
      <div class="sidebar-tagline">&#128274; Employee Policy Portal</div>
    </div>
    """, unsafe_allow_html=True)

    # Available Policies section
    st.markdown(f"""
    <div class="sidebar-divider-label">
      {img(ICON_DOC_B64, 18)}&nbsp; Available Policies
    </div>
    """, unsafe_allow_html=True)

    policy_cards = [
        (ICON_LOAN_B64, "Loan Policy",            "LP-001",   "Eligibility, types, documents, approval"),
        (ICON_KYC_B64,  "KYC Policy",             "KYC-002",  "Identity verification, documents"),
        (ICON_COMP_B64, "Customer Complaint",      "CCP-003",  "Complaint channels & escalation"),
        (ICON_CARD_B64, "Credit Card Policy",      "CCP-004",  "Issuance, limits, billing, disputes"),
        (ICON_ACC_B64,  "Account Opening Policy",  "AOP-005",  "Account types, requirements, process"),
    ]
    for icon_b64, name, code, desc in policy_cards:
        st.markdown(f"""
        <div class="policy-card">
          <img src="data:image/svg+xml;base64,{icon_b64}" height="32" alt="{name}"/>
          <div class="policy-card-text">
            <strong>{name}</strong>
            <small>{code} — {desc}</small>
          </div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # Sample questions section
    st.markdown(f"""
    <div class="sample-q-header">
      {img(ICON_Q_B64, 18)}&nbsp; Sample Questions
    </div>
    """, unsafe_allow_html=True)

    sample_questions = [
        "What documents are needed for a loan?",
        "What is the minimum credit score for a credit card?",
        "How long does KYC verification take?",
        "How do I file a customer complaint?",
        "What is the minimum balance for a savings account?",
        "What is the interest rate for a home loan?",
    ]
    for i, q in enumerate(sample_questions):
        if st.button(q, use_container_width=True, key=f"sample_{i}"):
            st.session_state.pending_question = q

    st.divider()

    if st.button("🗑️ Clear Chat History", use_container_width=True):
        total = len([m for m in st.session_state.messages if m["role"] == "user"])
        orchestrator.on_session_end(total)
        st.session_state.messages = []
        st.rerun()

    # Sidebar footer branding
    st.markdown(f"""
    <div class="sidebar-footer">
      {img(SHIELD_B64, 28)}<br>
      <div class="sidebar-footer-text">
        NexaBank Internal Systems<br>
        Powered by AI Policy Engine<br>
        © 2024 NexaBank
      </div>
    </div>
    """, unsafe_allow_html=True)


# ── Policy icon map (used in source badge) ────────────────────────────────────
POLICY_ICON_MAP = {
    "Loan Policy":              ICON_LOAN_B64,
    "KYC Policy":               ICON_KYC_B64,
    "Customer Complaint Policy":ICON_COMP_B64,
    "Credit Card Policy":       ICON_CARD_B64,
    "Account Opening Policy":   ICON_ACC_B64,
}

def policy_icon(policy_name: str) -> str:
    for key, b64 in POLICY_ICON_MAP.items():
        if key.lower() in (policy_name or "").lower():
            return img(b64, 20)
    return img(ICON_DOC_B64, 20)


# ── Initialize chat history ───────────────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = []
    st.session_state.messages.append({
        "role": "assistant",
        "content": "__welcome__",
        "policy_name": None,
        "found": True,
    })

# ── Render chat messages ──────────────────────────────────────────────────────
for message in st.session_state.messages:
    avatar = USER_AVATAR_URI if message["role"] == "user" else BOT_AVATAR_URI
    with st.chat_message(message["role"], avatar=avatar):

        if message["content"] == "__welcome__":
            # Rich welcome panel with icons per topic
            topic_items = [
                (ICON_LOAN_B64, "Loan Policies"),
                (ICON_KYC_B64,  "KYC Requirements"),
                (ICON_COMP_B64, "Complaint Procedures"),
                (ICON_CARD_B64, "Credit Card Policies"),
                (ICON_ACC_B64,  "Account Opening"),
            ]
            chips = "".join(
                f'<div class="topic-chip">{img(b64, 18)}&nbsp;{label}</div>'
                for b64, label in topic_items
            )
            st.markdown(f"""
            <div class="welcome-panel">
              <h4>{img(BOT_AVATAR_B64, 26)}&nbsp; Hello! I'm your NexaBank Policy Assistant.</h4>
              I can answer questions about official bank policies. Select a topic or type your question below.
              <div class="topic-grid">{chips}</div>
            </div>
            """, unsafe_allow_html=True)

        else:
            if message["role"] == "assistant":
                # Compliance alert banner
                if message.get("compliance_alert"):
                    st.warning(message["compliance_alert"])

                # Policy source badge
                names = message.get("policy_names") or (
                    [message["policy_name"]] if message.get("policy_name") else []
                )
                if names:
                    badges = "&nbsp;·&nbsp;".join(
                        f'{policy_icon(n)}&nbsp;<strong>{n}</strong>' for n in names
                    )
                    st.markdown(
                        f'<div class="policy-badge">{img(ICON_DOC_B64, 18)}&nbsp;Source:&nbsp;{badges}</div>',
                        unsafe_allow_html=True,
                    )

                # Agent used chip
                if message.get("agent_used"):
                    st.markdown(
                        f'<small style="color:#6699aa;">🤖 Agent: <code>{message["agent_used"]}</code>'
                        f'&nbsp;|&nbsp;Intent: <code>{message.get("intent","—")}</code></small>',
                        unsafe_allow_html=True,
                    )

                # Fallback warning box
                if not message.get("found"):
                    st.markdown(f"""
                    <div class="fallback-box">
                      {img(ICON_WARN_B64, 28)}
                      <div class="fb-text">
                        <strong>No exact policy match found.</strong><br>
                        Please contact the <strong>Compliance</strong> or <strong>Operations</strong> team.<br>
                        {img(ICON_TEL_B64, 16)}&nbsp;<code>compliance@bank.internal</code>
                        &nbsp;&nbsp;
                        {img(ICON_TEL_B64, 16)}&nbsp;<code>operations@bank.internal</code>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                # Main answer
                if message.get("content") and message["content"] != "__welcome__":
                    st.markdown(message["content"])

                # Email draft expander (EscalationAgent)
                if message.get("email_draft"):
                    with st.expander("📧 View Draft Escalation Email"):
                        st.markdown(message["email_draft"])
            else:
                st.markdown(message["content"])


# ── Handle sidebar sample question click ─────────────────────────────────────
if "pending_question" in st.session_state:
    pending = st.session_state.pop("pending_question")
    st.session_state.messages.append({"role": "user", "content": pending})
    result = orchestrator.handle(pending)
    st.session_state.messages.append({
        "role":             "assistant",
        "content":          result["answer"],
        "policy_name":      result.get("policy_name"),
        "policy_names":     result.get("policy_names", []),
        "found":            result["found"],
        "intent":           result.get("intent"),
        "agent_used":       result.get("agent_used"),
        "escalated":        result.get("escalated", False),
        "email_draft":      result.get("email_draft"),
        "compliance_alert": result.get("compliance_alert"),
    })
    st.rerun()


# ── Chat input ────────────────────────────────────────────────────────────────
user_input = st.chat_input("Ask a policy question, e.g. 'What are the KYC requirements?'")

if user_input:
    st.session_state.messages.append({"role": "user", "content": user_input})

    with st.chat_message("user", avatar=USER_AVATAR_URI):
        st.markdown(user_input)

    result = orchestrator.handle(user_input)

    with st.chat_message("assistant", avatar=BOT_AVATAR_URI):
        # Compliance alert
        if result.get("compliance_alert"):
            st.warning(result["compliance_alert"])

        # Source badge(s)
        names = result.get("policy_names") or (
            [result["policy_name"]] if result.get("policy_name") else []
        )
        if names:
            badges = "&nbsp;·&nbsp;".join(
                f'{policy_icon(n)}&nbsp;<strong>{n}</strong>' for n in names
            )
            st.markdown(
                f'<div class="policy-badge">{img(ICON_DOC_B64, 18)}&nbsp;Source:&nbsp;{badges}</div>',
                unsafe_allow_html=True,
            )

        # Agent used chip
        st.markdown(
            f'<small style="color:#6699aa;">🤖 Agent: <code>{result.get("agent_used","—")}</code>'
            f'&nbsp;|&nbsp;Intent: <code>{result.get("intent","—")}</code></small>',
            unsafe_allow_html=True,
        )

        # Fallback warning
        if not result["found"]:
            st.markdown(f"""
            <div class="fallback-box">
              {img(ICON_WARN_B64, 28)}
              <div class="fb-text">
                <strong>No exact policy match found.</strong><br>
                Please contact the <strong>Compliance</strong> or <strong>Operations</strong> team.<br>
                {img(ICON_TEL_B64, 16)}&nbsp;<code>compliance@bank.internal</code>
                &nbsp;&nbsp;
                {img(ICON_TEL_B64, 16)}&nbsp;<code>operations@bank.internal</code>
              </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown(result["answer"])

        # Email draft from EscalationAgent
        if result.get("email_draft"):
            with st.expander("📧 View Draft Escalation Email"):
                st.markdown(result["email_draft"])

    st.session_state.messages.append({
        "role":             "assistant",
        "content":          result["answer"],
        "policy_name":      result.get("policy_name"),
        "policy_names":     result.get("policy_names", []),
        "found":            result["found"],
        "intent":           result.get("intent"),
        "agent_used":       result.get("agent_used"),
        "escalated":        result.get("escalated", False),
        "email_draft":      result.get("email_draft"),
        "compliance_alert": result.get("compliance_alert"),
    })
