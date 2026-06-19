# 🏦 Internal Bank Employee Assistant

An AI-powered chatbot that helps bank employees instantly search internal
policy information — no internet required, no external APIs.

---

## Features

- Answers questions from **5 official bank policy documents**
- **Keyword-based search** with smart scoring (most relevant policy wins)
- **Focused snippet extraction** — returns the relevant section, not the full doc
- Fallback message directing to compliance/operations if no answer found
- Clean **Streamlit chat UI** with sidebar sample questions
- Policy source badge shown with every answer

---

## Policies Covered

| Policy | Code | Topics |
|---|---|---|
| Loan Policy | LP-001 | Eligibility, loan types, documents, approval |
| KYC Policy | KYC-002 | Identity verification, document requirements |
| Customer Complaint Policy | CCP-003 | Channels, timelines, escalation |
| Credit Card Policy | CCP-004 | Card types, billing, disputes, cancellation |
| Account Opening Policy | AOP-005 | Account types, documents, process |

---

## Project Structure

```
bank_assistant/
├── app.py                          ← Streamlit UI (run this)
├── requirements.txt                ← Dependencies
├── README.md                       ← This file
│
├── policies/                       ← Knowledge base (plain text)
│   ├── loan_policy.txt
│   ├── kyc_policy.txt
│   ├── customer_complaint_policy.txt
│   ├── credit_card_policy.txt
│   └── account_opening_policy.txt
│
└── utils/
    ├── __init__.py
    ├── knowledge_base.py           ← Loads policies + search logic
    └── agent.py                    ← Core agent: matches question → answer
```

---

## Setup & Run

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Start the app

```bash
streamlit run app.py
```

### 3. Open in browser

Streamlit will print a local URL, e.g.:
```
Local URL: http://localhost:8501
```

---

## How It Works

```
Employee types a question
        ↓
agent.py receives the question
        ↓
knowledge_base.py scores each policy by keyword matches
        ↓
Best matching policy is selected
        ↓
extract_answer() finds the most relevant paragraph
        ↓
Answer displayed in chat with source badge
        ↓
If no match → fallback message shown
```

---

## Adding New Policies

1. Create a new `.txt` file in the `policies/` folder
2. Add a keyword list for it in `KEYWORD_MAP` inside `utils/knowledge_base.py`
3. Add its display name in `POLICY_DISPLAY_NAMES` inside `utils/agent.py`
4. Restart the app — it loads automatically

---

## Sample Questions to Try

- "What documents are needed for a loan?"
- "What is the minimum credit score for a credit card?"
- "How long does KYC verification take?"
- "How do I file a customer complaint?"
- "What is the minimum balance for a savings account?"
- "What is the interest rate for a home loan?"
- "What happens if a credit card is lost or stolen?"
- "How do I open a business account?"

---

## Fallback Behavior

If the question doesn't match any policy:
> "I'm sorry, I could not find a relevant policy for your question.
> Please contact the **compliance or operations team** for further assistance.
> 📧 compliance@bank.internal | 📧 operations@bank.internal"
