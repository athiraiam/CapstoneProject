# Architecture Diagram — Internal Bank Employee Assistant

## System Architecture (Mermaid)

```mermaid
flowchart TD
    subgraph UI["🖥️ Presentation Layer"]
        A["👤 Bank Employee"]
        B["Streamlit Chat UI\n(app.py)"]
    end

    subgraph AGENT["🤖 Agent Layer"]
        C["Agent Core\n(agent.py)"]
        D["Knowledge Base Loader\n(knowledge_base.py)"]
    end

    subgraph KB["📚 Knowledge Base (Local)"]
        E1["loan_policy.txt\nLP-001"]
        E2["kyc_policy.txt\nKYC-002"]
        E3["customer_complaint_policy.txt\nCCP-003"]
        E4["credit_card_policy.txt\nCCP-004"]
        E5["account_opening_policy.txt\nAOP-005"]
    end

    subgraph SEARCH["🔍 Search Engine"]
        F["Keyword Scorer\n(KEYWORD_MAP)"]
        G["Snippet Extractor\n(extract_answer)"]
    end

    subgraph OUTPUT["📤 Output Layer"]
        H{"Match\nFound?"}
        I["✅ Policy Answer\n+ Source Badge"]
        J["⚠️ Fallback Message\n'Contact Compliance Team'"]
    end

    A -->|"Types question"| B
    B -->|"Sends question"| C
    C -->|"Loads on startup"| D
    D -->|"Reads files"| E1
    D -->|"Reads files"| E2
    D -->|"Reads files"| E3
    D -->|"Reads files"| E4
    D -->|"Reads files"| E5
    C -->|"Searches"| F
    F -->|"Best match policy text"| G
    G -->|"Relevant snippet"| H
    H -->|"Yes"| I
    H -->|"No"| J
    I -->|"Displays in chat"| B
    J -->|"Displays in chat"| B
```

## Component Description

| Component | File | Responsibility |
|---|---|---|
| Streamlit Chat UI | app.py | Receives employee input, renders chat messages, displays source badge |
| Agent Core | utils/agent.py | Orchestrates search and returns structured answer |
| Knowledge Base Loader | utils/knowledge_base.py | Reads .txt policy files into memory at startup |
| Keyword Scorer | utils/knowledge_base.py | Scores each policy by keyword overlap with question |
| Snippet Extractor | utils/knowledge_base.py | Returns focused paragraph instead of full policy |
| Policy Files | policies/*.txt | Plain-text source of truth for all bank policies |

## Data Flow

```
Employee Question
      │
      ▼
Keyword Scoring (all 5 policies scored simultaneously)
      │
      ▼
Best Policy Selected (highest keyword match score)
      │
      ▼
Snippet Extraction (finds most relevant paragraph within policy)
      │
      ▼
Answer Returned to UI with Source Policy Name
```
