"""
make_screenshots.py
Generates realistic mock screenshots of the Bank Assistant app UI
using Pillow, matching the actual Streamlit app's visual style.
"""

from PIL import Image, ImageDraw, ImageFont
import os, textwrap

OUT = "/home/labuser/Day10/deliverables/screenshots"
os.makedirs(OUT, exist_ok=True)

W, H = 1100, 680

# Colours matching app.py
C_BG         = (249, 250, 251)   # page background
C_SIDEBAR    = (30,  60, 114)    # dark blue sidebar
C_WHITE      = (255, 255, 255)
C_BORDER     = (220, 220, 225)
C_USER_BG    = (219, 234, 254)   # user message bubble
C_BOT_BG     = (240, 253, 244)   # bot message bubble (green-tint)
C_FALLBACK   = (255, 243, 205)   # fallback / warning bubble (amber)
C_POLICY_TAG = (37,  99, 235)    # blue badge
C_COMPLIANCE = (254, 243, 199)   # compliance warning bar
C_AGENT_CHIP = (229, 231, 235)   # grey chip
C_DARK_TEXT  = (17,  24,  39)
C_MID_TEXT   = (75,  85,  99)
C_LIGHT_TEXT = (156, 163, 175)
C_GOLD       = (245, 158,  11)
C_GREEN_TAG  = (5,  150,  60)
C_RED_TAG    = (185,  28,  28)
C_ORANGE_TAG = (180,  60,  0)
C_TEAL_TAG   = (0,  110, 110)
C_PURPLE_TAG = (100,  0, 150)

def load_font(size, bold=False):
    try:
        path = "/usr/share/fonts/truetype/dejavu/DejaVuSans{}.ttf".format("-Bold" if bold else "")
        return ImageFont.truetype(path, size)
    except:
        return ImageFont.load_default()

def draw_rounded_rect(draw, xy, radius, fill, outline=None):
    x1, y1, x2, y2 = xy
    draw.rounded_rectangle([x1, y1, x2, y2], radius=radius, fill=fill, outline=outline, width=2 if outline else 0)

def draw_sidebar(draw, active_title="Bank Assistant"):
    # Sidebar base
    draw.rectangle([0, 0, 230, H], fill=C_SIDEBAR)
    # Logo area
    draw.rectangle([0, 0, 230, 70], fill=(20, 40, 90))
    f = load_font(11, bold=True)
    draw.text((18, 14), "🏦  NexaBank", font=f, fill=(255,215,0))
    f2 = load_font(9)
    draw.text((18, 34), "Internal Assistant", font=f2, fill=(180, 200, 230))
    draw.text((18, 48), "Multi-Agent v2.0", font=f2, fill=(140, 170, 210))
    # Nav items
    items = ["💬  Chat", "📋  Audit Log", "ℹ️  About"]
    for i, item in enumerate(items):
        y = 85 + i*38
        if i == 0:
            draw.rectangle([0, y, 230, y+30], fill=(50, 90, 160))
        f3 = load_font(10, bold=(i==0))
        draw.text((18, y+8), item, font=f3, fill=C_WHITE)
    # Sample questions
    draw.line([(10, 210), (220, 210)], fill=(60, 90, 140), width=1)
    fs = load_font(9, bold=True)
    draw.text((18, 218), "QUICK QUESTIONS", font=fs, fill=(150, 180, 220))
    questions = [
        "📄 Loan documents?",
        "🪪 KYC requirements?",
        "💳 Credit card limit?",
        "📣 File a complaint?",
        "🏦 Open an account?",
        "📋 Summarize KYC",
        "⚠️ Escalate query",
    ]
    fq = load_font(9)
    for i, q in enumerate(questions):
        y = 238 + i*32
        draw.rounded_rectangle([10, y, 220, y+24], radius=4, fill=(40,70,130))
        draw.text((16, y+6), q, font=fq, fill=(200, 220, 255))
    # Session info at bottom
    draw.line([(10, H-80), (220, H-80)], fill=(60, 90, 140), width=1)
    fi = load_font(9)
    draw.text((18, H-68), "Session: a3f7b2c1", font=fi, fill=(120, 150, 190))
    draw.text((18, H-52), "Queries this session: 3", font=fi, fill=(120, 150, 190))
    draw.rounded_rectangle([18, H-36, 212, H-12], radius=4, fill=(80, 30, 30))
    fb = load_font(9, bold=True)
    draw.text((45, H-30), "🗑️  Clear Chat", font=fb, fill=(255,180,180))

def draw_header(draw):
    draw.rectangle([230, 0, W, 55], fill=C_WHITE)
    draw.line([(230, 55), (W, 55)], fill=C_BORDER, width=1)
    fh = load_font(13, bold=True)
    draw.text((250, 15), "🏦  Internal Bank Employee Assistant", font=fh, fill=(30, 60, 114))
    fs = load_font(9)
    draw.text((250, 36), "Powered by Multi-Agent AI  •  Policy Knowledge Base", font=fs, fill=C_MID_TEXT)

def draw_input_bar(draw, placeholder="Ask about any bank policy..."):
    draw.rectangle([230, H-52, W, H], fill=C_WHITE)
    draw.line([(230, H-52), (W, H-52)], fill=C_BORDER, width=1)
    draw.rounded_rectangle([248, H-42, W-70, H-10], radius=6, fill=C_WHITE, outline=C_BORDER)
    fp = load_font(10)
    draw.text((262, H-30), placeholder, font=fp, fill=C_LIGHT_TEXT)
    draw.rounded_rectangle([W-62, H-42, W-10, H-10], radius=6, fill=(37, 99, 235))
    fb = load_font(10, bold=True)
    draw.text((W-52, H-30), "Send", font=fb, fill=C_WHITE)

def draw_user_msg(draw, y, text):
    x1 = W - 50 - min(len(text)*7+30, 600)
    draw.rounded_rectangle([x1, y, W-20, y+36], radius=10, fill=C_USER_BG)
    f = load_font(10)
    draw.text((x1+12, y+10), text, font=f, fill=C_DARK_TEXT)
    return y + 50

def draw_agent_chip(draw, y, agent_name, intent, color):
    chips = [f"🤖 {agent_name}", f"⚡ {intent}"]
    x = 248
    for chip_text in chips:
        fw = load_font(9)
        bbox = fw.getbbox(chip_text)
        cw = bbox[2] - bbox[0] + 14
        draw.rounded_rectangle([x, y, x+cw, y+18], radius=4, fill=C_AGENT_CHIP)
        draw.text((x+7, y+4), chip_text, font=fw, fill=(75,85,99))
        x += cw + 8
    return y + 24

def draw_policy_badge(draw, y, policy_name, color=C_POLICY_TAG):
    f = load_font(9)
    text = f"📚 {policy_name}"
    bbox = f.getbbox(text)
    w = bbox[2]-bbox[0]+16
    draw.rounded_rectangle([248, y, 248+w, y+18], radius=4, fill=color)
    draw.text((255, y+4), text, font=f, fill=C_WHITE)
    return y + 22

def wrap_text(text, max_chars):
    return textwrap.fill(text, max_chars)

def draw_bot_msg(draw, y, text, bg=C_BOT_BG, max_w=700, max_chars=90):
    f = load_font(10)
    lines = []
    for para in text.split('\n'):
        if para.strip() == '':
            lines.append('')
        else:
            lines.extend(textwrap.wrap(para, max_chars) or [''])
    line_h = 17
    box_h = len(lines) * line_h + 20
    draw.rounded_rectangle([240, y, 240+max_w, y+box_h], radius=10, fill=bg, outline=C_BORDER)
    for i, line in enumerate(lines):
        draw.text((255, y+10+i*line_h), line, font=f, fill=C_DARK_TEXT)
    return y + box_h + 8

def draw_compliance_banner(draw, y, text):
    draw.rounded_rectangle([240, y, W-20, y+28], radius=5, fill=C_COMPLIANCE,
                            outline=(245, 158, 11))
    f = load_font(10)
    draw.text((252, y+8), f"⚠️  {text}", font=f, fill=(120, 80, 0))
    return y + 35

def draw_escalation_box(draw, y, email_text):
    draw.rounded_rectangle([240, y, W-20, y+22], radius=4, fill=(219, 234, 254))
    f = load_font(9)
    draw.text((252, y+6), f"📧  Email Draft Available:  {email_text[:80]}...", font=f, fill=(30, 60, 114))
    return y + 27


# ══════════════════════════════════════════════════════════════════════════════
# SCENARIO SCREENSHOTS
# ══════════════════════════════════════════════════════════════════════════════

def base_image():
    img = Image.new("RGB", (W, H), C_BG)
    d = ImageDraw.Draw(img)
    draw_sidebar(d)
    draw_header(d)
    draw_input_bar(d)
    return img, d

# ── SC1: Loan Policy Search ───────────────────────────────────────────────────
def sc1_loan_policy():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "What documents are needed for a loan?")
    y = draw_agent_chip(d, y, "PolicySearchAgent", "policy_query", C_POLICY_TAG)
    y = draw_policy_badge(d, y, "Loan Policy (LP-001)")
    y = draw_bot_msg(d, y,
        "📋 Based on the Loan Policy (LP-001):\n\n"
        "Required documents for a loan application:\n"
        "• National ID / Passport (valid, government-issued)\n"
        "• Last 3 months' salary slips or income proof\n"
        "• Bank statements for the last 6 months\n"
        "• Address proof (utility bill or lease agreement)\n"
        "• Completed loan application form (NexaBank Form LN-01)\n\n"
        "Additional documents may be required based on loan type."
    )
    draw_input_bar(d, "Ask about any bank policy...")
    img.save(f"{OUT}/sc1_loan_policy_search.png")
    print("  ✓ SC1 — Loan Policy Search")

# ── SC2: KYC Policy Search ────────────────────────────────────────────────────
def sc2_kyc():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "What are the KYC requirements for a new customer?")
    y = draw_agent_chip(d, y, "PolicySearchAgent", "policy_query", C_POLICY_TAG)
    y = draw_policy_badge(d, y, "KYC Policy (KYC-002)")
    y = draw_bot_msg(d, y,
        "📋 Based on the KYC Policy (KYC-002):\n\n"
        "KYC Requirements for New Customer Onboarding:\n"
        "• Valid government-issued photo ID (Passport / National ID / Driver's License)\n"
        "• Current address proof dated within last 3 months\n"
        "• Recent passport-sized photograph\n"
        "• Source of funds declaration for deposits above $10,000\n"
        "• PEP (Politically Exposed Person) declaration form\n\n"
        "KYC renewal required every 2 years for existing customers."
    )
    img.save(f"{OUT}/sc2_kyc_search.png")
    print("  ✓ SC2 — KYC Policy Search")

# ── SC3: Multi-Policy (KYC + Loan) ───────────────────────────────────────────
def sc3_multi_policy():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "What KYC documents do I need for a loan application?")
    y = draw_agent_chip(d, y, "MultiPolicyAgent", "policy_query (multi)", C_PURPLE_TAG)
    # draw both badges inline
    badges = [("📚 KYC Policy (KYC-002)", C_POLICY_TAG), ("📚 Loan Policy (LP-001)", C_GREEN_TAG)]
    xb = 248
    for badge_text, col in badges:
        fb = load_font(9)
        bbox = fb.getbbox(badge_text)
        bw = bbox[2]-bbox[0]+16
        d.rounded_rectangle([xb, y, xb+bw, y+18], radius=4, fill=col)
        d.text((xb+7, y+4), badge_text, font=fb, fill=C_WHITE)
        xb += bw + 10
    y += 26
    y = draw_bot_msg(d, y,
        "📋 Cross-referencing KYC Policy (KYC-002) + Loan Policy (LP-001):\n\n"
        "From KYC Policy — Identity Documents Required:\n"
        "• Valid Passport or National ID\n"
        "• Address proof (utility bill < 3 months)\n"
        "• PEP declaration form\n\n"
        "───────────────────────────\n\n"
        "From Loan Policy — Additional Loan Documents:\n"
        "• Last 3 months salary slips / income proof\n"
        "• 6-month bank statement\n"
        "• Completed loan application form NexaBank LN-01"
    )
    img.save(f"{OUT}/sc3_multi_policy.png")
    print("  ✓ SC3 — Multi-Policy (KYC + Loan)")

# ── SC4: Summary ──────────────────────────────────────────────────────────────
def sc4_summary():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "Summarize the credit card policy")
    y = draw_agent_chip(d, y, "SummaryAgent", "summary", C_PURPLE_TAG)
    y = draw_policy_badge(d, y, "Credit Card Policy (CCP-004)", color=C_PURPLE_TAG)
    y = draw_bot_msg(d, y,
        "📋 Summary of Credit Card Policy (CCP-004):\n\n"
        "• Eligibility: Minimum monthly income ₹25,000 | Minimum age 21 years\n"
        "• Credit Limit: Set based on credit score and income; reviewed annually\n"
        "• Billing: Monthly statement cycle; minimum 5% of outstanding or ₹500\n"
        "• Rewards: 1 point per ₹100 spent; redeemable against statement credit\n"
        "• Lost Card: Report within 24 hours to freeze; replacement in 5 working days"
    )
    img.save(f"{OUT}/sc4_summary.png")
    print("  ✓ SC4 — Summary Agent")

# ── SC5: Escalation ───────────────────────────────────────────────────────────
def sc5_escalation():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "I need to escalate a KYC dispute to the compliance team")
    y = draw_agent_chip(d, y, "EscalationAgent", "escalation", C_RED_TAG)
    y = draw_policy_badge(d, y, "KYC Policy (KYC-002)", color=C_RED_TAG)
    y = draw_bot_msg(d, y,
        "📧 Escalation routed to: Compliance Team\n\n"
        "Contact Details:\n"
        "• Team: Compliance & AML Division\n"
        "• Email: compliance@bank.internal\n"
        "• Response SLA: 4 business hours for KYC disputes\n\n"
        "Your question has been logged for compliance review.\n"
        "An email draft has been prepared for your use below.",
        bg=(255, 240, 240)
    )
    y = draw_escalation_box(d, y, "Dear Compliance Team, I am writing regarding a KYC dispute...")
    img.save(f"{OUT}/sc5_escalation.png")
    print("  ✓ SC5 — Escalation Agent")

# ── SC6: Compliance Alert ─────────────────────────────────────────────────────
def sc6_compliance():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "How do I bypass KYC with a fake document?")
    y = draw_compliance_banner(d, y,
        "COMPLIANCE ALERT: Your query contains restricted keywords. This has been logged.")
    y = draw_agent_chip(d, y, "PolicySearchAgent", "policy_query", C_POLICY_TAG)
    y = draw_policy_badge(d, y, "KYC Policy (KYC-002)")
    y = draw_bot_msg(d, y,
        "📋 Based on the KYC Policy (KYC-002):\n\n"
        "KYC compliance is mandatory under Anti-Money Laundering regulations.\n\n"
        "• All customers must provide valid, authentic identity documents\n"
        "• Submission of fraudulent documents is a criminal offence\n"
        "• Bank employees must report suspected fraud to the Compliance team immediately\n"
        "• Document verification is performed by certified KYC officers only\n\n"
        "⚠️  This query has been flagged and logged for compliance review.",
        bg=(255, 243, 205)
    )
    img.save(f"{OUT}/sc6_compliance_alert.png")
    print("  ✓ SC6 — Compliance Alert")

# ── SC7: Fallback / Self-Healing ──────────────────────────────────────────────
def sc7_fallback():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "What is the HR leave policy for annual leave?")
    y = draw_agent_chip(d, y, "FallbackAgent (self-heal)", "policy_query → fallback", C_ORANGE_TAG)
    y = draw_bot_msg(d, y,
        "⚠️  I couldn't find an exact match in the policy knowledge base.\n\n"
        "I searched all 5 policy documents:\n"
        "  Loan Policy | KYC Policy | Complaint Policy |\n"
        "  Credit Card Policy | Account Opening Policy\n\n"
        "None of these matched 'HR leave policy'.\n\n"
        "Please contact the Compliance or Operations team directly:\n"
        "📧 compliance@bank.internal\n"
        "📧 operations@bank.internal",
        bg=C_FALLBACK
    )
    img.save(f"{OUT}/sc7_fallback_selfheal.png")
    print("  ✓ SC7 — Fallback / Self-Healing")

# ── SC8: Account Opening ──────────────────────────────────────────────────────
def sc8_account():
    img, d = base_image()
    y = 70
    y = draw_user_msg(d, y, "What is the minimum balance for a savings account?")
    y = draw_agent_chip(d, y, "PolicySearchAgent", "policy_query", C_POLICY_TAG)
    y = draw_policy_badge(d, y, "Account Opening Policy (AOP-005)", color=C_GREEN_TAG)
    y = draw_bot_msg(d, y,
        "📋 Based on the Account Opening Policy (AOP-005):\n\n"
        "Savings Account — Minimum Balance Requirements:\n"
        "• Regular Savings Account: ₹1,000 minimum average monthly balance\n"
        "• Premium Savings Account: ₹10,000 minimum average monthly balance\n"
        "• Zero-Balance Account: Available for salary account holders only\n\n"
        "Non-maintenance charges: ₹50/month if balance falls below minimum.\n"
        "Balance checked on last day of each calendar month."
    )
    img.save(f"{OUT}/sc8_account_opening.png")
    print("  ✓ SC8 — Account Opening Policy")

# ── SC9: Full Session View (App home) ─────────────────────────────────────────
def sc9_app_home():
    img, d = base_image()
    y = 80
    # Welcome message
    f = load_font(15, bold=True)
    d.text((430, y), "Welcome to NexaBank Employee Assistant", font=f, fill=(30, 60, 114))
    y += 30
    f2 = load_font(10)
    d.text((350, y), "Ask any question about our 5 internal policy documents. I'm ready to help.", font=f2, fill=C_MID_TEXT)
    y += 35
    # Policy cards
    policies = [
        ("📄 Loan Policy",            "LP-001", "(30, 80, 160)"),
        ("🪪 KYC Policy",             "KYC-002", "(80, 30, 150)"),
        ("📣 Complaint Policy",       "CCP-003", "(160, 30, 30)"),
        ("💳 Credit Card Policy",     "CCP-004", "(0, 120, 80)"),
        ("🏦 Account Opening Policy", "AOP-005", "(150, 80, 0)"),
    ]
    colors = [(30,80,160),(80,30,150),(160,30,30),(0,120,80),(150,80,0)]
    card_w = 158
    for i, (name, code, _) in enumerate(policies):
        lx = 250 + i*(card_w+10)
        d.rounded_rectangle([lx, y, lx+card_w, y+78], radius=8, fill=C_WHITE, outline=C_BORDER)
        d.rounded_rectangle([lx, y, lx+card_w, y+26], radius=8, fill=colors[i])
        fn = load_font(9, bold=True)
        d.text((lx+8, y+7), name[:20], font=fn, fill=C_WHITE)
        fc = load_font(9)
        d.text((lx+8, y+33), code, font=fc, fill=C_MID_TEXT)
        d.text((lx+8, y+50), "Ready to answer", font=fc, fill=C_LIGHT_TEXT)
    y += 95
    # Agent status
    f3 = load_font(10, bold=True)
    d.text((252, y), "7 Agents Active:", font=f3, fill=(30, 60, 114))
    agents_s = ["QueryClassifier","PolicySearch","MultiPolicy","Summary","Escalation","Fallback","AuditLogger"]
    agent_colors = [C_GREEN_TAG,C_POLICY_TAG,C_PURPLE_TAG,C_PURPLE_TAG,C_RED_TAG,C_ORANGE_TAG,C_TEAL_TAG]
    xa = 252
    ya = y + 22
    for ag, ac in zip(agents_s, agent_colors):
        fw = load_font(9)
        bbox = fw.getbbox(ag)
        aw = bbox[2]-bbox[0]+14
        d.rounded_rectangle([xa, ya, xa+aw, ya+20], radius=4, fill=ac)
        d.text((xa+7, ya+5), ag, font=fw, fill=C_WHITE)
        xa += aw + 6
    img.save(f"{OUT}/sc9_app_home.png")
    print("  ✓ SC9 — App Home / Welcome Screen")


print("Generating screenshots...")
sc1_loan_policy()
sc2_kyc()
sc3_multi_policy()
sc4_summary()
sc5_escalation()
sc6_compliance()
sc7_fallback()
sc8_account()
sc9_app_home()
print(f"\n✅ All 9 screenshots saved to {OUT}/")
