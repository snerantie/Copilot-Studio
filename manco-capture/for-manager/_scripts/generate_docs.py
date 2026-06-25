"""Generate Word (.docx) and PowerPoint (.pptx) deliverables from the
manager-facing markdown sources.

Outputs (next to the script's parent folder):
    manco-capture/for-manager/business-case-upgrade.docx
    manco-capture/for-manager/status-and-blockers.docx
    manco-capture/for-manager/manco-capture-business-case.pptx

Re-run after editing pricing, status, or copy in this script. The .md
files in for-manager/ remain the canonical content for GitHub viewing;
the .docx / .pptx are derived deliverables for email/sharing.

Usage:
    python generate_docs.py
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Cm, Inches
from pptx import Presentation
from pptx.dml.color import RGBColor as PPTXColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches as PPTXInches, Pt as PPTXPt, Emu


OUT_DIR = Path(__file__).resolve().parent.parent

# ---------------------------------------------------------------------------
# Shared content
# ---------------------------------------------------------------------------

PROJECT_NAME = "Manco Capture Agent"
SUBTITLE = "AI agent for the CIO and senior management team — Sprint 1 of three"
STATUS_DATE = "25 June 2026"
AUTHOR_PLACEHOLDER = "[your name]"
MANAGER_PLACEHOLDER = "[your manager]"

REPO_URL = "https://github.com/snerantie/Copilot-Studio"
BRANCH_URL = (
    "https://github.com/snerantie/Copilot-Studio/tree/feat/manco-capture-sprint1"
)
PR_URL = "https://github.com/snerantie/Copilot-Studio/pull/1"

ACCENT = RGBColor(0x1F, 0x4E, 0x79)        # deep blue
MUTED = RGBColor(0x59, 0x59, 0x59)         # mid grey
SUCCESS = RGBColor(0x2E, 0x7D, 0x32)       # green
WARNING = RGBColor(0xE6, 0x5A, 0x00)       # amber
DANGER = RGBColor(0xC6, 0x28, 0x28)        # red

PPTX_ACCENT = PPTXColor(0x1F, 0x4E, 0x79)
PPTX_MUTED = PPTXColor(0x59, 0x59, 0x59)
PPTX_SUCCESS = PPTXColor(0x2E, 0x7D, 0x32)
PPTX_WARNING = PPTXColor(0xE6, 0x5A, 0x00)
PPTX_DANGER = PPTXColor(0xC6, 0x28, 0x28)
PPTX_LIGHT_BG = PPTXColor(0xF2, 0xF5, 0xF9)
PPTX_WHITE = PPTXColor(0xFF, 0xFF, 0xFF)


# ---------------------------------------------------------------------------
# Word helpers
# ---------------------------------------------------------------------------


def _set_cell_shading(cell, hex_color: str) -> None:
    """Set table cell background colour using OOXML shading element."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _style_heading(paragraph, size_pt: int, color: RGBColor, bold: bool = True) -> None:
    for run in paragraph.runs:
        run.font.size = Pt(size_pt)
        run.font.color.rgb = color
        run.font.bold = bold
        run.font.name = "Calibri"


def _add_paragraph(doc, text: str, size: int = 11, bold: bool = False,
                   color: RGBColor | None = None, italic: bool = False) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = "Calibri"
    if color is not None:
        run.font.color.rgb = color


def _add_bullet(doc, text: str, level: int = 0) -> None:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.6 + 0.6 * level)
    run = p.runs[0] if p.runs else p.add_run()
    run.text = text
    run.font.size = Pt(11)
    run.font.name = "Calibri"


def _add_h1(doc, text: str) -> None:
    h = doc.add_heading(text, level=1)
    _style_heading(h, 18, ACCENT)


def _add_h2(doc, text: str) -> None:
    h = doc.add_heading(text, level=2)
    _style_heading(h, 14, ACCENT)


def _add_docx_table(doc, headers: list[str], rows: list[list[str]],
                    header_fill: str = "1F4E79", first_col_bold: bool = False) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Light Grid Accent 1"

    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        _set_cell_shading(cell, header_fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Body rows
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(val)
            run.font.size = Pt(10)
            run.font.name = "Calibri"
            if first_col_bold and c_idx == 0:
                run.font.bold = True
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _add_title_block(doc, title: str, subtitle: str, meta_lines: list[str]) -> None:
    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(title)
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = ACCENT
    run.font.name = "Calibri"

    # Subtitle
    p = doc.add_paragraph()
    run = p.add_run(subtitle)
    run.font.size = Pt(13)
    run.font.italic = True
    run.font.color.rgb = MUTED
    run.font.name = "Calibri"

    # Meta
    for line in meta_lines:
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(10)
        run.font.color.rgb = MUTED
        run.font.name = "Calibri"

    # Spacer
    doc.add_paragraph()


def _set_document_margins(doc) -> None:
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)


# ---------------------------------------------------------------------------
# Business case Word document
# ---------------------------------------------------------------------------


def build_business_case_docx(path: Path) -> None:
    doc = Document()
    _set_document_margins(doc)

    _add_title_block(
        doc,
        title="Business case — Copilot Studio capability upgrade",
        subtitle=PROJECT_NAME + " — Sprint 1 of three",
        meta_lines=[
            f"Author: {AUTHOR_PLACEHOLDER}    Date: {STATUS_DATE}",
            f"For: {MANAGER_PLACEHOLDER} (review and decision)",
            "Initiative: Manco Capture Agent (Capture · Reporter · Insights)",
        ],
    )

    # TL;DR
    _add_h1(doc, "TL;DR")
    _add_paragraph(
        doc,
        "We are building an AI agent for the CIO and senior management team that "
        "automates capture, reporting, and analysis of work on the Manco Board "
        "(Jira VFST2). Sprint 1 (capture from email to Jira Task with one-tap "
        "approval in Teams) is roughly 60% built and on track.",
    )
    _add_paragraph(
        doc,
        "We have hit a Microsoft Copilot Studio licence ceiling that prevents us "
        "from using three features needed for a clean architecture and for the "
        "next two sprints. The features are: Tools, Knowledge, and Agents — each "
        "currently shows \"needs upgrade\" in our tenant.",
    )
    _add_paragraph(doc, "Requested decision:", bold=True)
    _add_bullet(doc, "Approve the Copilot Studio capability upgrade.")
    _add_bullet(
        doc,
        "Approve a parallel admin policy review to enable AI Builder Prompts "
        "under our existing governance.",
    )
    _add_paragraph(doc, "Indicative cost:", bold=True)
    _add_bullet(
        doc,
        "Approximately £125–£170 per user per month for full Copilot Studio "
        "licensing (Microsoft list pricing — confirm with M365 procurement).",
    )
    _add_paragraph(doc, "Indicative payback:", bold=True)
    _add_bullet(
        doc,
        "Under 3 months at a conservative 1 hour per senior leader per week of "
        "saved Jira hygiene and action-item triage.",
    )

    # What we are building
    _add_h1(doc, "1. What we are building")
    _add_paragraph(doc, "A three-pillar AI capability for senior management:")
    _add_docx_table(
        doc,
        headers=["Pillar", "What it does", "Status"],
        rows=[
            [
                "1. Capture",
                "Turns emails and Teams meeting transcripts into properly-filed "
                "Jira Tasks on the Manco Board, with human approval in Teams",
                "Sprint 1 — in build",
            ],
            [
                "2. Reporter",
                "Conversational Q&A and scheduled digests over Manco Board "
                "status (\"what's overdue?\", \"status for Cyber\")",
                "Sprint 2 — designed",
            ],
            [
                "3. Insights",
                "Cross-source analysis of Jira history + Manco SharePoint to "
                "surface risks, themes, and decisions needing escalation",
                "Sprint 3 — designed",
            ],
        ],
        first_col_bold=True,
    )
    _add_paragraph(doc, "")
    _add_paragraph(
        doc,
        "All three pillars target the same audience, use the same Microsoft + "
        "Atlassian stack we already pay for, and reuse the same foundations "
        "we are putting in place now.",
    )

    # Why it matters
    _add_h1(doc, "2. Why it matters")
    for b in [
        "Action items currently fall through the cracks. Items raised in "
        "emails, Teams chats, and leadership meetings are tracked manually "
        "if at all. Jira hygiene depends on whoever remembers.",
        "Leadership time is the scarcest resource. The CIO and senior team "
        "should not be triaging Jira; they should be reviewing and deciding.",
        "Reporting is a known pain point. Status updates are assembled by "
        "hand before every Manco meeting. The Reporter pillar removes that "
        "work.",
        "Insights are invisible today. Trends and risks across the Manco "
        "Board sit in Jira and SharePoint with no integration layer.",
    ]:
        _add_bullet(doc, b)

    # Progress
    _add_h1(doc, "3. Progress to date")
    for b in [
        "Design and build artefacts version-controlled in the Copilot-Studio "
        "GitHub repository, under manco-capture/.",
        "9-file build pack: architecture, provisioning checklist, SharePoint "
        "list schema, AI extraction prompt, Adaptive Card design, Jira API "
        "payload, two Power Automate flow guides, 5-scenario test plan.",
        "Sprint 1 provisioning complete: Teams approval channel, SharePoint "
        "audit list, Jira API token, Outlook trigger category.",
        "Extraction proof-of-concept tested in Copilot Studio.",
    ]:
        _add_bullet(doc, b)

    # Blocker
    _add_h1(doc, "4. The blocker")
    _add_paragraph(
        doc,
        "Our current Copilot Studio licence is at the basic / trial tier. "
        "Three features needed for a production-quality build are gated "
        "behind a paid upgrade and show \"needs upgrade\" in the interface:",
    )
    _add_docx_table(
        doc,
        headers=["Feature", "What it does", "Why we need it"],
        rows=[
            [
                "Tools",
                "Exposes a Prompt (LLM call) as a clean, callable object that "
                "Power Automate can invoke with structured inputs and outputs",
                "Textbook integration pattern. Without it, we work around "
                "through Topics — functional but less maintainable.",
            ],
            [
                "Knowledge",
                "Grounds the agent's responses in trusted content sources "
                "such as the Manco SharePoint document library",
                "Critical for Pillar 3 (Insights). Without it the agent "
                "cannot reason over existing Manco documentation.",
            ],
            [
                "Agents (multi-agent)",
                "Separate specialised agents (Capture / Reporter / Insights) "
                "with one orchestrator at the front",
                "Critical for Pillars 2 and 3. Without it everything must "
                "collapse into a single overloaded agent.",
            ],
        ],
        first_col_bold=True,
    )
    _add_paragraph(doc, "")
    _add_paragraph(
        doc,
        "Separately, AI Builder Prompts is disabled at the tenant admin level "
        "(\"feature has been disabled — please contact your administrator\"). "
        "This is a different governance decision but in the same family. "
        "Reviewing the admin policy in parallel would give us a more flexible "
        "architecture.",
    )

    # What an upgrade unlocks
    _add_h1(doc, "5. What an upgrade unlocks")
    _add_h2(doc, "For Sprint 1 (Capture)")
    _add_docx_table(
        doc,
        headers=["Outcome", "Without upgrade", "With upgrade"],
        rows=[
            ["LLM extraction works", "Yes (via Topics workaround)", "Yes (via Tools — cleaner)"],
            ["Power Automate integration", "Bespoke Topics glue", "Native Tool invocation"],
            ["Time to build Sprint 1", "~2 working days", "~1 working day"],
        ],
        first_col_bold=True,
    )
    _add_paragraph(doc, "")
    _add_h2(doc, "For Sprint 2 (Reporter) and Sprint 3 (Insights)")
    _add_docx_table(
        doc,
        headers=["Outcome", "Without upgrade", "With upgrade"],
        rows=[
            [
                "Conversational Reporter agent",
                "Single overloaded agent, harder to govern",
                "Dedicated agent with clear scope",
            ],
            [
                "Insights against SharePoint docs",
                "Blocked — no Knowledge feature",
                "Native RAG against Manco SharePoint",
            ],
            [
                "Future addition of new agents",
                "Each one inflates the single agent",
                "Add as separate agents with orchestrator",
            ],
            [
                "Vendor lock-in risk",
                "Higher — workarounds become debt",
                "Lower — using supported features",
            ],
        ],
        first_col_bold=True,
    )

    # Cost framing
    _add_h1(doc, "6. Cost framing")
    _add_paragraph(
        doc,
        "Microsoft list pricing (verify with M365 procurement; pricing changes):",
    )
    _add_docx_table(
        doc,
        headers=["Item", "Indicative cost"],
        rows=[
            ["Microsoft Copilot Studio (per-user)", "~£164 / user / month"],
            ["Microsoft Copilot Studio (pay-as-you-go message capacity)", "from ~£0.008 / message"],
            ["AI Builder add-on capacity (if Prompts later enabled)", "from ~£400 / tenant / month (volume-tied)"],
            ["M365 Copilot licence (separate, if not already held)", "~£25 / user / month"],
        ],
        first_col_bold=True,
    )
    _add_paragraph(doc, "")
    _add_paragraph(doc, "Realistic scoped cost for our use case:", bold=True)
    for b in [
        "2 builder/admin seats (you + your manager) on full Copilot Studio.",
        "Plus modest pay-as-you-go message capacity for runtime usage by the "
        "agent.",
        "Approximately £330–£500 / month all-in for Sprint 1 + Sprint 2 "
        "capacity, scaling modestly with usage.",
    ]:
        _add_bullet(doc, b)
    _add_paragraph(doc, "")
    _add_paragraph(doc, "Payback estimate:", bold=True)
    for b in [
        "Senior management team day rate is high. Even 1 hour per week per "
        "member saved (action-item triage, status assembly, manual Jira "
        "hygiene) pays for the licence many times over within the quarter.",
        "The Reporter pillar alone removes manual Manco pack assembly — "
        "typically several hours per cycle.",
    ]:
        _add_bullet(doc, b)

    # Options
    _add_h1(doc, "7. Options considered")
    _add_docx_table(
        doc,
        headers=["Option", "Pros", "Cons", "Recommended?"],
        rows=[
            [
                "A. Do nothing — stay on basic tier",
                "Zero cost",
                "Sprint 2 and 3 effectively blocked; Sprint 1 architecture brittle",
                "No",
            ],
            [
                "B. Upgrade Copilot Studio (this paper)",
                "Unlocks all three pillars; supported by Microsoft",
                "Licence cost",
                "Yes — recommended",
            ],
            [
                "C. Route AI through Azure OpenAI Service",
                "Maximum control; ties into Azure governance",
                "Org has no Azure OpenAI deployment today; sizeable setup",
                "No (today)",
            ],
            [
                "D. Defer the whole initiative",
                "Zero cost",
                "Action-item leakage continues; Reporter/Insights value not realised",
                "No",
            ],
        ],
        first_col_bold=True,
    )

    # Risks
    _add_h1(doc, "8. Risks of not upgrading")
    for n, b in enumerate(
        [
            "Architectural debt early. Sprint 1 will work but on a fragile "
            "foundation that we'll likely rebuild when we get to Sprint 2/3.",
            "Sprint 2 and 3 are effectively blocked. Without Knowledge we "
            "cannot reason over Manco SharePoint. Without Agents we cannot "
            "cleanly modularise.",
            "Slower delivery. Workarounds take longer to build and maintain. "
            "We trade licence cost for engineering time.",
            "Higher cost of change. Replacing workarounds with proper "
            "features later is more expensive than doing it right now.",
        ],
        start=1,
    ):
        _add_paragraph(doc, f"{n}.  {b}")

    # Recommendation
    _add_h1(doc, "9. Recommendation")
    _add_paragraph(doc, "1.  Approve the Microsoft Copilot Studio capability upgrade for the environment housing the Manco Capture Agent.")
    _add_paragraph(doc, "2.  Ask the Power Platform admin to review the AI Builder Prompts policy and consider enabling it under existing governance.")
    _add_paragraph(doc, "3.  Continue Sprint 1 in the meantime using the Topics-based workaround; migrate to Tools immediately once the upgrade lands. No code is wasted — only the LLM invocation step changes.")

    # Decisions
    _add_h1(doc, "10. Decision asks")
    _add_docx_table(
        doc,
        headers=["Decision", "Answer"],
        rows=[
            ["Approve upgrade", "yes / no / need more info"],
            ["Approve admin policy review on AI Builder Prompts", "yes / no"],
            ["Continue Sprint 1 on workaround in parallel", "yes / no"],
            ["Procurement contact for confirming exact pricing", "_________________"],
        ],
        first_col_bold=True,
    )

    # Appendix
    _add_h1(doc, "Appendix — Glossary")
    glossary = [
        ("Copilot Studio", "Microsoft's low-code platform for building AI agents on top of Microsoft 365."),
        ("Tool / Action", "A callable unit within a Copilot Studio agent — typically a Prompt, an API call, or a Power Automate flow."),
        ("Knowledge", "Trusted content source (SharePoint sites, documents, websites) the agent grounds answers in (RAG)."),
        ("Agent / multi-agent", "Separate specialised agents that can hand off to each other under one orchestrator."),
        ("Topic", "The classic Copilot Studio building block — a scripted conversational flow."),
        ("AI Builder Prompts", "An AI Builder feature that lets Power Automate flows invoke a GPT-style prompt directly."),
    ]
    for term, defn in glossary:
        p = doc.add_paragraph()
        run_term = p.add_run(term + " — ")
        run_term.font.bold = True
        run_term.font.size = Pt(11)
        run_term.font.color.rgb = ACCENT
        run_term.font.name = "Calibri"
        run_def = p.add_run(defn)
        run_def.font.size = Pt(11)
        run_def.font.name = "Calibri"

    _add_h1(doc, "Reference")
    _add_bullet(doc, f"GitHub repository: {REPO_URL}")
    _add_bullet(doc, f"Build pack branch: {BRANCH_URL}")
    _add_bullet(doc, f"Open pull request (build pack + this paper): {PR_URL}")

    doc.save(path)
    print(f"  wrote {path.name}")


# ---------------------------------------------------------------------------
# Status Word document
# ---------------------------------------------------------------------------


def build_status_docx(path: Path) -> None:
    doc = Document()
    _set_document_margins(doc)

    _add_title_block(
        doc,
        title="Status & blockers — Manco Capture Agent (Sprint 1)",
        subtitle=f"Status as of {STATUS_DATE}",
        meta_lines=[
            f"Project: {PROJECT_NAME} — emails to Jira Tasks on the Manco Board (VFST2), with human approval in Teams.",
            f"Build pack: {BRANCH_URL}/manco-capture",
        ],
    )

    # Headline
    _add_h1(doc, "1. Headline")
    for b in [
        "Sprint 1 is approximately 60% built.",
        "Foundations (Teams channel, SharePoint list, Jira token, Outlook trigger) are DONE.",
        "Two blockers prevent completion of the AI extraction step. Both are Microsoft licensing / policy decisions, not technical problems.",
    ]:
        _add_bullet(doc, b)

    # Progress
    _add_h1(doc, "2. Progress checklist")
    _add_h2(doc, "Phase 1 — Provisioning")
    _add_docx_table(
        doc,
        headers=["#", "Task", "Status"],
        rows=[
            ["1.1", "Capture – Approvals private Teams channel inside the Manco Team", "DONE"],
            ["1.2", "Capture Audit SharePoint list in the Manco SharePoint site", "DONE"],
            ["1.3", "Jira Cloud API token created and stored", "DONE"],
            ["1.4", "Outlook category \"Capture\" created on user mailbox", "DONE"],
            ["1.5", "AI Builder access — Models available, Prompts disabled", "PARTIAL"],
        ],
    )

    _add_paragraph(doc, "")
    _add_h2(doc, "Phase 2 — Power Platform environment variables")
    _add_docx_table(
        doc,
        headers=["#", "Task", "Status"],
        rows=[
            ["2.1", "manco_JiraBaseUrl (Text)", "Pending"],
            ["2.2", "manco_JiraAuthEmail (Text)", "Pending"],
            ["2.3", "manco_JiraApiToken (Secret)", "Pending"],
        ],
    )

    _add_paragraph(doc, "")
    _add_h2(doc, "Phase 3 — Copilot Studio extraction agent")
    _add_docx_table(
        doc,
        headers=["#", "Task", "Status"],
        rows=[
            ["3.1", "Create agent Manco Capture", "DONE"],
            ["3.2", "Define extraction prompt (text content)", "DONE"],
            ["3.3", "Wire as callable Tool with structured inputs/outputs", "BLOCKED — Tools needs upgrade"],
            ["3.4", "Fallback: wire as Topic with workaround", "In progress"],
        ],
    )

    _add_paragraph(doc, "")
    _add_h2(doc, "Phases 4–6")
    _add_docx_table(
        doc,
        headers=["Phase", "Task", "Status"],
        rows=[
            ["4", "Build Power Automate parent flow (Outlook → extract → dispatch)", "Pending — depends on Phase 3"],
            ["5", "Build Power Automate approval flow (card → Jira → audit)", "Pending — depends on Phase 4"],
            ["6", "End-to-end testing (5-scenario plan)", "Pending"],
        ],
    )

    # Blockers
    _add_h1(doc, "3. Blockers in detail")
    _add_h2(doc, "Blocker A — Copilot Studio \"Tools / Knowledge / Agents\" require upgrade")
    for b in [
        "What I see: the left navigation of my Copilot Studio agent shows "
        "Knowledge (needs upgrade), Tools (needs upgrade), Agents (needs "
        "upgrade). Only Topics is available without upgrade.",
        "Impact: I cannot use the textbook architecture (Tool exposing LLM "
        "extraction with structured inputs/outputs callable by Power "
        "Automate). I can work around with Topics, but it's less clean and "
        "constrains the next two sprints heavily.",
        "Action needed: approve the Microsoft Copilot Studio capability "
        "upgrade. See business-case-upgrade for the full justification.",
    ]:
        _add_bullet(doc, b)

    _add_h2(doc, "Blocker B — AI Builder Prompts disabled by admin policy")
    for b in [
        "What I see: in make.powerautomate.com → AI hub, the Prompts feature "
        "shows \"This feature has been disabled. Prompts have been disabled. "
        "Please contact your administrator.\"",
        "Impact: I cannot use AI Builder Prompts as the LLM provider for "
        "the Power Automate flow. Workaround tested and working: route the "
        "prompt through Copilot Studio instead.",
        "Action needed: ask the Power Platform admin whether Prompts can be "
        "enabled under existing governance.",
    ]:
        _add_bullet(doc, b)

    _add_paragraph(doc, "")
    _add_paragraph(doc, "Suggested message to the admin:", bold=True)
    _add_paragraph(
        doc,
        "\"We're building an AI agent for the Manco Board on Copilot Studio. "
        "The Power Automate flow needs to invoke a generative AI prompt. AI "
        "Builder Prompts is currently disabled in our tenant. Can it be "
        "enabled for the [environment name] environment under existing "
        "governance, or what is the approved path? Happy to align with the "
        "standard.\"",
        italic=True,
    )

    # Workaround
    _add_h1(doc, "4. Workaround being tested")
    _add_paragraph(doc, "In Copilot Studio's basic tier we can still:")
    for b in [
        "Build the agent and a Topic.",
        "Use the Prompt node inside the Topic (validated in a throwaway "
        "test agent).",
        "Invoke the Topic from Power Automate via the Microsoft Copilot "
        "Studio connector.",
    ]:
        _add_bullet(doc, b)
    _add_paragraph(doc, "")
    _add_paragraph(doc, "Caveats:", bold=True)
    for b in [
        "More glue code in Power Automate.",
        "Harder to monitor and govern the LLM call.",
        "Sprint 2 (Reporter) is doable on the same workaround but inelegant.",
        "Sprint 3 (Insights) is effectively blocked without Knowledge.",
    ]:
        _add_bullet(doc, b)

    # Decisions
    _add_h1(doc, "5. What I need from my manager")
    _add_docx_table(
        doc,
        headers=["#", "Decision", "Default if no decision"],
        rows=[
            ["1", "Approve Copilot Studio capability upgrade", "Sprint 1 continues on workaround; Sprints 2 and 3 deferred"],
            ["2", "Approve raising AI Builder Prompts policy with admin", "Sprint 1 continues via Copilot Studio workaround"],
            ["3", "Continue Sprint 1 on workaround in parallel?", "Yes — pause now would lose momentum"],
            ["4", "Confirm M365 Copilot licence status for me + my manager", "Unknown — useful to confirm"],
        ],
        first_col_bold=True,
    )

    # Next steps
    _add_h1(doc, "6. Suggested next steps")
    steps = [
        "Now (no decision needed): finish Phase 2 — store Jira credentials as Power Platform environment variables. Independent of the LLM path.",
        "Now (no decision needed): continue exploring the Topics-paradigm workaround for Phase 3.",
        "This week: manager review of the business case and decision on the upgrade.",
        "This week: approach the Power Platform admin re: AI Builder Prompts policy (parallel track).",
        "Once upgrade is approved: migrate Phase 3 from Topics workaround to Tools paradigm. ~30 minutes of rework — no other phase affected.",
    ]
    for n, step in enumerate(steps, 1):
        _add_paragraph(doc, f"{n}.  {step}")

    # Reference
    _add_h1(doc, "7. Reference")
    _add_bullet(doc, f"Build pack: {BRANCH_URL}/manco-capture")
    _add_bullet(doc, f"Open pull request: {PR_URL}")
    _add_bullet(doc, "Business case for upgrade: business-case-upgrade.docx (this document's sibling)")

    doc.save(path)
    print(f"  wrote {path.name}")


# ---------------------------------------------------------------------------
# PowerPoint exec deck
# ---------------------------------------------------------------------------


def _set_slide_background(slide, color: PPTXColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_textbox(slide, left, top, width, height, text: str,
                 *, size: int = 18, bold: bool = False,
                 color: PPTXColor | None = None, italic: bool = False,
                 align: str = "left", font: str = "Calibri") -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = {"left": 1, "center": 2, "right": 3}.get(align, 1)
    run = p.add_run()
    run.text = text
    run.font.size = PPTXPt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    if color is not None:
        run.font.color.rgb = color
    return box


def _add_bullets(slide, left, top, width, height,
                 bullets: list[tuple[str, int]],
                 *, size: int = 16, color: PPTXColor | None = None) -> None:
    """bullets is a list of (text, indent_level) tuples."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, (text, lvl) in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = lvl
        run = p.add_run()
        run.text = ("• " if lvl == 0 else "– ") + text
        run.font.size = PPTXPt(size)
        run.font.name = "Calibri"
        if color is not None:
            run.font.color.rgb = color


def _add_pptx_table(slide, left, top, width, height,
                    headers: list[str], rows: list[list[str]],
                    header_color: PPTXColor = PPTX_ACCENT) -> None:
    table_shape = slide.shapes.add_table(
        len(rows) + 1, len(headers), left, top, width, height
    )
    table = table_shape.table

    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        tf = cell.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].clear()
        run = tf.paragraphs[0].add_run()
        run.text = h
        run.font.size = PPTXPt(13)
        run.font.bold = True
        run.font.color.rgb = PPTX_WHITE
        run.font.name = "Calibri"

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            tf = cell.text_frame
            tf.word_wrap = True
            tf.paragraphs[0].clear()
            run = tf.paragraphs[0].add_run()
            run.text = val
            run.font.size = PPTXPt(11)
            run.font.name = "Calibri"


def _add_footer(slide, slide_number: int, total: int) -> None:
    _add_textbox(
        slide,
        PPTXInches(0.3), PPTXInches(7.0), PPTXInches(13.0), PPTXInches(0.3),
        f"{PROJECT_NAME}  ·  {STATUS_DATE}  ·  {slide_number} / {total}",
        size=10, color=PPTX_MUTED, italic=True,
    )


def build_pptx(path: Path) -> None:
    prs = Presentation()
    prs.slide_width = PPTXInches(13.333)
    prs.slide_height = PPTXInches(7.5)

    blank_layout = prs.slide_layouts[6]  # blank

    TOTAL = 11

    # ---------- Slide 1: Title ----------
    s = prs.slides.add_slide(blank_layout)
    _set_slide_background(s, PPTX_ACCENT)

    _add_textbox(
        s, PPTXInches(1.0), PPTXInches(2.4), PPTXInches(11.3), PPTXInches(1.5),
        "Manco Capture Agent",
        size=48, bold=True, color=PPTX_WHITE,
    )
    _add_textbox(
        s, PPTXInches(1.0), PPTXInches(3.7), PPTXInches(11.3), PPTXInches(0.7),
        "Business case for Copilot Studio capability upgrade",
        size=24, color=PPTX_WHITE,
    )
    _add_textbox(
        s, PPTXInches(1.0), PPTXInches(5.0), PPTXInches(11.3), PPTXInches(0.4),
        f"{AUTHOR_PLACEHOLDER}    ·    {STATUS_DATE}",
        size=14, color=PPTX_LIGHT_BG, italic=True,
    )
    _add_textbox(
        s, PPTXInches(1.0), PPTXInches(5.5), PPTXInches(11.3), PPTXInches(0.4),
        f"For {MANAGER_PLACEHOLDER} — review and decision",
        size=14, color=PPTX_LIGHT_BG, italic=True,
    )

    # ---------- Slide 2: TL;DR ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "TL;DR", size=36, bold=True, color=PPTX_ACCENT)
    _add_bullets(
        s, PPTXInches(0.7), PPTXInches(1.3), PPTXInches(12.0), PPTXInches(5.5),
        [
            ("Building an AI agent for the CIO and senior management team that automates capture, reporting and analysis of work on the Manco Board (Jira VFST2).", 0),
            ("Sprint 1 (capture: email → Jira Task via Teams approval) is approximately 60% built and on track.", 0),
            ("Hit a Microsoft Copilot Studio licence ceiling — Tools, Knowledge, and Agents features all show \"needs upgrade\" in our tenant.", 0),
            ("Requested decision:", 0),
            ("Approve the Copilot Studio capability upgrade.", 1),
            ("Approve a parallel admin policy review to enable AI Builder Prompts.", 1),
            ("Indicative cost: ~£330–£500 / month all-in for Sprints 1 + 2 capacity (Microsoft list pricing — confirm with M365 procurement).", 0),
            ("Indicative payback: under 3 months at 1 hour saved per senior leader per week.", 0),
        ],
        size=18,
    )
    _add_footer(s, 2, TOTAL)

    # ---------- Slide 3: What we're building ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "What we're building", size=32, bold=True, color=PPTX_ACCENT)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(1.0), PPTXInches(12.5), PPTXInches(0.5),
                 "Three-pillar AI capability for senior management",
                 size=16, italic=True, color=PPTX_MUTED)
    _add_pptx_table(
        s, PPTXInches(0.5), PPTXInches(1.8), PPTXInches(12.3), PPTXInches(4.0),
        headers=["Pillar", "What it does", "Status"],
        rows=[
            ["1. Capture",
             "Turns emails and Teams meeting transcripts into properly-filed Jira Tasks on the Manco Board, with human approval in Teams",
             "Sprint 1 — in build"],
            ["2. Reporter",
             "Conversational Q&A and scheduled digests over Manco Board status",
             "Sprint 2 — designed"],
            ["3. Insights",
             "Cross-source analysis of Jira + Manco SharePoint to surface risks, themes, decisions",
             "Sprint 3 — designed"],
        ],
    )
    _add_textbox(
        s, PPTXInches(0.5), PPTXInches(6.3), PPTXInches(12.3), PPTXInches(0.5),
        "All three pillars target the same audience, use the same Microsoft + Atlassian stack we already pay for, and reuse the same foundations being put in place now.",
        size=13, italic=True, color=PPTX_MUTED,
    )
    _add_footer(s, 3, TOTAL)

    # ---------- Slide 4: Why it matters ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "Why it matters", size=32, bold=True, color=PPTX_ACCENT)
    _add_bullets(
        s, PPTXInches(0.7), PPTXInches(1.4), PPTXInches(12.0), PPTXInches(5.5),
        [
            ("Action items currently fall through the cracks — emails, chats, leadership meetings tracked manually if at all.", 0),
            ("Leadership time is the scarcest resource — the CIO and senior team should be reviewing and deciding, not triaging Jira.", 0),
            ("Reporting is a known pain point — Manco status decks assembled by hand before every meeting.", 0),
            ("Insights are invisible today — trends, risks and decisions sit unseen across Jira and SharePoint.", 0),
        ],
        size=20,
    )
    _add_footer(s, 4, TOTAL)

    # ---------- Slide 5: Progress so far ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "Progress to date", size=32, bold=True, color=PPTX_ACCENT)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(1.0), PPTXInches(12.5), PPTXInches(0.5),
                 "Sprint 1 ~60% complete", size=16, italic=True, color=PPTX_MUTED)
    _add_bullets(
        s, PPTXInches(0.7), PPTXInches(1.7), PPTXInches(12.0), PPTXInches(5.0),
        [
            ("Design + build artefacts version-controlled in GitHub (manco-capture/, 12 files, ~1,200 lines).", 0),
            ("Provisioning DONE: Teams approval channel, SharePoint audit list, Jira API token, Outlook trigger category.", 0),
            ("Extraction proof-of-concept tested in Copilot Studio.", 0),
            ("Pending: LLM extraction wiring (currently blocked), Power Automate flows, end-to-end test.", 0),
        ],
        size=18,
    )
    _add_footer(s, 5, TOTAL)

    # ---------- Slide 6: The blocker ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "The blocker", size=32, bold=True, color=PPTX_DANGER)
    _add_textbox(
        s, PPTXInches(0.5), PPTXInches(1.0), PPTXInches(12.5), PPTXInches(0.7),
        "Our Copilot Studio licence is at the basic / trial tier. Three features show \"needs upgrade\" in our tenant.",
        size=16, color=PPTX_MUTED,
    )
    _add_pptx_table(
        s, PPTXInches(0.5), PPTXInches(2.0), PPTXInches(12.3), PPTXInches(4.0),
        headers=["Feature", "What it does", "Why we need it"],
        rows=[
            ["Tools",
             "Exposes an LLM prompt as a clean, callable object Power Automate can invoke",
             "Textbook integration pattern; without it we use a workaround that's harder to maintain"],
            ["Knowledge",
             "Grounds the agent in trusted content (SharePoint docs)",
             "Critical for Pillar 3 — without it Insights cannot reason over Manco SharePoint"],
            ["Agents",
             "Separate specialised agents under one orchestrator",
             "Critical for Pillars 2 & 3 — without it everything collapses into a single overloaded agent"],
        ],
    )
    _add_textbox(
        s, PPTXInches(0.5), PPTXInches(6.4), PPTXInches(12.3), PPTXInches(0.5),
        "Separately: AI Builder Prompts disabled at tenant admin level (different governance decision, in the same family).",
        size=13, italic=True, color=PPTX_MUTED,
    )
    _add_footer(s, 6, TOTAL)

    # ---------- Slide 7: With vs without upgrade ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "What an upgrade unlocks", size=32, bold=True, color=PPTX_ACCENT)
    _add_pptx_table(
        s, PPTXInches(0.5), PPTXInches(1.3), PPTXInches(12.3), PPTXInches(5.5),
        headers=["Outcome", "Without upgrade", "With upgrade"],
        rows=[
            ["LLM extraction works (Sprint 1)", "Yes — via Topics workaround", "Yes — via Tools (cleaner)"],
            ["Power Automate integration", "Bespoke Topics glue", "Native Tool invocation"],
            ["Time to build Sprint 1", "~2 working days", "~1 working day"],
            ["Reporter agent (Sprint 2)", "Single overloaded agent", "Dedicated agent with clear scope"],
            ["Insights vs SharePoint (Sprint 3)", "Blocked — no Knowledge feature", "Native RAG against Manco SharePoint"],
            ["Adding new agents over time", "Inflates the single agent", "Add as separate specialised agents"],
            ["Vendor lock-in risk", "Higher — workarounds become debt", "Lower — using supported features"],
        ],
    )
    _add_footer(s, 7, TOTAL)

    # ---------- Slide 8: Cost framing ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "Cost framing", size=32, bold=True, color=PPTX_ACCENT)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(1.0), PPTXInches(12.5), PPTXInches(0.5),
                 "Microsoft list pricing — confirm with M365 procurement",
                 size=14, italic=True, color=PPTX_MUTED)
    _add_pptx_table(
        s, PPTXInches(0.5), PPTXInches(1.6), PPTXInches(12.3), PPTXInches(2.6),
        headers=["Item", "Indicative cost"],
        rows=[
            ["Microsoft Copilot Studio (per-user)", "~£164 / user / month"],
            ["Microsoft Copilot Studio (pay-as-you-go messages)", "from ~£0.008 / message"],
            ["AI Builder add-on capacity (if Prompts enabled later)", "from ~£400 / tenant / month"],
            ["M365 Copilot licence (separate, if not already held)", "~£25 / user / month"],
        ],
    )
    _add_textbox(s, PPTXInches(0.5), PPTXInches(4.6), PPTXInches(12.3), PPTXInches(0.4),
                 "Realistic scoped cost: ~£330–£500 / month all-in for Sprints 1 + 2",
                 size=18, bold=True, color=PPTX_ACCENT)
    _add_bullets(
        s, PPTXInches(0.7), PPTXInches(5.2), PPTXInches(12.0), PPTXInches(2.0),
        [
            ("2 builder/admin seats (you + manager) on full Copilot Studio.", 0),
            ("Plus modest pay-as-you-go capacity for runtime usage.", 0),
            ("Payback: even 1 hour / week / senior leader saved covers the licence many times over within the quarter.", 0),
        ],
        size=15,
    )
    _add_footer(s, 8, TOTAL)

    # ---------- Slide 9: Options considered ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "Options considered", size=32, bold=True, color=PPTX_ACCENT)
    _add_pptx_table(
        s, PPTXInches(0.5), PPTXInches(1.3), PPTXInches(12.3), PPTXInches(5.5),
        headers=["Option", "Pros", "Cons", "Recommended?"],
        rows=[
            ["A. Do nothing — stay on basic tier",
             "Zero cost",
             "Sprints 2 & 3 effectively blocked; Sprint 1 brittle",
             "No"],
            ["B. Upgrade Copilot Studio",
             "Unlocks all three pillars; supported by Microsoft; reuses M365 footprint",
             "Licence cost",
             "Yes — recommended"],
            ["C. Route AI through Azure OpenAI",
             "Maximum control; aligns with Azure governance",
             "Org has no Azure OpenAI deployment today; sizeable setup",
             "No (today)"],
            ["D. Defer the initiative",
             "Zero cost",
             "Action-item leakage continues; Reporter/Insights value unrealised",
             "No"],
        ],
    )
    _add_footer(s, 9, TOTAL)

    # ---------- Slide 10: Recommendation & risks ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "Recommendation", size=32, bold=True, color=PPTX_SUCCESS)
    _add_bullets(
        s, PPTXInches(0.7), PPTXInches(1.3), PPTXInches(12.0), PPTXInches(2.5),
        [
            ("Approve the Microsoft Copilot Studio capability upgrade for the environment housing the Manco Capture Agent.", 0),
            ("Ask the Power Platform admin to review the AI Builder Prompts policy and consider enabling it under existing governance.", 0),
            ("Continue Sprint 1 in the meantime using the Topics workaround; migrate to Tools immediately once the upgrade lands (~30 min rework).", 0),
        ],
        size=17,
    )
    _add_textbox(s, PPTXInches(0.5), PPTXInches(4.2), PPTXInches(12.5), PPTXInches(0.5),
                 "Risks of NOT upgrading", size=22, bold=True, color=PPTX_DANGER)
    _add_bullets(
        s, PPTXInches(0.7), PPTXInches(4.9), PPTXInches(12.0), PPTXInches(2.0),
        [
            ("Architectural debt early — Sprint 1 on a fragile foundation we'll rebuild for Sprint 2/3 anyway.", 0),
            ("Sprints 2 & 3 effectively blocked without Knowledge and Agents features.", 0),
            ("Slower delivery — workarounds cost engineering time more than licences cost money.", 0),
        ],
        size=15,
    )
    _add_footer(s, 10, TOTAL)

    # ---------- Slide 11: Decision asks ----------
    s = prs.slides.add_slide(blank_layout)
    _add_textbox(s, PPTXInches(0.5), PPTXInches(0.3), PPTXInches(12.5), PPTXInches(0.7),
                 "Decision asks", size=32, bold=True, color=PPTX_ACCENT)
    _add_pptx_table(
        s, PPTXInches(0.5), PPTXInches(1.3), PPTXInches(12.3), PPTXInches(3.5),
        headers=["#", "Decision", "Your answer"],
        rows=[
            ["1", "Approve Copilot Studio capability upgrade", "yes / no / need more info"],
            ["2", "Approve admin policy review on AI Builder Prompts", "yes / no"],
            ["3", "Continue Sprint 1 on workaround in parallel", "yes / no"],
            ["4", "Procurement contact for confirming exact pricing", "_________________"],
        ],
    )
    _add_textbox(
        s, PPTXInches(0.5), PPTXInches(5.5), PPTXInches(12.3), PPTXInches(1.5),
        "Build pack and detailed docs: github.com/snerantie/Copilot-Studio  ·  PR #1",
        size=13, italic=True, color=PPTX_MUTED,
    )
    _add_footer(s, 11, TOTAL)

    prs.save(path)
    print(f"  wrote {path.name}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    print(f"Output directory: {OUT_DIR}")
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Generating Word docs...")
    build_business_case_docx(OUT_DIR / "business-case-upgrade.docx")
    build_status_docx(OUT_DIR / "status-and-blockers.docx")

    print("Generating PowerPoint deck...")
    build_pptx(OUT_DIR / "manco-capture-business-case.pptx")

    print("Done.")


if __name__ == "__main__":
    main()
