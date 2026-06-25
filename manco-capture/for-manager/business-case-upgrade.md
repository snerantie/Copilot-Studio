# Business case — Copilot Studio capability upgrade

**Author:** [your name] · **Date:** 25 June 2026
**For:** Manager review · **Initiative:** Manco Capture Agent (Sprint 1 of three)

---

## TL;DR

We are building an AI agent for the CIO and senior management team that
automates capture, reporting, and analysis of work on the **Manco Board**
(Jira `VFST2`). Sprint 1 (capture from email → Jira Task with one-tap
approval in Teams) is ~60% built and on track.

We have hit a **Microsoft Copilot Studio licence ceiling** that prevents us
from using three features needed for a clean architecture and for the next
two sprints. The features are: **Tools**, **Knowledge**, and **Agents** —
each currently shows *"needs upgrade"* in our tenant.

**Requested decision:** approve a Copilot Studio capability upgrade and a
parallel admin policy review to enable AI Builder Prompts under our existing
governance.

**Indicative cost:** £125–£170 per user / month for full Copilot Studio
licensing (Microsoft list pricing — confirm with M365 procurement).
**Indicative payback:** under 3 months at a conservative 1 hour / senior
leader / week of saved Jira hygiene and action-item triage.

---

## 1. What we are building

A three-pillar AI capability for senior management:

| Pillar | What it does | Status |
|---|---|---|
| **1. Capture** | Turns emails and Teams meeting transcripts into properly-filed Jira Tasks on the Manco Board, with human approval in Teams | Sprint 1 — in build |
| **2. Reporter** | Conversational Q&A and scheduled digests over Manco Board status ("what's overdue?", "give me a status for Cyber") | Sprint 2 — designed |
| **3. Insights** | Cross-source analysis of Jira history + Manco SharePoint docs to surface risks, themes, and decisions needing escalation | Sprint 3 — designed |

All three pillars target the same audience (CIO + senior management team), use
the same Microsoft + Atlassian stack we already pay for, and reuse the same
foundations we are putting in place now.

## 2. Why it matters

- **Action items currently fall through the cracks.** Items raised in
  emails, Teams chats, and leadership meetings are tracked manually if at
  all. Jira hygiene depends on whoever remembers.
- **Leadership time is the scarcest resource.** The CIO and senior team
  should not be triaging Jira; they should be reviewing and deciding.
- **Reporting is a known pain point.** Status updates are assembled by hand
  before every Manco meeting. The Reporter pillar removes that work.
- **Insights are invisible today.** Trends and risks across the Manco
  Board sit in Jira and SharePoint with no integration layer. The Insights
  pillar surfaces them automatically.

## 3. Progress to date

- Design and build artefacts version-controlled in the
  [`Copilot-Studio` GitHub repository](https://github.com/snerantie/Copilot-Studio), under
  [`manco-capture/`](https://github.com/snerantie/Copilot-Studio/tree/feat/manco-capture-sprint1/manco-capture).
- 9-file build pack: architecture, provisioning checklist, SharePoint list
  schema, AI extraction prompt, Adaptive Card design, Jira API payload,
  Power Automate flow guides, 5-scenario test plan.
- Sprint 1 provisioning complete: Teams approval channel, SharePoint audit
  list, Jira API token, Outlook trigger category.
- Extraction proof-of-concept tested in Copilot Studio.

## 4. The blocker

Our current Copilot Studio licence is at the **basic / trial tier**. Three
features needed for a production-quality build are gated behind a paid
upgrade and show *"needs upgrade"* in the interface:

| Feature | What it does | Why we need it |
|---|---|---|
| **Tools** | Lets us expose a Prompt (LLM call) as a clean, callable object that Power Automate can invoke with structured inputs and outputs | The textbook integration pattern. Without it, we work around through Topics — functional but less maintainable, harder to monitor, harder to evolve. |
| **Knowledge** | Grounds the agent's responses in trusted content sources such as the Manco SharePoint document library | **Critical for Pillar 3 (Insights).** Without it the Insights agent cannot read or reason over our existing Manco documentation. |
| **Agents (multi-agent)** | Lets us build separate specialised agents (Capture / Reporter / Insights) with one orchestrator at the front | **Critical for Pillars 2 and 3.** Without it everything must collapse into a single overloaded agent — bad architecture, hard to govern, hard to extend. |

Separately, **AI Builder Prompts** is disabled at the tenant admin level
(*"feature has been disabled — please contact your administrator"*). This
is a different governance decision but in the same family: it limits where
generative AI can be invoked from. Reviewing the admin policy in parallel
would give us a more flexible architecture.

## 5. What an upgrade unlocks

### For Sprint 1 (Capture)

| Outcome | Without upgrade | With upgrade |
|---|---|---|
| LLM extraction works | Yes (via Topics workaround) | Yes (via Tools — cleaner) |
| Power Automate integration | Yes but bespoke Topics-based glue | Native Tool invocation |
| Time to build Sprint 1 | ~2 working days | ~1 working day |

### For Sprint 2 (Reporter) and Sprint 3 (Insights)

| Outcome | Without upgrade | With upgrade |
|---|---|---|
| Conversational Reporter agent | Single overloaded agent, harder to govern | Dedicated agent with clear scope |
| Insights against SharePoint docs | **Blocked** — no Knowledge feature | Native RAG against Manco SharePoint |
| Future addition of new agents (e.g., Risk, Audit) | Each one inflates the single agent | Add as separate agents with orchestrator |
| Vendor lock-in risk | Higher — workarounds become bespoke debt | Lower — using supported features |

## 6. Cost framing

Microsoft list pricing (verify with M365 procurement; pricing changes):

| Item | Indicative cost |
|---|---|
| Microsoft Copilot Studio (per-user) | ~£164 / user / month |
| Microsoft Copilot Studio (pay-as-you-go message capacity) | from ~£0.008 / message |
| AI Builder add-on capacity (if Prompts later enabled) | from ~£400 / tenant / month (volume-tied) |
| M365 Copilot licence (separate, if not already held) | ~£25 / user / month |

**Realistic scoped cost for our use case:**

- 2 builder/admin seats (you + your manager) on full Copilot Studio
- Plus modest pay-as-you-go message capacity for runtime usage by the agent
- **~£330–£500 / month all-in** for Sprint 1 + Sprint 2 capacity, scaling
  modestly with usage.

**Payback estimate:**

- Senior management team day rate is high. Even **1 hour per week per
  member** saved (action-item triage, status assembly, manual Jira hygiene)
  pays for the licence many times over within the quarter.
- The Reporter pillar alone removes manual Manco pack assembly — typically
  several hours per cycle.

## 7. Options considered

| Option | Pros | Cons | Recommended? |
|---|---|---|---|
| **A. Do nothing — stay on basic tier** | Zero cost | Sprint 2 and 3 effectively blocked; Sprint 1 architecture is brittle | No |
| **B. Upgrade Copilot Studio** *(this paper)* | Unlocks all three pillars; supported by Microsoft; reuses existing M365 footprint | Licence cost | **Yes — recommended** |
| **C. Route generative AI through Azure OpenAI Service** | Maximum control, ties into existing Azure governance | Requires Azure OpenAI deployment which our org does not currently have; sizeable upfront setup | No (today) |
| **D. Defer the whole initiative** | Zero cost | Action-item leakage continues; Reporter and Insights value not realised | No |

## 8. Risks of not upgrading

1. **Architectural debt early.** Sprint 1 will work but on a fragile
   foundation that we'll likely rebuild when we get to Sprint 2/3 anyway.
2. **Sprint 2 and 3 are effectively blocked.** Without Knowledge we cannot
   reason over Manco SharePoint. Without Agents we cannot cleanly modularise.
3. **Slower delivery.** Workarounds take longer to build and longer to
   maintain. We trade licence cost for engineering time.
4. **Higher cost of change.** Replacing workarounds with proper features
   later is more expensive than doing it right now.

## 9. Recommendation

1. **Approve** the Microsoft Copilot Studio capability upgrade for the
   environment housing the Manco Capture Agent.
2. **Ask** the Power Platform admin to review the AI Builder Prompts policy
   and consider enabling it under existing governance.
3. **Continue** Sprint 1 in the meantime using the Topics-based workaround;
   migrate to Tools immediately once the upgrade lands. No code is wasted —
   the only piece that changes is the LLM invocation step.

## 10. Decision asks

- [ ] **Approve upgrade**: yes / no / need more info
- [ ] **Approve admin policy review on AI Builder Prompts**: yes / no
- [ ] **Continue Sprint 1 on workaround in parallel**: yes / no
- [ ] **Procurement contact for confirming exact pricing**: _________

---

## Appendix A — Glossary

- **Copilot Studio**: Microsoft's low-code platform for building AI agents
  on top of Microsoft 365.
- **Tool / Action**: a callable unit within a Copilot Studio agent —
  typically a Prompt (LLM call), an API call, or a Power Automate flow.
- **Knowledge**: trusted content source (SharePoint sites, documents,
  websites) the agent grounds its answers in (Retrieval-Augmented
  Generation).
- **Agent / multi-agent**: separate specialised agents that can hand off to
  each other under one orchestrator.
- **Topic**: the classic Copilot Studio building block — a scripted
  conversational flow with phrases and nodes.
- **AI Builder Prompts**: an AI Builder feature that lets Power Automate
  flows directly invoke a GPT-style prompt.

## Appendix B — Reference architecture (Sprint 1)

See [`manco-capture/README.md`](../README.md) in the repository.

## Appendix C — Build pack

See [`manco-capture/`](../) — 9 files covering provisioning, schema,
prompt, Adaptive Card, Jira payload, both Power Automate flows, and test
plan.
