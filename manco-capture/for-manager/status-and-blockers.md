# Status & blockers — Manco Capture Agent (Sprint 1)

**Status date:** 25 June 2026
**Project:** Manco Capture Agent — turns leadership emails into Jira Tasks
on the Manco Board (`VFST2`), with human approval in Teams.
**Build pack:** [GitHub repo](https://github.com/snerantie/Copilot-Studio/tree/feat/manco-capture-sprint1/manco-capture)

---

## 1. Headline

- Sprint 1 is **~60% built**.
- Foundations (Teams channel, SharePoint list, Jira token, Outlook trigger)
  are **done**.
- **Two blockers** prevent completion of the AI extraction step. Both are
  Microsoft licensing / policy decisions, not technical problems.

## 2. Progress checklist

### Phase 1 — Provisioning

| # | Task | Status |
|---|---|---|
| 1.1 | `Capture – Approvals` private Teams channel inside the Manco Team | ✅ Done |
| 1.2 | `Capture Audit` SharePoint list in the Manco SharePoint site | ✅ Done |
| 1.3 | Jira Cloud API token created and stored | ✅ Done |
| 1.4 | Outlook category `Capture` created on user mailbox | ✅ Done |
| 1.5 | AI Builder access — Models available, **Prompts disabled** | ⚠️ Partial |

### Phase 2 — Power Platform environment variables

| # | Task | Status |
|---|---|---|
| 2.1 | `manco_JiraBaseUrl` | ☐ Pending |
| 2.2 | `manco_JiraAuthEmail` | ☐ Pending |
| 2.3 | `manco_JiraApiToken` (Secret type) | ☐ Pending |

### Phase 3 — Copilot Studio extraction agent

| # | Task | Status |
|---|---|---|
| 3.1 | Create agent `Manco Capture` | ✅ Done |
| 3.2 | Define extraction prompt (text content) | ✅ Done |
| 3.3 | Wire as callable Tool with structured inputs/outputs | 🚫 **Blocked — Tools needs upgrade** |
| 3.4 | Fallback: wire as Topic with workaround | ⏳ In progress (Topics-paradigm UI is unfamiliar; investigating) |

### Phase 4 — Power Automate parent flow (Outlook → extract → dispatch)

| # | Task | Status |
|---|---|---|
| 4.x | Build flow | ☐ Pending — depends on Phase 3 |

### Phase 5 — Power Automate approval flow (card → Jira → audit)

| # | Task | Status |
|---|---|---|
| 5.x | Build flow | ☐ Pending — depends on Phase 4 |

### Phase 6 — End-to-end testing

| # | Task | Status |
|---|---|---|
| 6.x | 5-scenario test plan (see `08-test-plan.md`) | ☐ Pending |

## 3. Blockers in detail

### 🚫 Blocker A — Copilot Studio "Tools / Knowledge / Agents" require upgrade

- **What I see:** the left navigation of my Copilot Studio agent shows
  `Knowledge (needs upgrade)`, `Tools (needs upgrade)`, and `Agents (needs
  upgrade)`. Only `Topics` is available without upgrade.
- **Impact:** I cannot use the textbook architecture (a Tool exposing the
  LLM extraction prompt with structured inputs/outputs callable by Power
  Automate). I can work around this with Topics, but it's a less clean
  integration and constrains the next two sprints heavily.
- **Action needed:** approve the Microsoft Copilot Studio capability
  upgrade. See `business-case-upgrade.md` for full justification, cost
  framing, and options considered.

### 🚫 Blocker B — AI Builder Prompts disabled by admin policy

- **What I see:** in `make.powerautomate.com` → AI hub, the **Prompts**
  feature shows *"This feature has been disabled. Prompts have been
  disabled. Please contact your administrator."*
- **Impact:** I cannot use AI Builder Prompts as the LLM provider for the
  Power Automate flow. Workaround tested and working: route the prompt
  through Copilot Studio instead. Not a hard blocker for Sprint 1 but the
  integration is cleaner if Prompts is enabled.
- **Action needed:** ask the Power Platform admin whether Prompts can be
  enabled under existing governance, or what the governed path is for
  generative AI from Power Automate. Suggested message:

  > *"We're building an AI agent for the Manco Board on Copilot Studio.
  > The Power Automate flow needs to invoke a generative AI prompt. AI
  > Builder Prompts is currently disabled in our tenant. Can it be enabled
  > for the [environment name] environment under existing governance, or
  > what is the approved path? Happy to align with the standard."*

## 4. Workaround being tested

In Copilot Studio's basic tier, we can still:

- Build the agent and a Topic
- Use the **Prompt** node inside the Topic (this works — already
  validated in a throwaway test agent)
- Invoke the Topic from Power Automate via the Microsoft Copilot Studio
  connector

This gets Sprint 1 working, but with caveats:

- More glue code in Power Automate.
- Harder to monitor and govern the LLM call.
- **Sprint 2 (Reporter)** is doable on the same workaround but inelegant.
- **Sprint 3 (Insights)** is effectively blocked without Knowledge, because
  it requires grounding the agent in the Manco SharePoint documentation.

## 5. What I need from my manager

| # | Decision | Default if no decision |
|---|---|---|
| 1 | Approve Copilot Studio capability upgrade (see business case) | Sprint 1 continues on workaround; Sprints 2 and 3 deferred |
| 2 | Approve raising AI Builder Prompts policy with admin | Sprint 1 continues via Copilot Studio workaround |
| 3 | Continue Sprint 1 on workaround in parallel? | Yes — pause now would lose momentum |
| 4 | Confirm M365 Copilot licence status for me + my manager (separate licence from Copilot Studio) | Unknown — useful to confirm |

## 6. Suggested next steps

1. **Now** (no decision needed): finish Phase 2 — store the Jira credentials
   as Power Platform environment variables. Independent of the LLM path.
2. **Now** (no decision needed): continue exploring the Topics-paradigm
   workaround for Phase 3. Get the extraction working end-to-end at the
   prompt level even if the integration plumbing is suboptimal.
3. **This week:** manager review of `business-case-upgrade.md` and
   decision on the upgrade.
4. **This week:** approach the Power Platform admin re: AI Builder Prompts
   policy (parallel track).
5. **Once upgrade is approved:** migrate Phase 3 from Topics workaround to
   Tools paradigm. ~30 minutes of rework — no other phase affected.

## 7. Reference

- Build pack: [`manco-capture/`](../) in the GitHub repository
- Open pull request: [#1 — feat(manco-capture): Sprint 1 build pack](https://github.com/snerantie/Copilot-Studio/pull/1)
- Business case for upgrade: [`business-case-upgrade.md`](./business-case-upgrade.md)
