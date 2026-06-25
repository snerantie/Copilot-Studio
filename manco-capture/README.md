# Manco Capture Agent — Sprint 1 Build Pack

AI agent that turns leadership emails and Teams meeting transcripts into Jira
Tasks on the **Manco Board** (project `VFST2`), with human approval in Teams.

## Locked design decisions

| # | Decision |
|---|---|
| 1 | Scope: 12 Features on the Manco Board (each is a `VFST2-N` parent issue) |
| 2 | Hierarchy: Feature → Epic → **Task** (agent creates the Task) |
| 3 | Jira Cloud project key: `VFST2` (Enterprise, Plans enabled) |
| 4 | Default issue type: `Task` |
| 5 | Default priority: `Medium` (AI may upgrade to High based on urgency cues) |
| 6 | Approvers: user + manager, first to approve wins |
| 7 | Approval surface: Adaptive Card in a **private channel "Capture – Approvals"** inside the existing **Manco Team** |
| 8 | Routing UX: **single combined Epic dropdown** labelled `Feature › Epic` (picking the Epic implicitly picks the Feature) |
| 9 | Audit + dedup: **SharePoint list "Capture Audit"** in the Manco SharePoint site |
| 10 | Trust model: **human-in-the-loop** for Sprint 1; flip to autonomous later once accuracy proven |
| 11 | Sources Sprint 1: Outlook emails flagged with category `Capture` |
| 12 | Sources Sprint 2: Teams meeting transcripts |

## The 12 Manco Features

Confirm names against your Jira `VFST2-1 … VFST2-12` issue keys:

1. Architecture
2. Service Management
3. Audit
4. Tech Assurance
5. Cyber
6. Risk
7. PI Planning
8. PI Delivery
9. Ways of Working
10. Resourcing
11. Budget
12. VFS Exco

## Architecture (Sprint 1)

```
┌──────────────────────────────────────────────────────────────────┐
│  Outlook email categorised "Capture"                             │
└──────────────────────────────────────────────────────────────────┘
                             │
                             ▼
┌──────────────────────────────────────────────────────────────────┐
│  Power Automate — "Manco Capture: Email"                          │
│    1. Fetch live Epic list from Jira (JQL)                       │
│    2. AI Builder GPT prompt → structured action items            │
│    3. For each item: write SharePoint audit row (Pending)        │
│    4. For each item: start child flow with audit ID              │
└──────────────────────────────────────────────────────────────────┘
                             │
                             ▼
┌──────────────────────────────────────────────────────────────────┐
│  Power Automate — "Manco Capture: Approval" (child)               │
│    1. Post Adaptive Card to "Capture – Approvals", wait response │
│    2. Approve → resolve assignee → create Jira Task              │
│         → update card with VFST2-XXX link                        │
│         → update audit row (Approved + JiraKey)                  │
│    3. Reject → update card + audit row (Rejected)                │
│    4. Timeout (14d) → update audit row (Expired)                 │
└──────────────────────────────────────────────────────────────────┘
                             │
                             ▼
┌──────────────────────────────────────────────────────────────────┐
│  Jira Cloud — Task created, parented to Epic, rolls up to Feature │
└──────────────────────────────────────────────────────────────────┘
```

## Build order

| Step | File | What it gives you |
|---|---|---|
| 1 | `01-provisioning-checklist.md` | Manual setup (SharePoint list, Teams channel, Jira token in PA, AI Builder access) |
| 2 | `02-sharepoint-list-schema.md` | Exact columns for the "Capture Audit" list |
| 3 | `03-ai-extraction-prompt.md` | The GPT prompt used by AI Builder |
| 4 | `04-adaptive-card-approval.json` | Approval card template |
| 5 | `05-jira-create-task-payload.json` | Jira REST API body template |
| 6 | `06-power-automate-flow-main.md` | Step-by-step flow build (parent flow) |
| 7 | `07-power-automate-flow-approval.md` | Step-by-step flow build (child approval flow) |
| 8 | `08-test-plan.md` | Three test scenarios to validate end-to-end |

Sprint 2 (Teams meeting transcripts) reuses everything except a new trigger flow.
