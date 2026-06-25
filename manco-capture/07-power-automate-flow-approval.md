# Step 7 — Power Automate flow: "Manco Capture: Approval" (child)

Handles a single draft ticket from card post → approval → Jira creation → audit
update. One run per draft.

## Trigger

- **Manually trigger a flow** with inputs (this lets it be called as a child flow):

| Input | Type |
|---|---|
| `auditId` | Number |
| `draftTitle` | Text |
| `draftDescription` | Text |
| `suggestedEpicKey` | Text |
| `suggestedEpicLabel` | Text |
| `suggestedAssigneeEmail` | Text |
| `suggestedPriority` | Text |
| `suggestedDueDate` | Text |
| `sourceTypeLabel` | Text |
| `sourceUrl` | Text |
| `sourceTimestamp` | Text |
| `sourceQuote` | Text |
| `confidence` | Number |
| `epicCatalog` | Text (newline-delimited `<key> | <Feature> › <Epic>` list) |

## Step-by-step actions

### 1. Compose — `epicChoices` (Adaptive Card choices array)

Build a JSON array from `epicCatalog`. In Power Automate:

- **Select** action:
  - From: `split(triggerBody()?['epicCatalog'], decodeUriComponent('%0A'))`
  - Map:
    - `title` = `trim(last(split(item(), '|')))`
    - `value` = `trim(first(split(item(), '|')))`
- Save into a variable `epicChoices`.

### 2. Compose — `confidenceLabel`

```
@{triggerBody()?['confidence']}/5
```

### 3. Post adaptive card and wait for a response (Teams)

- Connector: **Microsoft Teams**
- Action: **Post adaptive card and wait for a response**
- Post as: **Flow bot**
- Post in: **Channel**
- Team: `Manco`
- Channel: `Capture – Approvals`
- Message: paste the JSON from `04-adaptive-card-approval.json`, substituting tokens with dynamic content:
  - `${draftTitle}` → input `draftTitle`
  - `${draftDescription}` → input `draftDescription`
  - `${suggestedEpicKey}` → input `suggestedEpicKey`
  - `${suggestedAssigneeEmail}` → input `suggestedAssigneeEmail`
  - `${suggestedPriority}` → input `suggestedPriority`
  - `${suggestedDueDate}` → input `suggestedDueDate`
  - `${sourceTypeLabel}` → input `sourceTypeLabel`
  - `${sourceTimestamp}` → input `sourceTimestamp`
  - `${sourceUrl}` → input `sourceUrl`
  - `${aiQuote}` → input `sourceQuote`
  - `${confidenceLabel}` → composed value above
  - `${epicChoices}` → the variable from step 1 (raw JSON)
  - `${auditId}` → input `auditId`
- Update message: leave default (`Thanks for your response!`).
- Set the timeout under **Settings → Time out**: `P14D` (14 days).

This action **pauses the flow** until someone clicks Approve or Reject. The
outputs include all form values + which button was pressed.

### 4. Get item — SharePoint Capture Audit (`auditId`)

Used to confirm the row exists and grab any fields we didn't pass through.

### 5. Switch — on `decision` from card response

#### Case `approve`

##### 5a. Resolve assignee accountId (if email provided)
- Condition: `assigneeEmail` from card response is not empty.
- HTTP GET:
  ```
  https://<your-domain>.atlassian.net/rest/api/3/user/search?query=@{outputs('Adaptive_Card_response')?['body/data/assigneeEmail']}
  ```
- Parse JSON, take first match's `accountId`. If no match, leave assignee unset.

##### 5b. Build description ADF
- Compose using template in `05-jira-create-task-payload.json`, substituting:
  - `${title}`, `${description}`, `${epicKey}`, `${priority}`, `${dueDateOrOmit}` (use `null` or omit the field if empty), `${assigneeAccountIdOrOmit}`, `${sourceQuote}`, `${sourceTypeLabel}`, `${sourceUrl}`, `${auditId}`.
- Important: if `dueDate` is empty, OMIT the field entirely from the body. Same for `assignee`.

##### 5c. HTTP POST — Create Jira Task
- Method: `POST`
- URI: `https://<your-domain>.atlassian.net/rest/api/3/issue`
- Auth: Basic (same credentials)
- Body: the composed JSON
- Headers: `Content-Type: application/json`, `Accept: application/json`

##### 5d. Parse JSON — Jira create response
- Capture `key` (e.g. `VFST2-2031`) and `self` URL.
- Build issue URL: `https://<your-domain>.atlassian.net/browse/@{body('Parse_Jira_response')?['key']}`

##### 5e. Update SharePoint audit row
- Status: `Approved`
- JiraKey: the new key
- JiraUrl: the issue URL
- Approver: from card response (the responder's user)
- DecisionTimestamp: `utcNow()`

##### 5f. Update Teams message — replace card with success
- Action: **Update an adaptive card in a chat or channel** (or post a reply if your tenant restricts updates).
- New card: a simple confirmation card showing `Created VFST2-XXX  ·  Approved by <user>  ·  <link>`.

#### Case `reject`

##### 5g. Update SharePoint audit row
- Status: `Rejected`
- Approver: from card response
- DecisionTimestamp: `utcNow()`

##### 5h. Update Teams message — replace card with rejected state
- Show `Rejected by <user>  ·  <timestamp>` and the original title for context.

### 6. Configure run-after — error handling

On any action's failure path:
- Update SharePoint audit row: Status = `Error`, ErrorMessage = the failure detail.
- Post a small error card into `Capture – Approvals` so the failure is visible.

### 7. Timeout

If the "Post adaptive card and wait for a response" times out after 14 days:
- Update SharePoint audit row: Status = `Expired`.
- Optionally post a follow-up nudge.
