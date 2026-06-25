# Step 6 — Power Automate flow: "Manco Capture: Email" (parent)

This is the main flow. Triggered by Outlook category change, extracts action
items, and dispatches each one to the child approval flow.

## Trigger

- Connector: **Office 365 Outlook**
- Trigger: **When an email arrives (V3)** (or **When a flagged email arrives** — see note)
- Folder: `Inbox`
- Importance: Any
- Only with Attachments: Either
- **Advanced — Category Contains**: `Capture`

> Note: Outlook doesn't expose a direct "category added" trigger. The common
> pattern is to trigger on every email arrival but use the "Has category"
> filter, OR to trigger on **Move/categorise** via a manual button + Power
> Automate "Selected email" trigger. For v1, simplest is: **two trigger flows
> in parallel** — one on the user's mailbox, one on the manager's, both
> filtered by category `Capture`.
>
> Alternative cleaner pattern: use the **"For a selected email"** manual trigger
> and have the approvers click "Run flow" from Outlook's flow menu on any
> email they want captured. No category required. Pick whichever fits your
> workflow.

## Step-by-step actions

### 1. Initialize variable — `epicCatalog`
- Type: String
- Value: empty (populated in step 3)

### 2. HTTP — Fetch Epics from Jira
- Method: `GET`
- URI:
  ```
  https://<your-domain>.atlassian.net/rest/api/3/search?jql=project%20%3D%20VFST2%20AND%20issuetype%20%3D%20Epic&fields=summary,parent&maxResults=200
  ```
- Auth: Basic. Username = your Atlassian email. Password = the API token stored in env variable.

### 3. Parse JSON — Epics response
- Content: `body('HTTP_Fetch_Epics')`
- Schema: paste the response schema (you can generate from a sample after first run).

### 4. Apply to each — Epic
- Inputs: `body('Parse_JSON_-_Epics_response')?['issues']`
- Inside the loop:
  - **Append to string variable** `epicCatalog`:
    ```
    @{items('Apply_to_each_Epic')?['key']} | @{items('Apply_to_each_Epic')?['fields']?['parent']?['fields']?['summary']} \u203a @{items('Apply_to_each_Epic')?['fields']?['summary']}
    
    ```
    (Each line ends with a newline so the prompt sees a clean list.)

### 5. Compose — `featureCatalog`
- Inputs:
  ```
  VFST2-1  | Architecture
  VFST2-2  | Service Management
  VFST2-3  | Audit
  VFST2-4  | Tech Assurance
  VFST2-5  | Cyber
  VFST2-6  | Risk
  VFST2-7  | PI Planning
  VFST2-8  | PI Delivery
  VFST2-9  | Ways of Working
  VFST2-10 | Resourcing
  VFST2-11 | Budget
  VFST2-12 | VFS Exco
  ```
  (Update these once you've confirmed the actual VFST2-N → name mapping in Jira.)

### 6. AI Builder — Create text with GPT using a prompt
- Prompt: the one in `03-ai-extraction-prompt.md`. Create it in AI Builder Studio, name it `ManCo Capture Extractor`.
- Inputs:
  - `sourceType`: `Email`
  - `sourceContent`: `body('Get_email_(V2)')?['Body']` (HTML stripped — use the `Html to text` action first if needed)
  - `epicCatalog`: `variables('epicCatalog')`
  - `featureCatalog`: `outputs('Compose_-_featureCatalog')`
  - `todayIso`: `formatDateTime(utcNow(), 'yyyy-MM-dd')`

### 7. Parse JSON — extraction output
- Content: the text output of the AI Builder action.
- Schema:
  ```json
  {
    "type": "object",
    "properties": {
      "items": {
        "type": "array",
        "items": {
          "type": "object",
          "properties": {
            "title": { "type": "string" },
            "description": { "type": "string" },
            "proposed_owner": { "type": ["string", "null"] },
            "due_date": { "type": ["string", "null"] },
            "priority_hint": { "type": "string" },
            "proposed_epic_key": { "type": ["string", "null"] },
            "proposed_epic_label": { "type": ["string", "null"] },
            "confidence": { "type": "integer" },
            "source_quote": { "type": "string" }
          }
        }
      }
    }
  }
  ```

### 8. Condition — items empty?
- If `length(body('Parse_JSON_-_extraction_output')?['items'])` is `0`:
  - **Terminate (Succeeded)** with status message "No action items found."
- Else: continue.

### 9. Get email permalink
- Action: **Get email (V2)** if not already loaded.
- Use `webLink` from the email properties for the source URL.

### 10. Apply to each — extracted item
- Inputs: `body('Parse_JSON_-_extraction_output')?['items']`
- Inside the loop:

  #### 10a. Create item — SharePoint "Capture Audit"
  Map fields:
  | SharePoint column | Value |
  |---|---|
  | Title | `items('Apply_to_each_item')?['title']` |
  | SourceType | `Email` |
  | SourceUrl | email's `webLink` |
  | SourceTimestamp | email's `DateTimeReceived` |
  | SourceSnippet | `items('Apply_to_each_item')?['source_quote']` |
  | ExtractedDescription | `items('Apply_to_each_item')?['description']` |
  | ProposedEpicKey | `items('Apply_to_each_item')?['proposed_epic_key']` |
  | ProposedEpicLabel | `items('Apply_to_each_item')?['proposed_epic_label']` |
  | ProposedAssignee | `items('Apply_to_each_item')?['proposed_owner']` |
  | ProposedDueDate | `items('Apply_to_each_item')?['due_date']` |
  | ProposedPriority | `items('Apply_to_each_item')?['priority_hint']` |
  | Confidence | `items('Apply_to_each_item')?['confidence']` |
  | Status | `Pending` |

  Capture the new item's `ID` for the next step.

  #### 10b. Run a Child Flow
  - Flow: **Manco Capture: Approval** (see `07-power-automate-flow-approval.md`)
  - Inputs:
    - `auditId`: the SharePoint item ID
    - `epicCatalog`: `variables('epicCatalog')` (for building the dropdown)
    - All draft fields the card needs

  Run the child flow **asynchronously** (un-tick "Wait for completion") so the
  parent isn't blocked. Each draft ticket then lives in its own child flow run,
  awaiting approval independently.
