# Step 2 — SharePoint list "Capture Audit" schema

Create a SharePoint list named **`Capture Audit`** in the Manco SharePoint site,
with the following columns. The internal name (in parentheses) is what Power
Automate sees in dynamic content.

## Columns

| Display name | Internal name | Type | Notes |
|---|---|---|---|
| Title | `Title` | Single line of text | Default column. Holds the AI-drafted ticket title. |
| Source type | `SourceType` | Choice | Choices: `Email`, `Transcript` |
| Source URL | `SourceUrl` | Hyperlink | Permalink back to the email or transcript |
| Source timestamp | `SourceTimestamp` | Date and time | When the source event occurred |
| Source snippet | `SourceSnippet` | Multiple lines of text (plain) | The verbatim AI quote that justified extraction |
| Extracted description | `ExtractedDescription` | Multiple lines of text (plain) | The drafted ticket description |
| Proposed Epic key | `ProposedEpicKey` | Single line of text | E.g. `VFST2-145` |
| Proposed Epic label | `ProposedEpicLabel` | Single line of text | E.g. `Cyber › IAM controls FY26` |
| Proposed assignee | `ProposedAssignee` | Single line of text | Email or name as extracted by AI |
| Proposed due date | `ProposedDueDate` | Date only | Null if AI couldn't infer one |
| Proposed priority | `ProposedPriority` | Choice | Choices: `Low`, `Medium`, `High`. Default: `Medium` |
| Confidence | `Confidence` | Number | 1–5 |
| Status | `Status` | Choice | Choices: `Pending`, `Approved`, `Rejected`, `Expired`, `Error`. Default: `Pending` |
| Approver | `Approver` | Person or group | Filled when card is responded to |
| Decision timestamp | `DecisionTimestamp` | Date and time | When the approval/rejection happened |
| Jira key | `JiraKey` | Single line of text | E.g. `VFST2-2031`. Empty until created. |
| Jira URL | `JiraUrl` | Hyperlink | Link to the created Task |
| Card message ID | `CardMessageId` | Single line of text | Teams message ID of the posted card, for later updates |
| Error message | `ErrorMessage` | Multiple lines of text (plain) | Set if `Status = Error` |

## Views

- **All items** (default): sort by `Created` desc.
- **Pending**: filter `Status = Pending`. Use this to spot stuck approvals.
- **Created this week**: filter `Status = Approved AND Created >= [Today]-7`. Useful for the future Reporter agent.

## Permissions

- Inherit from the Manco SharePoint site.
- The Power Automate flow's connection user must have at minimum **Contribute** rights.
