# Step 4 — Adaptive Card: approval template

Paste this JSON into the **Post adaptive card and wait for a response** action
in the child flow (`07-power-automate-flow-approval.md`), substituting the
`${...}` tokens with dynamic content from the flow.

```json
{
  "type": "AdaptiveCard",
  "version": "1.5",
  "body": [
    {
      "type": "Container",
      "style": "emphasis",
      "items": [
        {
          "type": "TextBlock",
          "text": "Manco Capture — approval needed",
          "weight": "Bolder",
          "size": "Medium",
          "color": "Accent",
          "wrap": true
        },
        {
          "type": "TextBlock",
          "text": "Source: ${sourceTypeLabel} · ${sourceTimestamp}",
          "isSubtle": true,
          "spacing": "None",
          "wrap": true
        }
      ]
    },
    {
      "type": "Input.Text",
      "id": "title",
      "label": "Title",
      "value": "${draftTitle}",
      "isRequired": true,
      "maxLength": 100,
      "errorMessage": "Title is required."
    },
    {
      "type": "Input.Text",
      "id": "description",
      "label": "Description",
      "value": "${draftDescription}",
      "isMultiline": true
    },
    {
      "type": "Input.ChoiceSet",
      "id": "epicKey",
      "label": "Feature › Epic",
      "value": "${suggestedEpicKey}",
      "isRequired": true,
      "errorMessage": "Pick an Epic. New Tasks must be parented to an Epic.",
      "choices": "${epicChoices}"
    },
    {
      "type": "Input.Text",
      "id": "assigneeEmail",
      "label": "Assignee email (leave blank to leave unassigned)",
      "value": "${suggestedAssigneeEmail}"
    },
    {
      "type": "Input.ChoiceSet",
      "id": "priority",
      "label": "Priority",
      "value": "${suggestedPriority}",
      "choices": [
        { "title": "Low", "value": "Low" },
        { "title": "Medium", "value": "Medium" },
        { "title": "High", "value": "High" }
      ]
    },
    {
      "type": "Input.Date",
      "id": "dueDate",
      "label": "Due date (optional)",
      "value": "${suggestedDueDate}"
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "AI quote:", "value": "${aiQuote}" },
        { "title": "Confidence:", "value": "${confidenceLabel}" },
        { "title": "Source:", "value": "[Open original](${sourceUrl})" }
      ]
    }
  ],
  "actions": [
    {
      "type": "Action.Submit",
      "title": "Approve & create",
      "style": "positive",
      "data": { "decision": "approve", "auditId": "${auditId}" }
    },
    {
      "type": "Action.Submit",
      "title": "Reject",
      "style": "destructive",
      "data": { "decision": "reject", "auditId": "${auditId}" }
    }
  ]
}
```

## How tokens are substituted

In the Power Automate **Post adaptive card** action, the message body field
accepts Power Automate dynamic content directly. Replace each `${token}` with
the matching dynamic content from the child flow's inputs:

| Token | Source |
|---|---|
| `${draftTitle}` | trigger input `draftTitle` |
| `${draftDescription}` | trigger input `draftDescription` |
| `${suggestedEpicKey}` | trigger input `suggestedEpicKey` |
| `${suggestedAssigneeEmail}` | trigger input `suggestedAssigneeEmail` |
| `${suggestedPriority}` | trigger input `suggestedPriority` |
| `${suggestedDueDate}` | trigger input `suggestedDueDate` |
| `${sourceTypeLabel}` | trigger input `sourceTypeLabel` |
| `${sourceTimestamp}` | trigger input `sourceTimestamp` |
| `${sourceUrl}` | trigger input `sourceUrl` |
| `${aiQuote}` | trigger input `sourceQuote` |
| `${confidenceLabel}` | composed `<n>/5` |
| `${epicChoices}` | the JSON array built from `epicCatalog` in step 1 of the child flow |
| `${auditId}` | trigger input `auditId` |

## A note on the Epic dropdown

`${epicChoices}` must be a **JSON array** like:

```json
[
  { "title": "Architecture › Reference architecture refresh", "value": "VFST2-101" },
  { "title": "Cyber › IAM controls FY26", "value": "VFST2-202" }
]
```

The "Select" action in step 1 of the child flow produces exactly this shape
when fed the `epicCatalog` string. Make sure the action that posts the card
inlines this as JSON, not as a string — most Teams "Post adaptive card"
actions handle this correctly when the Message field is set to **Code view**.
