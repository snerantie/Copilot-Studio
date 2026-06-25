# Step 3 — AI Builder extraction prompt

Used by the **"Create text with GPT using a prompt"** AI Builder action in the
parent Power Automate flow. Returns structured JSON we then parse.

## Prompt inputs (defined in AI Builder)

| Input name | Type | Description |
|---|---|---|
| `sourceType` | Text | `Email` or `Transcript` |
| `sourceContent` | Text | The raw email body or transcript text |
| `epicCatalog` | Text | A newline-delimited list of available Epics |
| `featureCatalog` | Text | A newline-delimited list of the 12 Features |
| `todayIso` | Text | Today's date in ISO format, e.g. `2026-06-25` |

## Prompt body

```
You are an action-item extractor for the Manco leadership board at VFS.
Today is {todayIso}.

INPUT
Source type: {sourceType}
Source content:
"""
{sourceContent}
"""

CONTEXT — the 12 Features on the Manco Board:
{featureCatalog}

CONTEXT — Epics where Tasks can be parented (format: <epicKey> | <Feature> › <Epic title>):
{epicCatalog}

YOUR JOB
Identify discrete action items in the source. An action item is a thing
someone must DO — not background, not discussion, not opinions.

For each action item, output an object with these fields:

- title           : <= 100 chars, imperative voice ("Confirm…", "Draft…", "Approve…")
- description     : 1–3 sentences. Include context and who asked. Plain text, no markdown.
- proposed_owner  : a name or email if EXPLICITLY mentioned as the doer. Otherwise null.
- due_date        : ISO date (YYYY-MM-DD) if EXPLICITLY mentioned or unambiguously derivable
                    (e.g. "by next Friday"). Otherwise null. Never invent.
- priority_hint   : "Low" | "Medium" | "High"
                    - High: explicit urgency, named deadline within 7 days, exec escalation
                    - Medium: default
                    - Low: nice-to-have, no deadline, informational follow-up
- proposed_epic_key   : the single best-matching epicKey from the catalogue above.
                        null if no good match (confidence < 3).
- proposed_epic_label : the matching "<Feature> › <Epic title>" string. null if no match.
- confidence      : integer 1–5, where 5 = certain action item with clear epic match,
                    1 = weak signal, possibly noise.
- source_quote    : the verbatim sentence (or two) from the source justifying this item.

RULES
- If there are no genuine action items, return: {"items": []}
- Do NOT invent owners, dates, or epic mappings.
- Do NOT extract general discussion, status updates, or FYIs as action items.
- Do NOT output more than 8 items per source — pick the strongest signals.
- Output ONLY valid JSON. No prose, no markdown fences, no commentary.

OUTPUT SHAPE
{
  "items": [
    {
      "title": "...",
      "description": "...",
      "proposed_owner": null,
      "due_date": null,
      "priority_hint": "Medium",
      "proposed_epic_key": "VFST2-145",
      "proposed_epic_label": "Cyber › IAM controls FY26",
      "confidence": 4,
      "source_quote": "..."
    }
  ]
}
```

## How `epicCatalog` is built

Earlier in the parent flow, an HTTP action queries Jira:

```
GET https://<your-domain>.atlassian.net/rest/api/3/search?jql=project%20%3D%20VFST2%20AND%20issuetype%20%3D%20Epic&fields=summary,parent&maxResults=200
```

For each Epic returned, look up the parent Feature's summary (from the same
response — `parent.fields.summary`) and format one line:

```
VFST2-145 | Cyber › IAM controls FY26
```

Concatenate with newlines and pass into the prompt.

## How `featureCatalog` is built

Static — just the 12 Feature names with their VFST2-N keys, e.g.:

```
VFST2-1  | Architecture
VFST2-2  | Service Management
...
VFST2-12 | VFS Exco
```

Confirm the actual key→name mapping once you can look up `VFST2-1 … VFST2-12`
in Jira and paste it back. Until then I've assumed the order matches your list.
