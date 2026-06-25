# Step 1 — Provisioning checklist

Do these before building any flow. Tick them off as you go.

## 1.1 Teams channel

- In the existing **Manco** Team, create a **private channel** named `Capture – Approvals`.
- Members: yourself + your manager only.
- Note the channel's full path (you'll need it in the Power Automate "Post adaptive card" action).

## 1.2 SharePoint list "Capture Audit"

- In the Manco SharePoint site (the one surfaced under the Manco Team's SharePoint tab), create a new **List** called `Capture Audit`.
- Schema is in `02-sharepoint-list-schema.md`. Add each column with the exact name and type.
- Set the default view to sort by `Created` descending.

## 1.3 Jira API token

- In Atlassian: account avatar → **Manage account** → **Security** → **Create and manage API tokens** → **Create API token**.
- Label it `Manco Capture Agent` so you can revoke later if needed.
- Copy the token immediately — Atlassian won't show it again.
- Store it as a **Power Automate environment variable** or a **secure input** (do not hardcode into the flow JSON). Recommended: Power Platform **Environment Variable** of type "Secret" backed by Azure Key Vault if available; otherwise an environment variable of type "Text" in a non-production environment for now.

## 1.4 Jira API base URL

- Your Jira tenant URL is `https://<your-domain>.atlassian.net`.
- The Atlassian account email you'll authenticate with is the email tied to the API token above.
- Auth header for all Jira API calls is `Basic base64(<email>:<token>)`. Power Automate's HTTP action can compose this for you if you pick **Basic** auth.

## 1.5 AI Builder access

- Confirm your Power Platform environment has **AI Builder** capacity. Most Copilot Studio licences include it.
- You'll use the action **"Create text with GPT using a prompt"** (in some tenants it's labelled **"Run a prompt"** under the AI Builder connector).
- No model deployment is needed — it uses the platform's hosted model.

## 1.6 Outlook category

- In Outlook (web or desktop), create a category named **`Capture`** (any colour).
- This is the signal the agent listens for. Flagging an email with this category opts it into processing.
- Tell your manager to do the same on their mailbox if you want both inboxes wired in.

## 1.7 Optional — Manco Team app permissions

- If your tenant restricts the Power Automate Teams connector, ask your M365 admin to allow it for the Manco Team.

---

## Confirmation prompts before you proceed

- [ ] `Capture – Approvals` channel exists, both approvers added
- [ ] `Capture Audit` SharePoint list exists with all columns from step 2
- [ ] Jira API token created and stored as a Power Platform secret
- [ ] Outlook `Capture` category created on at least your mailbox
- [ ] AI Builder available in your Power Platform environment
