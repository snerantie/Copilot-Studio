# Step 8 — Test plan

Three scenarios to validate Sprint 1 end-to-end before going live with your
manager. Use a test inbox or a dedicated Outlook category until you're confident.

## Test 1 — Clean happy path (single clear action item)

**Setup**: send yourself an email with body:

> Subject: FY26 capex envelope — sign-off needed
>
> Hi — I need your sign-off on the FY26 capex split by next Friday. Alex Patel
> will pull the numbers but the final shape needs CIO endorsement before we
> circulate to the board. Many thanks, CFO.

Apply the `Capture` Outlook category.

**Expected**:
1. Within ~60s, an Adaptive Card appears in `Capture – Approvals`.
2. Title resembles "Approve FY26 capex split / sign off FY26 capex envelope".
3. Suggested Epic dropdown is pre-set to a Budget-related Epic.
4. Suggested assignee is Alex (or blank if no Jira user matches).
5. Suggested due date is the calculated "next Friday" ISO date.
6. Priority = Medium (or High if the model picks up urgency).
7. Confidence ≥ 3.

**Action**: tap Approve.

**Expected**:
1. Card updates to "Created VFST2-XXXX · Approved by <you>".
2. `Capture Audit` SharePoint list shows a new row, Status = Approved, JiraKey populated.
3. In Jira, the Task exists under the chosen Epic, with description containing the source quote and a link back to the email.

## Test 2 — Multi-item source

**Setup**: send an email with three discrete asks:

> Need three things before Friday: 1) Risk to confirm the top-10 register is signed off.
> 2) Architecture to draft a one-pager on the new IAM approach. 3) Service Management
> to pull last quarter's incident trends for the Manco pack.

**Expected**:
1. Three separate Adaptive Cards appear, each routed to plausible Epics
   (Risk, Architecture, Service Management).
2. Each can be approved/rejected independently.
3. `Capture Audit` shows three rows.

## Test 3 — No action items / noise

**Setup**: send an email with body:

> Just sharing the deck from yesterday — nothing needed from you. Thanks.

**Expected**:
1. No card is posted.
2. Parent flow terminates with the "No action items found" success path.
3. `Capture Audit` has no new rows.

## Test 4 — Rejection path

Repeat Test 1 but tap **Reject**.

**Expected**:
1. Card updates to "Rejected by <you>".
2. `Capture Audit` row Status = Rejected, JiraKey empty.
3. No Jira issue created.

## Test 5 — Bad Epic mapping

**Setup**: send an email where the AI is likely to mismap (e.g., a Cyber email
that mentions budget heavily). AI suggests a Budget Epic.

**Expected**:
1. Card surfaces low-ish confidence.
2. You change the Epic dropdown to a Cyber Epic and approve.
3. Jira Task is created under the **corrected** Epic, not the AI's suggestion.
4. Audit row reflects the original AI suggestion AND the final approved value
   (we don't yet log the override delta — add this to the backlog if useful for
   tuning the model later).

## Sign-off

Move to autonomous mode (no human-in-the-loop) only after **20+ approved tickets
with no corrections** across at least 5 different Features.
