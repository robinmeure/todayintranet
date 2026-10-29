# Microsoft Graph widget load-test data

The [PowerShell seeder](../scripts/Seed-GraphTestData.ps1) creates **1,100 items per selected
type** for the interactively signed-in user. It runs separately from SPFx: dashboard permissions
remain unchanged. Use an isolated **test account and test Planner plan**, not production.
Planner tasks are visible to other plan members.

## Requirements and sign-in

- PowerShell 7 or newer (`pwsh`).
- A licensed work/school account with Exchange, To Do, and access to a **basic** Planner plan
  for the selected services. Premium Planner plans are not supported by these Graph endpoints.
- Install only the Graph authentication module (not the entire SDK):

  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```

The script signs in using `Connect-MgGraph -ContextScope Process`. It requests delegated
`User.Read`, `Tasks.ReadWrite` for Planner/To Do, and `Calendars.ReadWrite` for Calendar.
Tenant policy may require admin consent for the Microsoft Graph PowerShell application.
`-TenantId '<tenant-id>'` selects a tenant; `-UseDeviceCode` enables device-code sign-in if
permitted by your organization's policy. Only the Global/public Graph cloud is supported.

If your tenant requires a dedicated application, use `-ClientId '<public-client-app-id>'`
with `-TenantId`. Configure that Entra application for public-client interactive authentication
and the delegated permissions above, following your organization's app registration policy.
Do not supply client secrets or access tokens. The script confirms the actual account and tenant
before any writes; use that same account in the dashboard.

## Preview without signing in or writing anything

From the repository root:

```powershell
pwsh -File .\scripts\Seed-GraphTestData.ps1 -DryRun
```

This prints the item count and representative Graph payloads. It does not import the Graph
module, create files, prompt for Planner destinations, or make network requests.

## Seed data

```powershell
# Prompts for an existing test plan ID and bucket ID, then asks you to type SEED.
pwsh -File .\scripts\Seed-GraphTestData.ps1

# Supply IDs from your basic test plan; the plan and bucket themselves are never created/deleted.
pwsh -File .\scripts\Seed-GraphTestData.ps1 -Count 1100 `
  -PlanId '<test-plan-id>' -BucketId '<test-bucket-id>'

# No Planner setup needed for only To Do and Calendar (run inside PowerShell for array arguments).
.\scripts\Seed-GraphTestData.ps1 -Types Todo,Calendar -Count 1200

# Put more than 1,000 meetings on a single date for the Calendar date-picker view.
.\scripts\Seed-GraphTestData.ps1 -Types Calendar -Count 1100 `
  -StartDate '2026-11-02T08:00:00Z' -Days 1
```

`-Count` is **per type**, not a combined count. Defaults create 3,300 items plus one To Do list.
All titles carry a unique `[IntranetSeed:<run-id>]` tag, varied enterprise topics and Lorem ipsum.
To Do bodies and Planner descriptions include longer synthetic content. Tasks remain open, with
varied priority, not-started/in-progress status, and due dates including overdue dates.
Planner tasks are explicitly assigned to `/me` so they appear in **My tasks**.
To Do uses a dedicated tagged list in the user's account, not the user's existing lists.

Calendar writes to the **default calendar** so the dashboard's `/me/calendarView` sees the events.
Events default to tomorrow at 08:00 UTC, distributed across seven days and sixteen half-hour
slots per day. They intentionally overlap for load testing. Events have **no attendees, no
online meetings, no reminders, and show as free**; no invitations are sent. They still clutter
the user's real calendar, which is why a test account is essential. Change `-StartDate` and
`-Days` to match the widget's selected date/range; dates with explicit offsets are converted to UTC.

Requests are sequential and paced by `-DelayMilliseconds` (default 200). Expect large runs to
take many minutes, especially Planner (create, read details, update description per task).
HTTP 429 respects `Retry-After` (seconds or HTTP date), with bounded exponential backoff if absent.
Only reads retry transient 502/503/504 responses. Other failures stop and surface the Graph error.
SDK automatic retries are temporarily disabled to avoid unintentional duplicate POSTs.
Planner preflight follows pagination and checks the published limits: 3,000 active / 9,000 total
tasks per plan, and 9,000 tasks assigned to a user. Concurrent changes and service-specific limits
can still cause a run to stop part-way.

`-Force` skips the explicit SEED/DELETE confirmation for intentional automation; it does not skip
identity checks, preflight or sign-in.

## Cleanup and interrupted runs

Every confirmed create is immediately recorded in a local manifest, by default
`temp\graph-seed\<run-id>.json` (already gitignored). Supply `-ManifestPath` to choose another
location. Existing manifests are never overwritten by a new seed run. **Keep the manifest**
until cleanup is complete; it contains resource/user/tenant IDs, never credentials.

```powershell
pwsh -File .\scripts\Seed-GraphTestData.ps1 -Cleanup `
  -ManifestPath '.\temp\graph-seed\<run-id>.json'
```

Sign in as the same user and tenant, then type DELETE. Cleanup deletes **only recorded IDs**,
using fresh ETags for Planner and immutable event IDs for Calendar. Existing Planner plans/buckets
are untouched. The generated To Do list is deleted only after it is empty; if you added other tasks,
move them out before rerunning cleanup. Do not manually move seeded To Do tasks to other lists
before cleanup. Keep manifests private and do not edit their IDs.

Cleanup checkpoints each deletion, accepts already-absent resources, and can be rerun after an
interruption. It retains the manifest for audit. There is no resume mode for creation: clean up a
partial run before starting a new one.

There is an unavoidable window between a remote create and its local manifest write. A timeout,
process termination, or disk failure can leave an unrecorded item. The script does **not** retry
ambiguous creates automatically. Inspect the printed run tag in Planner, To Do and Outlook for
orphans before reseeding; calendar creates also carry a stable per-item `transactionId`.
Cleanup is deliberately not a broad title-search-and-delete operation.

## Exercise the widgets

1. Sign in to the dashboard with the seeded account. Refresh the widgets (or reload the page)
   after writes propagate through Microsoft 365.
2. In **Tasks**, enable both **To Do** and **Planner** sources and disable **Due only** to include
   the full seeded date spread. Check sorting, paging, source warnings, and responsiveness.
3. In **Calendar**, select the seeded date or an upcoming range covering it. Enable pagination
   to check bounded fetching; disable it to exercise all pages of that date/range.
4. Current safeguards are intentional: Tasks stops at **1,000 records / 12 requests per source**
   (including To Do list enumeration) and reports incomplete results. Calendar with pagination
   enabled only fetches its bounded display window. Creating 1,100 items does not mean every
   widget will load or render 1,100 items. This script does not change those limits.
5. Cleanup when finished and refresh again.

This seeder does not create mail, flagged emails, or Viva Engage content.

## Offline validation

```powershell
pwsh -File .\scripts\Test-GraphTestData.ps1
```

The dependency-free test harness uses fake Graph responses and temporary manifests; it never
signs in or calls Microsoft 365.

## API references

- [Create Planner task](https://learn.microsoft.com/en-us/graph/api/planner-post-tasks?view=graph-rest-1.0)
- [Create To Do task](https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks?view=graph-rest-1.0)
- [Create calendar event](https://learn.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-1.0)
- [Graph throttling](https://learn.microsoft.com/en-us/graph/throttling)
- [Planner limits](https://learn.microsoft.com/en-us/planner/planner-limits)
