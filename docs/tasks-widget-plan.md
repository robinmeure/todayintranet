# Tasks widget: read-only aggregation and personal organization

## Goal and decisions

Turn the existing `m365.tasks` tile into one place to review due work from Microsoft To Do, Planner, Loop task lists, and flagged Outlook messages. A task row shows its title, due date (if any), and a link to the source when a reliable link is available. Users can assign a **personal priority**, add **personal tags**, reorder tasks, and filter by source.

**Source systems are read-only.** The widget must not create, update, complete, flag, or delete any Microsoft 365 task or message. Priority, tags, and order are not Graph properties for this feature and must never be written back to Graph. Source title, due date, and open/closed state remain authoritative.

**Personal organization follows the user across devices** when permitted: use a per-user record stored alongside the dashboard layout but versioned and saved independently of it, with an explicitly labelled browser-only fallback. "Local" means local to this dashboard experience, not an edit to the originating app. No background service or application-wide Graph permission is required.

The original tile read only the default To Do list, fetched `maxItems + 10` tasks before sorting, and had no task link. Its settings are part of the dashboard layout; task-level metadata needs a separate lifetime and save path.

### Implementation decisions

- Successful saves show no organization status or explanatory section in the normal task view. Loading, pending, browser-only, error and conflict notices remain visible. Export and clear actions remain available under **Manage personal organization** in dashboard edit mode.
- Implemented all applicable To Do lists, assigned Planner work and opt-in flagged Outlook messages, with independent source state/cache entries. Loop remains a Planner classification, not a separate feed or filter. To Do and Planner use labelled app links; no unsupported task deep links are generated.
- Tenant responses rejected query options on both To Do list and task collection reads (`RequestBroker--ParseUri`). Both initial requests now use bare endpoints; server pagination links are followed unchanged and completion filtering is client-side. All returned records, including completed tasks, count toward the existing source budget. This intentionally forgoes `$select` on To Do; normalized records still retain only the required fields.
- The screenshot inspires the compact task table, source buttons and accent divider. The existing shared Sass typography, theme tokens and widget frame supply the application styling. Source-completion checkboxes are deliberately absent. Tags add, rename and remove inline (Enter/blur saves, Escape cancels), and the star opens an inline priority picker. A rename that collides case-insensitively with another tag on the same task is reported rather than silently dropping one, keyboard focus returns to the chip that was edited, and the add button is restored if another tile removes the tag being renamed. At narrow widths tag controls and due dates move below titles.
- Retained `Mail.ReadBasic`, which Microsoft's message-list documentation lists as least privileged; no broader Graph permissions were added. Live flag-field/filter validation, permissions and roaming remain tenant acceptance gates, not claims established by fixture tests.
- Added an `enabled` argument and scope isolation to `useGraphData`, retaining its default behavior for existing widgets. There are three source hooks, not a fourth Loop request. Source request keys exclude display-only settings.
- Personal organization uses a `TaskOrganizationJson` column and `tasks:v1:<encoded-user>` item in the existing list. Keys contain no `|`, unlike existing layout keys. Duplicate tiles share a store; separate browser tabs use separate persisted outboxes.
- Automatic organization saves reuse known ETags and await a conditional update response without reading the item back. If the response omits an ETag, a fresh version is fetched only when needed for a later save. Conflicts invalidate the cached version and trigger the existing bounded read/merge/retry path. Creation retains duplicate-record verification; 429/503 errors do not trigger immediate write-recovery reads. The 600 ms debounce and explicit/event-driven remote refresh behavior are unchanged.
- A document is limited to 500 entries and 60,000 serialized characters, with at most five 32-character tags per task. It carries a per-task logical revision (`modified`) for deterministic newest-revision merges. Concurrent order-group changes require a choice of the complete browser or SharePoint organization.
- Compact and List expose the same controls. Other stored views fall back to Compact without modifying persisted settings. Starting-source preferences use dashboard settings; the source buttons are temporary navigation, matching the existing calendar pattern.
- Default priority ordering is followed by explicit manual positions, then due-day/key ordering for the untouched tail. A task grip supports pointer drag-and-drop with an insertion indicator and Up/Down keyboard ordering. Dropping before/after another row inserts rather than swaps; crossing priority groups adopts the target priority, as requested. A drag renumbers only the rows whose position actually changes: tasks already sorted clear of the drop point keep their stored position and revision, and the unordered tail is left implicit. This keeps a drag bounded by the drop depth rather than by the number of loaded tasks, so a large group cannot exhaust the document limit on a first drag, and a reorder cannot outrank a concurrent edit to a task it did not move. Hidden/retained group positions are preserved. No automatic retention timestamps or pruning writes are performed.
- Automated tests cover adapters, limits, shared stores, independent ETags, fallback/error/recovery cases and UI interactions. A local fixture exercises the real compiled component at wide and narrow sizes. The hosted workbench requires sign-in; tenant acceptance remains outstanding.

## 1. Source inventory and identity

| Filter | Read path | Open-task rule | Source link |
| --- | --- | --- | --- |
| To Do | `GET /me/todo/lists`, then `GET /me/todo/lists/{listId}/tasks` for each applicable list | Exclude completed tasks | Graph `todoTask` has no documented task `webUrl`. Verify a supported deep link; otherwise link to the To Do app and label it accordingly. |
| Planner | `GET /me/planner/tasks` (tasks assigned to the signed-in user) | `percentComplete < 100` | Graph `plannerTask` has no documented `webUrl`. Verify a supported deep link; otherwise offer a clearly labelled Planner app link. |
| Loop | Loop task-list tasks synchronized into Planner, not a second copy of all Planner tasks | Same Planner open-task rule | Use the verified Planner task link if available; do not manufacture a Loop-page URL from a plan or task ID. |
| Outlook | `GET /me/messages` for flagged, incomplete messages in the user's mailbox, including those not synchronized to To Do | `flag.flagStatus === 'flagged'` | Message `webLink`, when present. |

The Outlook choice is deliberate: the source includes **all qualifying flagged mailbox messages**, not just To Do's "Flagged email" list. Exclude that special To Do list from the To Do feed to avoid two rows for the same flagged email. Document that primary-mailbox, delegated-mailbox, and folder coverage depend on the mailbox API and permissions; do not silently substitute a recent Inbox-only query. Verify the Graph flag filter and query behavior with a representative tenant; if server filtering is unsupported, page and filter client-side with a visible completeness limit rather than quietly omitting old flagged messages.

Loop tasks are represented by Planner's synchronized tasks. **Do not label a Planner task "Loop" based on its title, plan name, or an undocumented heuristic.** First establish whether a documented Graph field or supported plan relationship reliably identifies Loop-origin tasks in the target tenant. If not, show those tasks once under Planner and do not offer a misleading independent Loop filter; communicate this limitation before claiming Loop is separately filterable. The same rule applies to identifying tasks originating in Collaborative notes.

Use a source-scoped stable key, never a title or array index:

```ts
type TaskSource = 'todo' | 'planner' | 'outlook';

interface AggregatedTask {
  key: string; // source + underlying resource ID; include To Do list ID where required
  source: TaskSource;
  title: string;
  due: { dateTime: string; timeZone?: string } | undefined;
  sourceUrl?: string;
}
```

For Loop, retain the underlying Planner task ID in the key so moving a task between Planner and Loop presentation does not create two identities. Treat the source label as classification, not proof that the original Loop page can be linked. Investigate Outlook's immutable message ID preference when messages move between folders; To Do task IDs can change when moved between lists, so key continuity across such moves is not guaranteed by Graph. Do not guess matches using title/due date: unmatched old metadata should be retained temporarily for recovery, not attached to the wrong task.

## 2. Fetch and normalize

Implement independent source adapters that return normalized records plus source-specific status. Use delegated Graph requests through the existing SPFx client. Request only fields needed for identity, title, due date, completion/flag status, and link. Request all relevant To Do lists, rather than silently falling back to one list if none is marked default. Follow Graph `@odata.nextLink` across list and task pages. Avoid a global `top(maxItems)` before merging: it can hide the earliest-due task in another source or a later page.

Fetch enabled sources concurrently but report errors separately. One denied or throttled source must not blank healthy sources. Reuse the widget's refresh and cache patterns where appropriate, but do not cache a composite response in a way that hides source errors or treats an incomplete result as a complete sync. Bound the number of requests/records for large mailboxes; when the bound is hit, show that results are incomplete and offer a retry or narrowed filter. Respect throttling and avoid one unbounded request per task for link discovery.

**Per-source state needs a change to the shared data plumbing.** `useGraphData` returns one `IWidgetDataState` with a single `error`, a single cache entry, and a TTL selected from `GRAPH_TTL_BY_SCOPE` by one scope string — so a composite fetch declared as `Tasks.Read` would also give flagged mail the 10-minute Tasks TTL instead of Mail's 3. Choose one approach explicitly before building adapters:

- one `useGraphData` call per source, each with its own scope, cache key and TTL, plus a small merge layer that combines four states; or
- an extension of `IWidgetDataState`/`WidgetView` with a partial-success shape that carries per-source errors alongside ready data.

The first is preferred: it keeps the shared hook untouched, gives each source its own TTL and retry, and means changing `maxItems` or a filter does not invalidate all four sources at once (the cache key includes the hook's `deps`). Give each source an explicit request/page cap so a six-row tile cannot fan out into dozens of Graph calls per load.

Normalize due dates without discarding To Do/Outlook `dateTimeTimeZone.timeZone`. Interpret date-only-style due dates in their stated zone, not as an arbitrary UTC midnight; compare overdue status by calendar day in the user's locale. Planner's due date is a UTC timestamp. Missing due dates sort after dated tasks and remain visible unless "due only" is enabled. Never derive the due date of a flagged message from its received date.

At the aggregation boundary, deduplicate by underlying identity, not by equal titles. Render only open source items. Do not cache source titles or due dates as personal metadata.

## 3. Personal metadata and persistence

Persist a **versioned, validated document** scoped by hosting site/web and signed-in user. Share it across Tasks widget instances (and dashboards) on that web; keep each tile's source filters, due-only choice, item count, and view in its existing per-instance settings. The personal document contains no source task snapshots:

```json
{
  "version": 1,
  "tasks": {
    "planner:source-task-id": {
      "priority": "high",
      "tags": ["This week", "Follow up"],
      "order": 20,
      "modified": "1790588000000:device"
    }
  }
}
```

Define one local priority vocabulary (for example `high`, `normal`, `low`, with `normal` as default), independent of To Do `importance` and Planner `priority`. A missing metadata entry has default priority, no tags, and no manual position. Normalize tag whitespace/case for matching; cap tag count and length, reject invalid persisted data, and avoid storing sensitive task titles in the personal document. Cap the whole serialized document as well: it lives in a SharePoint `Note` column with a character limit of roughly 64,000, so bound the number of retained task entries and fail loudly rather than truncating. Reorder using stable integer positions within the destination priority group, renumbering it when needed. Dropping across groups adopts the destination priority. Default order is priority, then overdue/earliest due date, then a stable source/key tie-breaker. An explicitly moved item follows its manual position within its group; changing priority through the picker moves it into that group's default ordering until reordered.

**Reuse the existing layouts list rather than provisioning a second one.** `TodayIntranetLayouts` is already provisioned hidden, with `ReadSecurity: 2` / `WriteSecurity: 2`, an indexed `Title` key, ETag conditional updates, lost-response recovery and duplicate-create detection. A second list would require its own owner bootstrap — list creation is owner-only, and a visitor arriving first gets the browser-only fallback — and would duplicate that logic almost line for line. Instead, store the personal document as a **separate item in the same list**, under its own namespaced `itemKey` and its own `Note` column, by generalizing `SharePointLayoutClient` over the field name and the parser it applies:

- The layout item is keyed `` `${dashboardKey}|${userKey}` ``. Namespace the metadata key so it cannot collide with any dashboard key, and keep it within the 255-character `Title` limit.
- A separate item means a separate ETag and a separate save path: organization edits save while the dashboard is not in edit mode and do not depend on the **Done** button, which is the requirement that rules out `LayoutJson` and widget settings.
- Adding the new column still needs a site owner to visit once. Until then, fall back to browser-only organization with the existing "saved in this browser only" message rather than failing the widget.

Privacy is scoped, not absolute: `ReadSecurity: 2` hides items from other users, but site owners and collection administrators can still read the list. Since tags are user-authored free text, state this in user documentation instead of implying isolation. Keep source task data out of the SharePoint record.

Checkpoint organization edits promptly in browser storage, then publish serialized/debounced changes to SharePoint. Clearly distinguish "saved to SharePoint", "pending", "conflict", and "saved in this browser only". If the user lacks list write permission, retain browser-only organization with an honest message that it will not roam. Do not claim a write succeeded on a failed checkpoint or a failed cloud publish. Scope browser keys by site, web, and user, and provide a recovery/export path for conflict or corrupt-data cases. A stale ETag must not overwrite edits from another device: reload and reconcile before rewriting. Because the document is a disjoint map keyed by task ID, **merge per key and let the newest edit win for that key** — this is cheap and safe for priority and tags. Reserve an explicit user-facing conflict choice for competing manual-order changes within the same group, where a silent merge would visibly reshuffle the list. Never make a Graph write while reconciling metadata.

## 4. Reconciliation and lifecycle

On every successful refresh, join current normalized source tasks to the user's metadata by stable key. New tasks receive defaults. Source title/date changes show immediately while personal priority/tags/order remain intact. Source completion, unassignment, unflagging, or deletion removes a task from the visible open list on a **complete, successful refresh of that source**. Filtering a source out, losing access, hitting a pagination cap, or receiving a partial response is **not** evidence that its tasks disappeared.

Keep metadata for keys that are no longer visible; do not delete organization because a task vanished from a refresh. **Do not implement timestamp-based retention in the first release.** Recording a "last confirmed sighting" per key would turn every page load by every user into a SharePoint write, with the ETag conflicts across tabs and devices that follow — a high recurring cost for a rare recovery case. Keep entries until the user clears them explicitly, and bound growth with the document size cap from section 3. If retention is added later, bucket the confirmation timestamp so it writes at most once a day, track it in the personal record, and prune only after a complete successful refresh for that source. If the same task reappears, its organization is still there. Treat a To Do list move or an Outlook message-ID change as a new key unless a supported stable correspondence is established; do not migrate metadata heuristically.

If SharePoint metadata fails to load, do not display tasks with an empty organization as though it were the user's saved state. Show tasks with an explicit organization-load error or pending indicator, and avoid overwriting remote metadata with defaults.

## 5. Widget interaction

Display each task as **title, due date, and source link**. Show personal priority and tags as small inline editable controls/chips without turning them into source-task data. Provide source filter toggles (To Do, Planner, Outlook, and Loop only if attribution is reliable), the existing "due only" setting, and a sensible max-items limit. Filters operate before the display limit. Give an "all sources" state and distinct empty versus partial-error messages. Provide keyboard-accessible priority/tag editing and drag-and-drop ordering with Up/Down keys on the grip as the keyboard equivalent; do not put interactive edit controls inside the source-link anchor. Keep source navigation labelled accurately when only an app-level link is possible.

Preserve the widget's stable registry type `m365.tasks`, existing `maxItems`, `dueOnly`, and `view` preferences, and refresh affordance. **Restrict `supportedViews` to the views that can host edit controls.** `m365.tasks` currently opts into `ALL_ITEM_VIEWS` (`list`, `compact`, `cards`, `gallery`, `adaptive`); the adaptive view renders through the Adaptive Card path, which cannot express editable chips or move-up/down. Narrow the list to the views that can, and fall back gracefully when a tile has a persisted `view` that is no longer supported instead of rendering nothing.

**Existing tiles must not change behavior silently.** Today this widget means "my default To Do list". Absent source-filter settings must therefore default to To Do only; Planner, Loop and Outlook are opt-in per tile, so an upgrade never pulls a user's flagged mail into a dashboard unannounced or triggers a permission error on a tile that worked yesterday. Update the widget description and documentation to reflect aggregation and local organization.

## 6. Permissions and validation gates

- Keep Graph **delegated, read-only** permissions. Existing `Tasks.Read` covers the documented To Do and assigned Planner read endpoints. Verify whether querying message flags requires `Mail.Read` rather than the widget's existing `Mail.ReadBasic`; if it does, request and document the broader consent before enabling Outlook, or show a source-specific permission error. Never add `Tasks.ReadWrite` or `Mail.ReadWrite` for this feature.
- Treat a new scope as a **deployment step, not a code change**: it edits `webApiPermissionRequests` in `package-solution.json` and requires a fresh tenant-administrator approval in every tenant that already runs the solution.
- `IWidgetDefinition.requiredPermission` now accepts a string or string array; the catalogue names `Tasks.Read` and the optional Outlook `Mail.ReadBasic` scope.
- Verify Outlook flag filtering, due-date/time-zone handling, pagination, and link behavior with real Graph responses; verify To Do/Planner links rather than relying on undocumented URL templates.
- Verify whether Loop-origin Planner tasks can be positively identified before exposing a separate Loop filter.
- Add focused tests for source mapping and pagination; due/undated ordering; To Do flagged-email exclusion; incomplete and failed source refreshes; cross-device metadata merges/conflicts; duplicate widget instances; storage denial; task disappearance/reappearance; and that all Graph operations are reads — assert the latter by giving the mocked client only a reading surface, so any `post`/`patch`/`delete` fails the test.
- Run the repository's widget/store tests and build/type-check after implementation; retain the live tenant acceptance checks above.

## Proposed delivery sequence

Verification must not block implementation. Decide the fallback up front: **if a supported deep link or a documented Loop discriminator cannot be confirmed, ship v1 with app-level links and no separate Loop filter** rather than pausing the work.

**v1 — aggregation.** All To Do lists plus assigned Planner tasks, paged and merged, with per-source error states and request caps, the existing `maxItems`/`dueOnly`/`view` preferences preserved, and sources defaulting to To Do only for tiles that predate the change. No personal metadata yet.

**v2 — personal organization.** Priority, tags and reordering on a namespaced item in the existing layouts list, with browser-only fallback, per-key merge, size caps, and the load-failure and permission-denied states from sections 3 and 4.

**v3 — remaining sources.** Flagged Outlook messages once the required scope is confirmed and consented, and a Loop filter only if Loop-origin tasks can be positively identified. Update Graph consent, registry description, and user documentation, then validate the complete read-only flow and cross-device/browser-only behavior.

## Relevant references

- Existing widget and Graph caching: [`TasksWidget.tsx`](../src/webparts/todayIntranet/widgets/graph/TasksWidget.tsx), [`useGraphData.ts`](../src/webparts/todayIntranet/widgets/graph/useGraphData.ts).
- Existing storage and UI patterns: [`SharePointListLayoutStore.ts`](../src/webparts/todayIntranet/services/SharePointListLayoutStore.ts), [`SharePointLayoutClient.ts`](../src/webparts/todayIntranet/services/SharePointLayoutClient.ts), [`IWidgetContent.ts`](../src/webparts/todayIntranet/widgets/content/IWidgetContent.ts).
- Microsoft Graph: [To Do task](https://learn.microsoft.com/en-us/graph/api/resources/todotask?view=graph-rest-1.0), [To Do lists](https://learn.microsoft.com/en-us/graph/api/resources/todotasklist?view=graph-rest-1.0), [assigned Planner tasks](https://learn.microsoft.com/en-us/graph/api/planneruser-list-tasks?view=graph-rest-1.0), [Planner task](https://learn.microsoft.com/en-us/graph/api/resources/plannertask?view=graph-rest-1.0), [Outlook message listing](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0), [follow-up flag](https://learn.microsoft.com/en-us/graph/api/resources/followupflag?view=graph-rest-1.0).
- Microsoft Support: [Loop task lists synchronize with Planner](https://support.microsoft.com/en-us/loop/manage-your-tasks-from-loop-task-lists-and-collaborative-notes-in-planner), [flagged Outlook email in To Do](https://support.microsoft.com/en-us/todo/using-microsoft-to-do-with-flagged-email-from-outlook).
