# Code Review — `todayintranet` SPFx dashboard

**Date:** 2026-09-28
**Scope:** Static review of the dashboard host, persistence services, data hooks and representative widgets under [`src/webparts/todayIntranet/`](./src/webparts/todayIntranet/) for SPFx, TypeScript, React, performance, SOLID and DRY concerns. This is not an exhaustive audit or runtime certification.
**Revision:** Incorporates a second-opinion check of the original findings and recommendations.

**Overall:** the code has useful architectural foundations: module-specific Fluent UI imports, a lazy-loaded Adaptive Cards renderer, per-widget error boundaries, strict layout validation, request coalescing and substantial test code. There are worthwhile maintainability and lifecycle improvements, but this review has not established a critical production defect or measured a performance bottleneck.

### Evidence and validation limits

- The original review reported a clean `npx eslint "src/**/*.{ts,tsx}"` run.
- The second opinion inspected the effective ESLint configuration for the task widget and confirmed that no `react-hooks/*` rules were enabled.
- Temporarily enabling `rules-of-hooks` and `exhaustive-deps` in memory for the dashboard, task widget and two data hooks produced six dependency warnings and no Rules-of-Hooks errors. No configuration files were changed.
- The build, unit tests and coverage were not run as part of these reviews. Test-file volume alone does not establish test quality or coverage.
- No React profiling, production bundle measurements or cross-browser download tests were performed. Performance and browser-compatibility findings below are candidates for validation, not demonstrated failures.

---

## Priority summary

Finding numbers are retained from the original review for traceability. Priority describes the next action, not a demonstrated incident severity.

| # | Finding | Evidence | Recommended action |
|---|---------|----------|--------------------|
| 2 | Hook lint rules are not enabled | Confirmed effective configuration | Address first; triage warnings |
| 4 / 14 | Recovery download cleanup differs between implementations | Confirmed code difference; failure unverified | Test supported browsers, then standardize |
| 1 | Spread dependencies prevent static verification | Confirmed; current callers use fixed-length arrays | Harden request identity and typing |
| 3 / 6 | Mutable session/history ownership during rendering | Confirmed patterns; runtime failure not demonstrated | Review lifecycle semantics and add regression tests |
| 5 | Dashboard updates can rerender unchanged widgets | Confirmed render structure; cost unmeasured | Profile before adding memoization |
| 7 | Widget implementations are eagerly imported | Confirmed imports; bundle impact unmeasured | Measure production bundle before splitting |
| 8 | Calendar derivation runs on clock ticks | Expected behavior; cost unmeasured | Optimize only if profiling warrants it |
| 9 / 12 / 13 | Persistence complexity and repeated hook/error logic | Maintainability observations | Refactor only with a clear benefit and behavioral tests |
| 10 | Most user-facing text is not localized | Confirmed | Prioritize according to language requirements |
| 11 | Local-only store is used in tests, not the production host | Confirmed usage | Clarify its role; do not remove automatically |

Suggested sequencing: enable and triage hook linting; verify recovery exports; profile realistic dashboards; make focused lifecycle improvements; then pursue localization and shared abstractions according to product requirements.

---

## Implementation status

Validated with `npm run build` (production build and package, 401 tests passing, 0 failures) and `npx eslint "src/**/*.{ts,tsx}"` (0 problems with the hook rules enabled).

| # | Status | Change |
|---|--------|--------|
| 2 | Done | Enabled `rules-of-hooks` (error) and `exhaustive-deps` (warn) in [`eslint.config.js`](./eslint.config.js). The full codebase showed 10 warnings across 8 sites, more than the 6 in the targeted sample. Seven were resolved by refactoring. Three intentional omissions now have scoped suppressions that state the reason: the dashboard's starter layout, task-organization store acquisition, and refresh-only sync. |
| 1 | Done | Both data hooks now depend on an explicit request key built by the existing cache-key function. `useGraphData` parameters are typed as primitives (`GraphRequestParameter`). `useSearchData` derives its identity from the whole request, so its redundant `deps` argument was removed; a new regression test covers this. `widgetDataCacheKey` and `getSearchData` accept just `{ spContext }`, so effects no longer depend on a per-render context object. |
| 4 / 14 | Done; browser check pending | Shared [`downloadJson`](./src/webparts/todayIntranet/common/downloadJson.ts) is used by both recovery exports. It attaches the link, revokes the URL after 1 s (immediately if the click throws), and has unit tests for the file name, content type, cleanup and failure path. A manual download check in supported browsers is still outstanding. |
| 3 | Done | `Dashboard` now keys an inner `DashboardSession` by store identity. The session is created once per mount in a ref, not in `useMemo`, and a store switch remounts. The existing store-switch, stale-callback and draft-checkpoint tests pass unchanged. |
| 6 | Done | Task history is React state. It advances through a pure, unit-tested `advanceTaskHistory` transition using the "adjust state during render" pattern, so the ref is no longer mutated during render. |
| 13 | Done | Shared `isTransientHttpStatus` for Graph classification and messages. Search keeps its own policy, because it treats HTTP 500 as a bad query. |
| Notes | Done | The Adaptive Cards Markdown override is set once, when the library loads. The Calendar effect depends on the stable `reload` callback rather than `state.reload`. |
| 11 | Done | Documented the local-only store's role (test and alternate implementation); it was not removed. |
| 7 | Measured; not pursued | The production main bundle is 627 KB and the Adaptive Cards chunk is 326 KB. Widget-specific source is about 117 KB unminified, and the starter dashboard renders four of those widgets immediately. The bundle is mostly library code, so splitting widgets would save little and add loading waterfalls. Revisit if the catalogue grows. |
| 5 / 8 | Deferred | No profiling was done. Per the review, memoization waits for measured cost. |
| 9 / 10 / 12 | Deferred | Refactoring the persistence store, localization and a shared data-hook core need product input or a clear benefit first. The explicit request keys narrow the remaining differences between the two data hooks. |

---

## React lifecycle and reliability

### 1. Spread dependency arrays weaken static verification

[`useGraphData.ts`](./src/webparts/todayIntranet/widgets/graph/useGraphData.ts) and [`useSearchData.ts`](./src/webparts/todayIntranet/widgets/search/useSearchData.ts):

```ts
}, [spContext, scope, scopeKey, reloadToken, refreshToken, enabled, ...deps]);
```

React requires a constant number and ordering of dependencies between renders. Current callers use fixed-length arrays, so the spread alone does not demonstrate a runtime bug. It does prevent the hook linter from verifying the complete dependency list and leaves future callers responsible for preserving the contract.

Prefer an explicit request identity built from typed, supported request parameters, with a clear relationship to the cache key. Alternatively, document and constrain the existing dependency contract.

Do not blindly replace `unknown[]` with `JSON.stringify(deps)`: serialization is not equivalent to React's dependency comparison, conflates some values, ignores functions and can throw for unsupported inputs. A serialized key is appropriate only when its input domain and normalization rules are explicit.

### 2. `react-hooks` rules are registered but never enabled

The effective configuration has no enabled `react-hooks/*` rules, despite the plugin being registered. Add to [`eslint.config.js`](./eslint.config.js):

```js
rules: {
  'react-hooks/rules-of-hooks': 'error',
  'react-hooks/exhaustive-deps': 'warn'
}
```

The targeted check produced six warnings: the dashboard's omitted `starterLayout`, the task widget's source-setting dependency, and missing/spread dependencies in the two data hooks.

Triage these rather than mechanically adding every reported dependency. For example, the dashboard deliberately avoids reloading when the starter layout changes, and adding an unstable context object to a fetch effect could cause excessive requests. Refactor dependency ownership where appropriate; use a narrowly scoped suppression with a reason where the behavior is intentional and tested.

These two rules do not detect all render-phase ref mutations or misuse of memoized objects as mutable state. Enabling them will not by itself resolve findings 3 and 6.

### 3. `useMemo` used as mutable instance state

[`Dashboard.tsx`](./src/webparts/todayIntranet/components/Dashboard.tsx):

```ts
const session = React.useMemo(() => ({ store, active: true, loaded: false, ... }), [store]);
const currentSession = React.useRef(session);
currentSession.current = session;
```

`useMemo` is intended as a performance optimization, while this object's mutable fields carry lifecycle and persistence semantics. That is a design concern, but no memo-eviction or session-loss failure has been demonstrated in the current React 17 host.

Review session ownership explicitly: creation when the store changes, activation after commit, invalidation of old asynchronous callbacks, and disposal. A state/ref-backed lifecycle or a keyed session component may be clearer, but simply replacing the memo with render-time ref reassignment is not a complete concurrency-safe fix.

Several hooks also assign latest callbacks or values to refs during render. Treat these as lifecycle review points, not proof that the current app fails under StrictMode. Moving every assignment into a layout effect could change cleanup ordering, particularly the draft-setting checkpoint on unmount. Preserve that behavior with store-switch, unmount and stale-callback regression tests.

### 4. Recovery-download cleanup needs browser validation

[`Dashboard.tsx`](./src/webparts/todayIntranet/components/Dashboard.tsx), `executeAction` export branch:

```ts
link.click();
} catch { ... } finally {
  link.remove();
  if (url) { URL.revokeObjectURL(url); }   // synchronous, immediately after click()
}
```

The dashboard revokes the URL immediately, while `exportJson` in [`TasksWidget.tsx`](./src/webparts/todayIntranet/widgets/graph/TasksWidget.tsx) delays revocation by one second. This inconsistency warrants investigation because recovery export is an important reliability path.

No supported-browser failure was reproduced. Do not label this a confirmed critical defect, and do not treat a one-second timeout as proof that a download has completed.

Test both paths in supported browsers, verifying the downloaded filename and JSON contents rather than only that `click()` was called. If shared behavior is appropriate, extract a small download utility with consistent error reporting and URL cleanup. Deferred revocation can be a pragmatic compatibility measure, subject to those tests.

---

## Performance candidates

### 5. Every widget re-renders on every dashboard state change

[`Dashboard.tsx`](./src/webparts/todayIntranet/components/Dashboard.tsx) builds a fresh `widgetContext` (including a fresh `updateSettings` closure) inside `widgets.map(...)`, and [`WidgetFrame`](./src/webparts/todayIntranet/components/WidgetFrame.tsx) is not memoized. So a store-status tick, a breakpoint change, an `operationError`, or opening a panel re-renders *all* tiles and their Graph/search subtrees.

Unchanged request dependencies prevent these renders from automatically refetching data. Their CPU cost has not been measured.

Profile title edits, store-status transitions and panel changes with realistic widget counts. If expensive, introduce a memoized widget-host boundary receiving stable instance props and callbacks, and construct context inside that child. `React.memo` alone will not help while object props change on every parent render.

Do not call a context-building hook inside `widgets.map(...)`; hooks belong in a component or custom hook with stable call order. Preserve instance identity for unchanged widgets where practical.

### 6. Render-phase mutation and unmemoized aggregation in `TasksWidget`

[`TasksWidget.tsx`](./src/webparts/todayIntranet/widgets/graph/TasksWidget.tsx):

```ts
if (history.current.scope !== scope) { history.current = { scope, tasks: {...} }; }
sources.forEach((source) => {
  history.current.tasks[source] = reconcileTaskSource(history.current.tasks[source], sourceState.data);
  ...
});
const ordered = sortTasks(tasks.filter(...), state.document);
```

The ref mutation makes history changes occur during rendering rather than at an explicit committed transition. This deserves lifecycle review, especially before adopting concurrent rendering. No user-visible failure has been reproduced.

The original keystroke claim was incorrect: the inline tag editor owns child-local state, so ordinary typing does not inherently rerender the parent task widget. Reconciliation and sorting repeat on the parent's own renders, including drag-state changes.

Separate committed history updates from pure display derivation. Consider state/reducer or data-layer ownership for history, and memoize filtering/sorting only if useful. Moving the existing mutation into `useMemo` would not fix render purity because its callback also executes during render.

Cache-key serialization is another possible profiling target, but its cost has not been measured. Avoid blanket memoization of small calculations.

### 7. All widget code is eagerly bundled

[`WidgetRegistry.tsx`](./src/webparts/todayIntranet/widgets/WidgetRegistry.tsx) statically imports widget implementations. The Adaptive Cards library is already dynamically imported, but optional widget implementations remain in the eager dependency graph.

Measure the production bundle and initial-load cost before splitting it. Larger, infrequently used widgets are better candidates than every small component; the starter dashboard immediately uses several widgets and could otherwise incur extra loading waterfalls.

If splitting is justified, keep catalogue metadata eager and load implementation/settings components on demand. Ensure static imports of settings, constants or barrel exports do not pull those implementations back into the initial bundle. Provide loading and chunk-failure behavior, and verify the emitted chunks.

### 8. Calendar clock recomputation is expected; optimize only if costly

[`CalendarWidget.tsx`](./src/webparts/todayIntranet/widgets/graph/CalendarWidget.tsx) ticks every 15s via `useCalendarClock`, re-running `calendarItems` (filter + sort + map + several `toLocaleDateString` calls per event).

This does not inherently cause refetch storms: `calendarWindow` normalizes to midnight, so `start`/`end` stay stable across ordinary ticks. A separate mechanism requests periodic data refresh.

Time-sensitive labels and filtering legitimately need updating. A memo depending on `now` would still recompute every 15 seconds and therefore would not eliminate the stated work. If profiling shows a cost, separate event-dependent parsing/sorting from time-dependent labels and filtering. Otherwise, retain the simpler implementation.

---

## SOLID and maintainability

### 9. Persistence orchestration is complex; extract along proven boundaries

[`SharePointListLayoutStore.ts`](./src/webparts/todayIntranet/services/SharePointListLayoutStore.ts) coordinates remote reads, conditional writes, retries, local conflicts, legacy import and status messages. Its many mutable fields make transitions harder to reason about.

Size alone does not prove an SRP violation. Retries, pending writes, revisions and conflicts are genuinely coupled; extracting each into a separate class could increase coordination complexity.

Prefer small, demonstrably useful boundaries: pure reconciliation decisions, localized status-message formatting, or named metadata options where boolean arguments are ambiguous. Extract a scheduler only if its independent contract simplifies the store.

Before restructuring, run and extend regression tests for stale callbacks, concurrent edits, conflicts, retries, reset and disposal. Existing test code is valuable evidence of intent, not proof that the state machine is correct.

### 10. User-facing copy is hardcoded throughout

[`mystrings.d.ts`](./src/webparts/todayIntranet/loc/mystrings.d.ts) declares exactly six strings — all property-pane labels. Every string a user actually reads ("Loading your dashboard…", "Saved to SharePoint.", all recovery guidance, every widget empty state) is inline.

This is a localization gap if multilingual support is required, not a functional blocker on every non-English tenant. Use the existing SPFx localized-resource mechanism for user-facing labels and messages.

For services, consider exposing status/reason data and formatting localized messages at a presentation boundary. A standalone English message map may improve organization, but it does not by itself provide localization or remove presentation concerns from the service.

### 11. Clarify the local-only store's supported role

[`TodayIntranetWebPart.ts`](./src/webparts/todayIntranet/TodayIntranetWebPart.ts) constructs the SharePoint-backed store. The local-only implementation in [`LocalStorageLayoutStore.ts`](./src/webparts/todayIntranet/services/LocalStorageLayoutStore.ts) is used by its own tests and by [`DashboardWidgetTitles.test.tsx`](./src/webparts/todayIntranet/components/DashboardWidgetTitles.test.tsx).

That is a legitimate test/alternate-implementation role. Clarify whether it is intentionally supported, but do not delete it merely because the production host does not select it. Likewise, adding a production fallback would change behavior and should be a product decision, not a cleanup.

Implementing the same interface provides a useful dependency-inversion seam; it does not alone prove behavioral substitutability under every failure condition.

---

## DRY opportunities

### 12. `useGraphData` and `useSearchData` are near-clones

Both hooks independently implement reload/refresh tracking, cache bypass, refreshing-state retention, cancellation and reload callbacks.

They also have meaningful differences: Graph has disabled-source and identity-scope handling, while Search handles an empty query as an empty result. A shared hook must preserve or deliberately reconcile those semantics, not only swap a loader and error mapper.

A small shared asynchronous-state core may reduce maintenance. Keep source-specific policies in adapters, use explicit request identity, and validate settings changes, identity changes, stale responses, authorization errors and refresh behavior before replacing either hook.

### 13. Transient HTTP status lists duplicated three times

`429 / 500 / 502 / 503 / 504` appears in [`useGraphData.ts`](./src/webparts/todayIntranet/widgets/graph/useGraphData.ts), [`WidgetDataError.ts`](./src/webparts/todayIntranet/widgets/data/WidgetDataError.ts), and (minus 500) [`SearchService.ts`](./src/webparts/todayIntranet/widgets/search/SearchService.ts).

Search deliberately treats HTTP 500 differently because it can signal a bad query; the source already documents that behavior. A shared status predicate could reduce duplication, but only if source-specific policy remains explicit. Similar-looking lists are not sufficient reason to force identical retry or stale-data behavior.

### 14. Two blob-download implementations

Covered in finding 4. The difference is confirmed; an actual download failure is not. Share the implementation only after defining and validating the desired behavior.

---

## Smaller notes

- [`WidgetRegistry.tsx`](./src/webparts/todayIntranet/widgets/WidgetRegistry.tsx) rebuilds search text during matching, and [`AddWidgetPanel.tsx`](./src/webparts/todayIntranet/components/AddWidgetPanel.tsx) repeatedly counts existing widget types. With the current small catalogue, these are minor cleanup candidates, not established bottlenecks.
- Prefer `.find(...)` over `.filter(...)[0]` when the intent is to find one item. Treat this as clarity and allocation cleanup rather than a performance priority.
- [`WidgetAdaptiveCard.tsx`](./src/webparts/todayIntranet/widgets/content/WidgetAdaptiveCard.tsx) assigns a library-global Markdown callback inside an instance effect. Consider configuring it once when the library loads, while verifying compatibility with multiple card instances. This observation is not a security certification.
- The pure grid helpers in [`Dashboard.tsx`](./src/webparts/todayIntranet/components/Dashboard.tsx) could be extracted for direct unit testing if that simplifies maintenance. Avoid moving code solely to reduce file length.

---

## What is already done well

- Module-specific Fluent UI imports (`@fluentui/react/lib/Button` etc.) support bundle hygiene; actual emitted size still needs measurement.
- `adaptivecards` is code-split behind a named `webpackChunkName` and loaded on demand.
- `WidgetErrorBoundary` contains descendant render failures, reducing the impact of a failing widget.
- `layoutDecoder.ts` is a genuinely strict decoder: rejects cycles, non-plain prototypes, symbol keys, accessor properties, holes and out-of-range geometry.
- `WidgetDataCache` coalesces in-flight requests by key, bounds entries, and keeps stale data for transient-error fallback.
- `WidgetView` centralizes loading/error/empty/refreshing presentation for widgets that use it.
- Data-cache keys include tenant, user, site and web identity, providing useful isolation. This does not by itself prove end-to-end absence of stale cross-identity UI state.
- `ILayoutStore` is a useful architectural boundary, and the repository includes substantial store/cache test code. Test execution and coverage remain to be verified.
