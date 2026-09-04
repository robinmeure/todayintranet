# Today intranet

A full-width SharePoint Framework (SPFx) web part that hosts a grid of widgets. Visitors switch the
dashboard into **edit mode from the page itself** — no SPFx page editing, no property pane — to add,
move, resize and remove widgets. Each person's arrangement is saved for them personally.

Built on SPFx **1.23.2** with the Heft toolchain (gulp is no longer used), React 17 and Fluent UI v8.

## Getting started

```bash
npm install
npm start                 # heft start --clean, opens the hosted workbench
npm run build             # heft test --production && heft package-solution --production
```

The hosted workbench target lives in `config/serve.json`. First run on a machine also needs the dev
certificate:

```bash
npx heft trust-dev-cert
```

Useful Heft actions (old gulp equivalents in brackets): `build` [build + bundle], `start` [serve],
`package-solution` [package-solution], `clean` [clean], `dev-deploy` (new — pushes built assets to a
test CDN), `trust-dev-cert`, `untrust-dev-cert`. `--production` replaces the old `--ship`.

Deploy `sharepoint/solution/todayintranet.sppkg` to the tenant app catalog. `skipFeatureDeployment`
is on, so the web part is available on every site once the package is approved. Approve the Graph
permission requests afterwards (see [Widgets](#widgets)).

## How the dashboard fits together

```
TodayIntranetWebPart.ts           creates the layout store, captures the theme, renders <Dashboard>
components/Dashboard.tsx          grid, edit mode, add/remove, keyboard nudging, debounced persistence
components/WidgetFrame.tsx        tile chrome: icon, title, drag handle, settings flyout, remove
components/WidgetErrorBoundary.tsx contains a crashing widget
components/AddWidgetPanel.tsx     widget catalogue
components/LayoutPresetPanel.tsx  layout picker: column presets and tile height
widgets/IWidget.ts                the widget contract
widgets/WidgetRegistry.tsx        every widget the catalogue offers
widgets/WidgetMessage.tsx         shared loading / empty / error presentation
widgets/graph/useGraphData.ts     Graph query hook with consent-aware error handling
model/IDashboardLayout.ts         persisted layout shape
model/LayoutPresets.ts            layout presets and the code that pours widgets into them
services/*LayoutStore.ts          persistence
```

Layout is authored on a 12 column grid (`react-grid-layout`). Narrower breakpoints are derived from
that canonical layout by clamping widths, and rearranging is only saved at the widest breakpoint so a
phone visit can never scramble someone's desktop layout.

The web part sets `supportsFullBleed: true`, so place it in a **full-width section** on a
communication-site page to get edge-to-edge rendering.

### Rearranging

In edit mode a widget can be dragged by its title bar, or moved from the keyboard: Tab to a title
bar, then arrow keys to move and Shift+arrows to resize. Because the grid compacts vertically, a
vertical nudge swaps the tile with whatever sits above or below it rather than leaving a gap.

### Layouts

**Choose a layout** in edit mode opens a picker with two choices that together define the grid:

| Columns | Tile height |
| --- | --- |
| Single column, Two columns, Three columns, Four columns | Short (4 rows) |
| Main and sidebar, Sidebar and main | Medium (6 rows) |
| Banner and two columns, Banner and three columns | Tall (9 rows) |

Applying a layout pours the widgets already on the dashboard into its slots, in their current order —
it never adds or removes anything. The chosen preset stays in force: adding or removing a widget
reflows into the same slots. Dragging or nudging a tile by hand switches the layout to **Custom**,
shown in the edit toolbar, and nothing reflows again until a preset is applied.

Presets are declared in `model/LayoutPresets.ts` as the column spans of one repeating band, so a new
one is a single entry in `LAYOUT_PRESETS`. Two rules keep a preset stable against the grid's vertical
compaction, which would otherwise pull tiles out of the arrangement the user picked: every tile in a
band gets the band's height, and a band cut short by an oversized widget stretches its last tile to
close the gap. Preset and row-size ids are persisted, so never rename one that has shipped.

## Widgets

| Type key | Widget | Graph permission |
| --- | --- | --- |
| `m365.calendar` | Calendar — upcoming events, configurable days ahead | `Calendars.ReadBasic` |
| `m365.mail` | My mail — inbox, optional unread-only | `Mail.ReadBasic` |
| `m365.tasks` | My tasks — open Microsoft To Do items | `Tasks.Read` |
| `demo.clock` | Clock | — |
| `demo.welcome` | Welcome greeting (reference implementation) | — |

The three Microsoft 365 widgets need tenant admin approval of the `webApiPermissionRequests` in
`config/package-solution.json`, granted in **SharePoint admin center > Advanced > API access** after
the package is deployed. Until that happens each widget shows a message naming the missing
permission rather than failing silently. The `.ReadBasic` scopes are deliberate: SPFx grants
permissions to the tenant-wide SharePoint Online Client Extensibility principal, not to this
solution alone, so the widgets ask only for the metadata they render — never message or event
bodies and attachments.

### Adding a widget

1. Write a component that takes an `IWidgetContext` (instance id, its own `settings`, the SPFx
   `WebPartContext`, `isEditing`, and `updateSettings`). Use `useGraphData` for anything that calls
   Microsoft 365 — it handles loading, cancellation, throttling and missing consent.
2. Register an `IWidgetDefinition` for it in `widgets/WidgetRegistry.tsx` with a **stable** `type`
   key — that key is what lives in saved layouts, so never rename one that has shipped.
3. Optionally implement `renderSettings` to get a gear icon and settings flyout on the tile.

Widgets are wrapped in an error boundary, so a widget that throws shows a contained message instead
of blanking the dashboard.

## Where layouts are stored

Per user, in a hidden list named **TodayIntranetLayouts** on the site hosting the page: `Title` holds
`<dashboard id>|<user login name>` and a `LayoutJson` note column holds the layout.

The **Dashboard id** is a property-pane setting that defaults to `default`, so out of the box a
person's arrangement is per site and survives the web part being removed, re-added or moved to
another page. Give a second dashboard on the same site its own id to keep the two apart. Changing
the id points everyone at a fresh set of saved layouts — treat it as a deliberate reset. The value is
lower-cased and reduced to `a-z 0-9 . _ -`, max 50 characters, because it ends up in a list item
title, an OData filter and a `localStorage` key.

The list is created on first use, which needs Manage Lists rights — in practice a site owner opening
the page provisions it. It is created hidden, with `ReadSecurity`/`WriteSecurity` set to 2 so users
only ever see their own item, and `Title` is indexed: the list holds one item per user, and filtering
an unindexed column stops working once a list passes the 5,000-item list view threshold.

If the list cannot be read or written (read-only visitor, missing list, throttling), the store falls
back to `localStorage` and the dashboard shows a warning saying the arrangement is browser-local.
Swapping persistence later is a matter of implementing `ILayoutStore`.
