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
components/WidgetFrame.tsx        tile chrome: icon, title, drag handle, refresh, settings flyout, footer link
components/WidgetErrorBoundary.tsx contains a crashing widget
components/AddWidgetPanel.tsx     widget catalogue: search, categories, permission hints
components/LayoutPresetPanel.tsx  layout picker: column presets and tile height
widgets/IWidget.ts                the widget contract
widgets/WidgetRegistry.tsx        every widget the catalogue offers
widgets/content/                  how widgets describe what they show (see below)
widgets/graph/useGraphData.ts     Graph query hook with consent-aware error handling
widgets/search/                   SharePoint search: service, hook and the query builder UI
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

## How a widget describes what it shows

Widgets do not style themselves. They describe their content with the shapes in
`widgets/content` and hand it to a shared primitive, which is what keeps a calendar tile, a mail
tile and anything added later looking like one product instead of five.

```
widgets/content/IWidgetContent.ts  the vocabulary: tones, list items, badges, errors, data state, views
widgets/content/WidgetView.tsx     one data state -> skeleton, error, empty state or content
widgets/content/WidgetItems.tsx    draws a collection in whichever view the tile is set to
widgets/content/WidgetList.tsx     the list view, comfortable or compact
widgets/content/WidgetCards.tsx    the cards and gallery views
widgets/content/WidgetAdaptiveCard.tsx  the Adaptive Card renderer, themed from the site
widgets/content/toAdaptiveCard.ts  turns ordinary widget items into an Adaptive Card payload
widgets/content/WidgetStat.tsx     a single headline figure
widgets/content/WidgetProse.tsx    heading plus body text
widgets/content/WidgetStates.tsx   the skeleton, empty and error presentations
widgets/content/WidgetSettings.tsx typed fields for the settings flyout, and the view picker
```

A data-driven widget is then a query plus a mapping:

```tsx
const state = useGraphData<Message[]>(context, 'Mail.ReadBasic', fetchMessages, [maxItems]);
const view = viewSetting(context, MAIL_DEFAULT_VIEW);

return (
  <WidgetView
    state={state}
    loading={{ label: 'Loading your inbox…', rows: 4 }}
    empty={{ iconName: 'Mail', text: 'Your inbox is empty.' }}
  >
    {(messages) => <WidgetItems items={messages.map(toItem)} view={view} />}
  </WidgetView>
);
```

`WidgetView` owns the four states every asynchronous tile has, so no widget re-implements them:

| State | What the user sees |
| --- | --- |
| Loading | A skeleton shaped like the rows that are coming, not a spinner |
| Refreshing | The data it already had, with a progress bar — a refresh never blanks a tile |
| Empty | An icon, a sentence and an optional way out ("Open Outlook") |
| Error | A message bar, with **Try again** unless an administrator has to act first |

`IWidgetListItem` carries a **tone** (`neutral`, `accent`, `success`, `warning`, `danger`) rather
than a colour, so emphasis means the same thing on every tile: a meeting in progress, an overdue
task and an unread message all lean on the same five tones and the same badge.

### Views

Because items are described rather than styled, the same data can be drawn several ways. A widget
declares `supportedViews`, the tile gets a **View** picker in its settings flyout, and the choice is
saved per tile — one person's inbox can be a compact list while another's is a gallery.

| View | What it draws |
| --- | --- |
| List | One row per item — icon, title, meta line |
| Compact | The same rows, tightened, so more fits in a short tile |
| Cards | Each item on its own card, with a tone-coloured edge |
| Gallery | Those cards flowed into a responsive grid |
| Adaptive card | The items rendered as a real Adaptive Card, themed from the site |

Widgets never choose a view themselves — they read `viewSetting(context, defaultView)` and hand the
items to `WidgetItems`, which does the rest.

#### The same items, five ways

Three calendar events mapped onto `IWidgetListItem` — one meeting running right now, one later
today, one tomorrow:

```tsx
const items: IWidgetListItem[] = [
  {
    key: 'AAMkAGI3…',
    title: 'Weekly leadership sync',
    meta: ['Today · 09:30–10:00', 'Microsoft Teams'],
    hasAccentBar: true,
    tone: 'success',
    isEmphasized: true,
    badge: { text: 'Now' },
    href: 'https://outlook.office.com/calendar/item/AAMkAGI3…'
  },
  {
    key: 'AAMkAGI4…',
    title: 'Intranet design review',
    meta: ['Today · 14:00–15:00', 'Room 4.12'],
    hasAccentBar: true,
    tone: 'accent',
    isEmphasized: true,
    href: 'https://outlook.office.com/calendar/item/AAMkAGI4…'
  },
  {
    key: 'AAMkAGI5…',
    title: 'Quarterly all hands',
    meta: ['Tomorrow · All day'],
    hasAccentBar: true,
    tone: 'neutral',
    href: 'https://outlook.office.com/calendar/item/AAMkAGI5…'
  }
];
```

Nothing in there says "blue", "row" or "card". The view decides that, and only that:

```
List                                        Compact
┌──────────────────────────────────────┐    ┌──────────────────────────────────────┐
│ ▌ Weekly leadership sync       Now   │    │ ▌ Weekly leadership sync       Now   │
│ ▌ Today · 09:30–10:00 · Teams        │    │ ▌ Today · 09:30–10:00 · Teams        │
│                                      │    │ ▌ Intranet design review             │
│ ▌ Intranet design review             │    │ ▌ Today · 14:00–15:00 · Room 4.12    │
│ ▌ Today · 14:00–15:00 · Room 4.12    │    │ ▌ Quarterly all hands                │
│                                      │    │ ▌ Tomorrow · All day                 │
│ ▌ Quarterly all hands                │    └──────────────────────────────────────┘
│ ▌ Tomorrow · All day                 │
└──────────────────────────────────────┘

Cards                                       Gallery
┌──────────────────────────────────────┐    ┌───────────────┐ ┌───────────────┐
│ ┌──────────────────────────────────┐ │    │ ▌       Now   │ │ ▌             │
│ │▌                          Now    │ │    │ ▌ Weekly lea… │ │ ▌ Intranet d… │
│ │▌ Weekly leadership sync          │ │    │ ▌ Today · 09… │ │ ▌ Today · 14… │
│ │▌ Today · 09:30–10:00 · Teams     │ │    └───────────────┘ └───────────────┘
│ └──────────────────────────────────┘ │    ┌───────────────┐
│ ┌──────────────────────────────────┐ │    │ ▌             │
│ │▌ Intranet design review          │ │    │ ▌ Quarterly … │
│ │▌ Today · 14:00–15:00 · Room 4.12 │ │    │ ▌ Tomorrow ·… │
│ └──────────────────────────────────┘ │    └───────────────┘
└──────────────────────────────────────┘
```

The **Adaptive card** view runs the same items through `itemsToAdaptiveCard`, which is where the
tones become Adaptive Card colours. The first item above becomes:

```jsonc
{
  "type": "Container",
  "spacing": "None",
  "separator": false,
  "selectAction": {
    "type": "Action.OpenUrl",
    "title": "Weekly leadership sync",
    "url": "https://outlook.office.com/calendar/item/AAMkAGI3…"
  },
  "items": [
    {
      "type": "ColumnSet",
      "spacing": "None",
      "columns": [
        {
          "type": "Column",
          "width": "stretch",
          "verticalContentAlignment": "Center",
          "items": [
            // tone: 'success' + isEmphasized
            { "type": "TextBlock", "text": "Weekly leadership sync",
              "weight": "Bolder", "color": "Good", "wrap": true, "spacing": "None" },
            { "type": "TextBlock", "text": "Today · 09:30–10:00 · Microsoft Teams",
              "size": "Small", "isSubtle": true, "wrap": false, "spacing": "None" }
          ]
        },
        {
          "type": "Column",
          "width": "auto",
          "verticalContentAlignment": "Center",
          "items": [
            { "type": "TextBlock", "text": "Now", "size": "Small", "weight": "Bolder",
              "color": "Good", "wrap": false, "spacing": "None" }
          ]
        }
      ]
    }
  ]
}
```

#### Tones

One tone vocabulary drives all five views, so the same meaning gets the same emphasis everywhere:

| Tone | Where it comes from | In list / cards | In an Adaptive Card |
| --- | --- | --- | --- |
| `neutral` | An event later in the week, a read message | Grey bar and icon | `Default` |
| `accent` | Unread mail, an event today | Theme colour | `Accent` |
| `success` | A meeting happening right now | Green | `Good` |
| `warning` | A high-importance task | Amber | `Warning` |
| `danger` | An overdue task | Red, plus an "Overdue" badge | `Attention` |

### Adaptive Cards

`WidgetAdaptiveCard` renders any Adaptive Card (schema 1.5 and lower) with the `adaptivecards`
package. Three things are worth knowing:

- **It is lazy loaded.** The renderer is a large dependency, so it is a separate webpack chunk that
  is only fetched when a tile actually shows a card.
- **It is themed.** The host config is built from the current Fluent/SPFx theme, so cards pick up the
  site's colours and fonts, and our tones map onto the Adaptive Card text colours (`accent` →
  `Accent`, `danger` → `Attention`, and so on).
- **Text is plain text.** Markdown processing is explicitly disabled, so a payload can never inject
  HTML — `Action.OpenUrl` opens in a new tab, `Action.Submit` is handed to the widget, and nothing
  else is honoured.

The **Adaptive card** widget in the catalogue renders a payload pasted into its settings, which is
also the reference for a widget that wants to hand over a card instead of items. An announcement
tile, for instance:

```json
{
  "type": "AdaptiveCard",
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "version": "1.5",
  "body": [
    { "type": "TextBlock", "text": "Office move: 14 October", "weight": "Bolder", "size": "Medium", "wrap": true },
    { "type": "TextBlock", "text": "Floors 3 and 4 move to the north wing. Pack your desk by Friday.", "wrap": true, "isSubtle": true },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Owner", "value": "Facilities" },
        { "title": "Questions", "value": "#office-move" }
      ]
    }
  ],
  "actions": [
    { "type": "Action.OpenUrl", "title": "Read the plan", "url": "https://contoso.sharepoint.com/sites/facilities/move" }
  ]
}
```

The payload is stored in that tile's settings, so it travels with the person's dashboard rather than
being published to everyone — handy for a personal card, and the reason a widget backed by a list or
an API is the better route for something the whole site should see.

### Tile chrome

`WidgetFrame` wraps every widget and owns everything outside the content:

- **Refresh** — widgets that set `isRefreshable` get a refresh button; pressing it bumps
  `IWidgetContext.refreshToken`, which `useGraphData` treats as a dependency. The frame never has to
  know where the data came from.
- **Quiet actions** — tile buttons fade in on hover or keyboard focus, and stay visible in edit mode
  and on touch screens, so a full dashboard does not read as a wall of icons.
- **Footer link** — a definition's `footerLink` becomes the tile's "Open calendar" style link.
- Each tile is a labelled region, so screen reader users can jump between widgets.

### Search widgets

`widgets/search` is a second data source alongside Graph, and the same
`IWidgetDataState` shape, so search widgets get skeletons, refresh, empty states and errors for free:

```
widgets/search/SearchService.ts       the REST call, row flattening and error mapping
widgets/search/useSearchData.ts       the search equivalent of useGraphData
widgets/search/searchQuery.ts         scopes, and how a query is assembled from them
widgets/search/SearchQuerySetting.tsx scope picker, query box and the live preview
```

A query is assembled from three parts, and the tile shows the result of that assembly:

| Part | Comes from | Example |
| --- | --- | --- |
| Base | The widget, fixed | `PromotedState:2` (news posts) |
| Scope | The **Look in** picker | `SPSiteURL:"https://contoso.sharepoint.com/sites/hr"` |
| Terms | Whatever the user typed | `FileType:docx handbook` |

| Scope | KQL it adds |
| --- | --- |
| This site | `SPWebUrl:"<web url>"` |
| This site collection | `SPSiteURL:"<site url>"` |
| This hub | `DepartmentId:{<hub id>}` — disabled when the site is not in a hub |
| Everywhere | nothing |

**The preview is the point.** Writing KQL blind is the worst part of configuring a search widget, so
the settings flyout shows the effective query and runs it, debounced, as it is edited:

```
Look in   [ This site collection      ▾ ]
Query     [ FileType:aspx               ]

┌─ PREVIEW ─────────────────────────── Run again ─┐
│ SPSiteURL:"https://contoso.sharepoint.com/      │
│ sites/hr" FileType:aspx                         │
│                                                 │
│ 11 results, first 5 shown                       │
│ · Employee handbook — HR                        │
│ · Onboarding checklist — HR                     │
│ · Expenses policy — HR                          │
└─────────────────────────────────────────────────┘
```

The preview runs exactly what the tile will run — including refusing to run when the widget would.
The Search results widget needs terms, so a scope on its own reports that rather than showing a
sample of the site the tile would never display.

## Widgets

| Type key | Widget | Graph permission |
| --- | --- | --- |
| `m365.calendar` | Calendar — upcoming events, configurable days ahead | `Calendars.ReadBasic` |
| `m365.mail` | My mail — inbox, optional unread-only | `Mail.ReadBasic` |
| `m365.tasks` | My tasks — open Microsoft To Do items | `Tasks.Read` |
| `sp.news` | News — news posts from a site, a hub, or everywhere | — |
| `sp.search` | Search results — any SharePoint search query | — |
| `card.adaptive` | Adaptive card — renders a payload from the tile's settings | — |
| `demo.clock` | Clock | — |
| `demo.welcome` | Welcome greeting (reference implementation) | — |

**News** and **Search results** run on the SharePoint search REST API as the signed in user, so they
need no admin consent and no permission request: results are already trimmed to what that person can
open. See [Search widgets](#search-widgets).

The three Microsoft 365 widgets need tenant admin approval of the `webApiPermissionRequests` in
`config/package-solution.json`, granted in **SharePoint admin center > Advanced > API access** after
the package is deployed. Until that happens each widget shows a message naming the missing
permission rather than failing silently. The `.ReadBasic` scopes are deliberate: SPFx grants
permissions to the tenant-wide SharePoint Online Client Extensibility principal, not to this
solution alone, so the widgets ask only for the metadata they render — never message or event
bodies and attachments.

### Adding a widget

1. Write a component that takes an `IWidgetContext` (instance id, its own `settings`, the SPFx
   `WebPartContext`, `isEditing`, `refreshToken`, and `updateSettings`). Use `useGraphData` for
   anything that calls Microsoft 365 — it handles loading, cancellation, throttling, missing consent
   and the refresh token for you.
2. Describe the content with the primitives in `widgets/content` rather than styling it yourself:
   `WidgetView` for the loading/error/empty/content states, then `WidgetItems` for a collection (it
   handles all five views), or `WidgetStat`, `WidgetProse` or `WidgetAdaptiveCard` for something
   else. A widget is usually a query plus a `toItem` mapping.
3. Register an `IWidgetDefinition` for it in `widgets/WidgetRegistry.tsx` with a **stable** `type`
   key — that key is what lives in saved layouts, so never rename one that has shipped. Give it a
   `category` and `keywords` so it can be found in the catalogue, `requiredPermission` if it needs
   Graph consent, `isRefreshable` if it re-reads on `refreshToken`, `supportedViews` and
   `defaultView` if it renders a collection, and a `footerLink` if there is a full experience to
   link out to.
4. Optionally implement `renderSettings`, using the typed fields in `widgets/content` — that is what
   puts a gear icon and settings flyout on the tile.

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
