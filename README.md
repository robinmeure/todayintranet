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
is on, so the web part is available on every site once the package is approved.

## How the dashboard fits together

```
TodayIntranetWebPart.ts          creates the layout store, renders <Dashboard>
components/Dashboard.tsx         grid, edit mode, add/remove, debounced persistence
components/WidgetFrame.tsx       shared tile chrome: icon, title, drag handle, remove
components/AddWidgetPanel.tsx    widget catalogue
widgets/IWidget.ts               the widget contract
widgets/WidgetRegistry.tsx       every widget the catalogue offers
model/IDashboardLayout.ts        persisted layout shape
services/*LayoutStore.ts         persistence
```

Layout is authored on a 12 column grid (`react-grid-layout`). Narrower breakpoints are derived from
that canonical layout by clamping widths, and rearranging is only saved at the widest breakpoint so a
phone visit can never scramble someone's desktop layout.

The web part sets `supportsFullBleed: true`, so place it in a **full-width section** on a
communication-site page to get edge-to-edge rendering.

## Adding a widget

1. Write a component that takes an `IWidgetContext` (instance id, its own `settings`, the SPFx
   `WebPartContext`, `isEditing`, and `updateSettings`).
2. Register an `IWidgetDefinition` for it in `widgets/WidgetRegistry.tsx` with a **stable** `type`
   key — that key is what lives in saved layouts, so never rename one that has shipped.

`demo.welcome` and `demo.clock` are placeholders that exercise both halves of the contract (persisted
per-instance settings, and live state that survives dragging). Replace them with the real calendar,
mail and tasks widgets; those will call Microsoft Graph via `context.msGraphClientFactory` and need
matching `webApiPermissionRequests` entries in `config/package-solution.json`.

## Where layouts are stored

Per user, in a hidden list named **TodayIntranetLayouts** on the site hosting the page: `Title` holds
`<web part instance id>|<user login name>` and a `LayoutJson` note column holds the layout.

The list is created on first use, which needs Manage Lists rights — in practice a site owner opening
the page provisions it. It is created hidden, with `ReadSecurity`/`WriteSecurity` set to 2 so users
only ever see their own item.

If the list cannot be read or written (read-only visitor, missing list, throttling), the store falls
back to `localStorage` and the dashboard shows a warning saying the arrangement is browser-local.
Swapping persistence later is a matter of implementing `ILayoutStore`.
