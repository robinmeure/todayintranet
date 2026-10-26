---
description: "Use when adding widgets, changing widget views, or editing dashboard styling. Covers shared typography, theme, responsive sizing, pagination, accessibility, and visual validation."
applyTo: "src/webparts/todayIntranet/widgets/**,src/webparts/todayIntranet/components/**,src/webparts/todayIntranet/styles/**"
---

# Widget look and feel

## Start with the shared presentation layer

Before implementing a widget, read the [content contract](../../README.md#how-a-widget-describes-what-it-shows)
and [visual style](../../README.md#visual-style). The README defines the approved hierarchy;
[widgetVisuals](../../src/webparts/todayIntranet/styles/_widgetVisuals.scss) implements its native tokens.
Keep both aligned when an approved design change affects the shared system.

- Map collections to `IWidgetListItem` and render through `WidgetItems`. This gives list,
  compact, cards, gallery, agenda, links, and Adaptive Card views the same presentation rules.
- Use `WidgetView` for loading, error, empty, refreshing, and stale-data states.
  Use `WidgetStat` or `WidgetProse` for single-content widgets.
- Register supported views and the default view in `WidgetRegistry`. `WidgetFrame` owns
  the widget heading, settings, actions, content surface, and intrinsic-height measurement.
- If a specialized interaction needs custom markup, reuse the shared Sass mixins and controls.
  Extend an existing primitive when the behavior is useful to other widgets.

## Typography and surfaces

- Keep Segoe UI and the shared font stack. Use `visual.item-title` for every collection item
  title, including custom tables; use `visual.text` for descriptions, `visual.meta` for
  dates and secondary information, and `visual.helper` for helper text and compact badge captions.
- Widget headings use `visual.widget-heading`; subsection headings use `visual.subsection-heading`.
  Keep the same item-title size, line height, and weight across widgets, views, and breakpoints.
  Narrow layouts wrap content rather than reducing title text.
- Preserve status meaning through existing tones, badges, and labels. Emphasis must not
  introduce another item-title font weight.
- Use shared spacing/radius tokens, semantic SPFx theme colors, `themed-border`, and `focus-ring`.
  Preserve dark-theme and forced-colors behavior. Use Fluent icons rather than text substitutes.
- Let the frame supply its rounded content surface. Use `contentSurface: 'none'` when a shared
  presentation already supplies its own surfaces, as My links does.
- Keep collection Adaptive Card typography aligned through
  [the host config](../../src/webparts/todayIntranet/widgets/content/adaptiveCardHostConfig.ts)
  and [item mapping](../../src/webparts/todayIntranet/widgets/content/toAdaptiveCard.ts).
  Authored cards retain explicit size choices; clock values and other non-item displays
  retain their distinct roles.

## Content sizing and pagination

- Keep all titles, descriptions, metadata, and actions visible. Wrap long text and URLs,
  use shrinkable grid/flex children (`min-width: 0`), and allow natural height growth.
  Widget content must not have internal scrollbars, ellipses, line clamps, or clipping.
  Normal vertical page scrolling and scrollable settings overlays remain allowed.
- Leave grid-height measurement to `WidgetFrame` and `Dashboard`. Measured heights are transient;
  only deliberate layout gestures persist geometry. Test growth and shrinkage without layout saves.
- Every collection must support optional per-instance pagination. Pass `paginationSize(context)`
  to `WidgetItems`; use `useWidgetPagination` and `WidgetPagination` for custom collections.
  Filter/sort before paging and keep every loaded matching item reachable.
- Use the frame's shared pagination settings unless the widget needs custom controls.
  Set `hasCustomPaginationSettings` only when `renderSettings` supplies the toggle and page-size
  control itself. Respect saved settings and document defaults, bounds, and remote fetch limits.
- Verify pagination-off behavior and reset/clamp paging after filters, ranges, or result counts
  change. Single-content widgets grow to fit rather than showing page controls.

## Completion checks

- Add or extend [typography contract tests](../../src/webparts/todayIntranet/widgets/content/widgetTypography.test.ts)
  for any new custom title surface. Cover pagination, empty/error/refresh states, and changed interactions.
- Validate with the repository's build, lint, and tests. For a deployable update, use the
  production build/package workflow in [package.json](../../package.json).
- Inspect computed styles and layout at 320, 375, 768, and 1440px with long titles,
  descriptions, URLs, empty data, and multiple pages. Confirm matching title typography,
  full text visibility, no horizontal page overflow, and no internal widget scrollbars.
- Check keyboard access, visible focus, accessible names for icon-only controls, and contrast
  in the existing theme. Keep essential actions available without hover.
- Distinguish local fixture verification from live SharePoint verification in the handoff;
  building a package does not deploy it or prove which version the tenant is running.
