/**
 * The shared way of representing widget content.
 *
 * A widget describes what it wants to show (`IWidgetListItem`, `IWidgetEmptyState`,
 * `IWidgetDataState`) and hands it to a primitive here. It never styles a list, an
 * empty state or a failure itself, so every tile on the dashboard stays consistent
 * and a new widget is mostly data mapping.
 */
export * from './IWidgetContent';
export * from './toAdaptiveCard';
export * from './WidgetAdaptiveCard';
export * from './WidgetCards';
export * from './WidgetItems';
export * from './WidgetList';
export * from './WidgetProse';
export * from './WidgetSettings';
export * from './WidgetStat';
export * from './WidgetStates';
export * from './WidgetView';
export { toneClass } from './tone';
