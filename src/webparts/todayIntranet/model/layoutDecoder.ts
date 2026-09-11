import { IDashboardLayout } from './IDashboardLayout';

function invalid(message: string): never {
  throw new Error(`Invalid dashboard layout: ${message}`);
}

function isObject(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    return false;
  }
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function copyJson(value: unknown, ancestors: object[], depth: number, context?: 'layout' | 'widgets' | 'widget'): unknown {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== 'object' || !value || depth > 100 || ancestors.indexOf(value) >= 0) {
    return invalid('values must be finite, acyclic JSON data.');
  }
  if (!Array.isArray(value) && !isObject(value)) {
    return invalid('objects must be plain JSON objects.');
  }
  if (Object.getOwnPropertySymbols(value).length) {
    return invalid('symbol properties are not supported.');
  }
  const next = ancestors.concat([value]);
  if (Array.isArray(value)) {
    const result: unknown[] = [];
    if (Object.keys(value).length !== value.length) {
      return invalid('arrays must not contain holes or extra properties.');
    }
    for (let index = 0; index < value.length; index++) {
      const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
      if (!descriptor || !('value' in descriptor)) {
        return invalid('accessor properties are not supported.');
      }
      result.push(copyJson(descriptor.value, next, depth + 1, context === 'widgets' ? 'widget' : undefined));
    }
    return result;
  }
  const result: Record<string, unknown> = {};
  for (const key of Object.getOwnPropertyNames(value)) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !descriptor.enumerable || !('value' in descriptor)) {
      return invalid('properties must be enumerable JSON values.');
    }
    // Existing layout clones include undefined optional fields; omit these just
    // as JSON does, but reject undefined inside widget settings or unknown data.
    const optional = context === 'layout' && (key === 'presetId' || key === 'rowSizeId') ||
      context === 'widget' && (key === 'title' || key === 'settings');
    if (descriptor.value === undefined && optional) {
      continue;
    }
    Object.defineProperty(result, key, {
      value: copyJson(descriptor.value, next, depth + 1, context === 'layout' && key === 'widgets' ? 'widgets' : undefined),
      enumerable: true, writable: true, configurable: true
    });
  }
  return result;
}

function optionalString(object: Record<string, unknown>, key: string): void {
  if (object[key] !== undefined && typeof object[key] !== 'string') {
    invalid('optional title and preset fields must be strings.');
  }
}

function assertLayout(layout: unknown): asserts layout is Record<string, unknown> & IDashboardLayout {
  if (!isObject(layout)) {
    invalid('the root must be a plain object.');
  }
  if (layout.version !== 1 && layout.version !== 2) {
    return invalid('unsupported schema version; only versions 1 and 2 are supported.');
  }
  if (!Array.isArray(layout.widgets)) {
    return invalid('widgets must be an array.');
  }
  optionalString(layout, 'presetId');
  optionalString(layout, 'rowSizeId');
  const ids = new Set<string>();
  for (const widget of layout.widgets) {
    if (!isObject(widget)) {
      return invalid('each widget must be a plain object.');
    }
    if (typeof widget.id !== 'string' || !widget.id.trim() || ids.has(widget.id)) {
      return invalid('widget ids must be nonempty, unique strings.');
    }
    ids.add(widget.id);
    if (typeof widget.type !== 'string' || !widget.type.trim()) {
      return invalid('widget types must be nonempty strings.');
    }
    optionalString(widget, 'title');
    if (widget.settings !== undefined && !isObject(widget.settings)) {
      return invalid('widget settings must be a plain object.');
    }
    const { x, y, w, h } = widget;
    if (typeof x !== 'number' || typeof y !== 'number' || typeof w !== 'number' || typeof h !== 'number' ||
      !Number.isSafeInteger(x) || !Number.isSafeInteger(y) || !Number.isSafeInteger(w) || !Number.isSafeInteger(h) ||
      x < 0 || y < 0 || w < 1 || h < 1 || x + w > 12 || !Number.isSafeInteger(y + h)) {
      return invalid('widget geometry must use nonnegative integer positions and positive integer sizes within 12 columns.');
    }
  }
}

/** Validate persisted or edited layouts without depending on the widget registry. */
export function decodeLayout(value: unknown): IDashboardLayout {
  if (!isObject(value)) {
    return invalid('the root must be a plain object.');
  }
  const layout = copyJson(value, [], 0, 'layout');
  assertLayout(layout);
  return layout;
}

export function parseLayout(raw: string): IDashboardLayout {
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return invalid('the stored JSON could not be parsed.');
  }
  return decodeLayout(value);
}
