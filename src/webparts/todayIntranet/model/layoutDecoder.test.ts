import { cloneLayout, IDashboardLayout } from './IDashboardLayout';
import { decodeLayout, parseLayout } from './layoutDecoder';

const layout: IDashboardLayout = {
  version: 2, presetId: 'custom-preset', rowSizeId: 'custom-height',
  widgets: [{ id: 'one', type: 'unregistered.widget', title: '', x: 0, y: 0, w: 12, h: 4,
    settings: { nested: { items: [null, true, 2, 'value'] } } }]
};

describe('layoutDecoder', () => {
  it.each([1, 2])('round trips version %s, unknown widget types and optional fields', (version) => {
    const value = { ...layout, version, futureOption: { enabled: true },
      widgets: [{ ...layout.widgets[0], futureWidgetOption: 'preserve' }] };
    expect(parseLayout(JSON.stringify(value))).toEqual(value);
  });

  it('deep clones settings and accepts absent optional fields from existing clones', () => {
    const decoded = decodeLayout(layout);
    expect(decoded).toEqual(layout);
    expect(decoded.widgets[0].settings?.nested).not.toBe(layout.widgets[0].settings?.nested);
    expect(decodeLayout(cloneLayout({ version: 1, widgets: [{ id: 'a', type: 'unknown', x: 0, y: 0, w: 1, h: 1 }] })))
      .toEqual({ version: 1, widgets: [{ id: 'a', type: 'unknown', x: 0, y: 0, w: 1, h: 1 }] });
  });

  it.each([null, [], 'data', {}, { version: 3, widgets: [] }, { version: 0, widgets: [] },
    { version: 2, widgets: {} }, { ...layout, presetId: 2 }, { ...layout, rowSizeId: null }])
  ('rejects invalid roots and unsupported versions (%j)', (value) => {
    expect(() => decodeLayout(value)).toThrow('Invalid dashboard layout:');
  });

  it.each([
    { id: '' }, { id: '  ' }, { type: '' }, { title: 3 }, { settings: [] }, { settings: null },
    { x: -1 }, { x: 1 }, { x: 0.5 }, { y: -1 }, { y: Infinity }, { w: 0 }, { w: 13 }, { h: 0 },
    { h: 1.5 }, { y: Number.MAX_SAFE_INTEGER }, { settings: new Date() }
  ])('rejects invalid widget fields (%j)', (fields) => {
    expect(() => decodeLayout({ ...layout, widgets: [{ ...layout.widgets[0], ...fields }] })).toThrow();
  });

  it('requires unique ids and does not leak identifiers in error messages', () => {
    expect(() => decodeLayout({ ...layout, widgets: [layout.widgets[0], layout.widgets[0]] }))
      .toThrow('widget ids must be nonempty, unique strings');
    expect(() => parseLayout('secret-personal-content')).toThrow('the stored JSON could not be parsed');
  });

  it.each([undefined, NaN, Infinity, () => undefined, Symbol('private')])
  ('rejects non-JSON settings values', (value) => {
    expect(() => decodeLayout({ ...layout, widgets: [{ ...layout.widgets[0], settings: { value } }] })).toThrow();
  });

  it('rejects cycles, sparse arrays, accessors and undefined unknown fields', () => {
    const cyclic: Record<string, unknown> = {};
    cyclic.self = cyclic;
    const accessor = Object.defineProperty({}, 'value', { enumerable: true, get: () => { throw new Error('private'); } });
    [cyclic, { list: new Array(2) }, accessor, { title: undefined }].forEach((settings) =>
      expect(() => decodeLayout({ ...layout, widgets: [{ ...layout.widgets[0], settings }] })).toThrow());
    expect(() => decodeLayout({ ...layout, extra: { title: undefined } })).toThrow();
  });

  it('preserves JSON __proto__ data without changing object prototypes', () => {
    const decoded = parseLayout('{"version":2,"widgets":[],"__proto__":{"polluted":true}}');
    expect(Object.getPrototypeOf(decoded)).toBe(Object.prototype);
    expect(Object.prototype.hasOwnProperty.call(decoded, '__proto__')).toBe(true);
  });
});
