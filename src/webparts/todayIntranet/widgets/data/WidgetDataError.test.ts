import {
  classifyGraphDataError,
  isWidgetDataAuthorizationError,
  widgetDataErrorCodes,
  widgetDataErrorStatus
} from './WidgetDataError';

describe('WidgetDataError', () => {
  it.each([
    [{ statusCode: 401 }, true],
    [{ status: 403 }, true],
    [{ code: 'accessDenied' }, true],
    [{ innerError: { code: 'invalidAuthenticationToken' } }, true],
    [{ code: 'ServiceUnavailable' }, false]
  ])('classifies authorization errors without reading message text (%p)', (error, expected) => {
    expect(isWidgetDataAuthorizationError(error)).toBe(expected);
  });

  it('reads numeric status variants and recursively collects machine-readable codes', () => {
    const error = {
      status: 503,
      code: 'outer',
      innerError: { code: 'middle', innererror: { code: 'timeout' } }
    };

    expect(widgetDataErrorStatus(error)).toBe(503);
    expect(widgetDataErrorCodes(error)).toEqual(['outer', 'middle', 'timeout']);
  });

  it.each([
    [{ statusCode: 429 }, 'stale'],
    [{ status: 500 }, 'stale'],
    [{ code: 'serviceUnavailable' }, 'stale'],
    [{ innerError: { code: 'timeout' } }, 'stale'],
    [{ code: 'accessDenied' }, 'invalidate'],
    [new TypeError('mapping failed'), 'invalidate'],
    [{ code: 'badRequest' }, 'invalidate']
  ] as const)('uses stale data only for recognized Graph transient failures (%p)', (error, expected) => {
    expect(classifyGraphDataError(error)).toBe(expected);
  });
});
