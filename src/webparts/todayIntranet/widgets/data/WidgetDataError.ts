export type WidgetDataErrorDisposition = 'stale' | 'invalidate';

interface IErrorShape {
  status?: unknown;
  statusCode?: unknown;
  code?: unknown;
  innerError?: unknown;
  innererror?: unknown;
}

const AUTHORIZATION_CODES = new Set<string>([
  'accessdenied',
  'authorization_requestdenied',
  'authenticationnotconfigured',
  'invalidauthenticationtoken',
  'insufficient_claims',
  'unauthenticated'
]);

const TRANSIENT_CODES = new Set<string>([
  'activitylimitreached',
  'extensionerror',
  'gatewaytimeout',
  'serviceunavailable',
  'timeout',
  'toomanyrequests',
  'transienterror'
]);

function errorShape(error: unknown): IErrorShape | undefined {
  return error && typeof error === 'object' ? error as IErrorShape : undefined;
}

export function widgetDataErrorStatus(error: unknown): number | undefined {
  const shape = errorShape(error);
  const status = shape?.statusCode ?? shape?.status;
  return typeof status === 'number' ? status : undefined;
}

export function widgetDataErrorCodes(error: unknown): string[] {
  const result: string[] = [];
  let current = errorShape(error);
  const visited = new Set<IErrorShape>();

  while (current && !visited.has(current)) {
    visited.add(current);
    if (typeof current.code === 'string') {
      result.push(current.code.toLowerCase());
    }
    current = errorShape(current.innerError ?? current.innererror);
  }

  return result;
}

export function isWidgetDataAuthorizationError(error: unknown): boolean {
  const status = widgetDataErrorStatus(error);
  return status === 401 ||
    status === 403 ||
    widgetDataErrorCodes(error).some((code) => AUTHORIZATION_CODES.has(code));
}

export function classifyGraphDataError(error: unknown): WidgetDataErrorDisposition {
  if (isWidgetDataAuthorizationError(error)) {
    return 'invalidate';
  }

  const status = widgetDataErrorStatus(error);
  if (status === 429 || status === 500 || status === 502 || status === 503 || status === 504) {
    return 'stale';
  }

  return widgetDataErrorCodes(error).some((code) => TRANSIENT_CODES.has(code))
    ? 'stale'
    : 'invalidate';
}
