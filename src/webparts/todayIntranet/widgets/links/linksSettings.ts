import { WidgetTone } from '../content/IWidgetContent';

export interface IMyLink {
  title: string;
  url: string;
  description?: string;
  iconName?: string;
  badge?: string;
  tone?: WidgetTone;
}

const SAMPLE_LINKS: IMyLink[] = [
  {
    title: 'Microsoft Learn',
    url: 'https://learn.microsoft.com/',
    description: 'Sample: training and documentation.',
    iconName: 'Education',
    tone: 'accent'
  },
  {
    title: 'Microsoft Support',
    url: 'https://support.microsoft.com/',
    description: 'Sample: product help and guidance.',
    iconName: 'Help',
    tone: 'success'
  },
  {
    title: 'MDN Web Docs',
    url: 'https://developer.mozilla.org/',
    description: 'Sample: web development references.',
    iconName: 'Code',
    tone: 'warning'
  },
  {
    title: 'GitHub Docs',
    url: 'https://docs.github.com/',
    description: 'Sample: guides for working with GitHub.',
    iconName: 'Documentation',
    tone: 'accent'
  },
  {
    title: 'Stack Overflow',
    url: 'https://stackoverflow.com/',
    description: 'Sample: programming questions and answers.',
    iconName: 'Chat',
    tone: 'warning'
  },
  {
    title: 'Wikipedia',
    url: 'https://www.wikipedia.org/',
    description: 'Sample: explore the free encyclopedia.',
    iconName: 'Globe',
    tone: 'success'
  }
];

export const DEFAULT_LINKS_SETTING: string = JSON.stringify(SAMPLE_LINKS, undefined, 2);

export type LinksSettingResult = { links: IMyLink[]; error?: undefined } | { error: string; links?: undefined };

function optionalString(item: Record<string, unknown>, key: keyof IMyLink, max: number, index: number): string | undefined {
  const value = item[key];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== 'string' || value.length > max) {
    throw new Error(`Link ${index} has an invalid ${key}.`);
  }
  return value.trim() || undefined;
}

function readTone(value: string | undefined, index: number): WidgetTone | undefined {
  switch (value) {
    case undefined:
    case 'neutral':
    case 'accent':
    case 'success':
    case 'warning':
    case 'danger':
      return value;
    // Preserve settings saved by the initial links implementation.
    case 'blue':
    case 'purple': return 'accent';
    case 'green': return 'success';
    case 'orange': return 'warning';
    case 'red': return 'danger';
    default: throw new Error(`Link ${index} has an invalid tone.`);
  }
}

function safeHttpsUrl(value: unknown, index: number): string {
  if (typeof value !== 'string' || value.length > 2048) {
    throw new Error(`Link ${index} has an invalid URL.`);
  }
  let parsed: URL;
  try {
    parsed = new URL(value);
  } catch {
    throw new Error(`Link ${index} must use an absolute HTTPS URL.`);
  }
  if (parsed.protocol !== 'https:' || parsed.username || parsed.password) {
    throw new Error(`Link ${index} must use an absolute HTTPS URL without credentials.`);
  }
  return parsed.toString();
}

export function parseLinksSetting(json: string): LinksSettingResult {
  let value: unknown;
  try {
    value = JSON.parse(json);
  } catch {
    return { error: 'The links setting is not valid JSON.' };
  }
  if (!Array.isArray(value)) {
    return { error: 'The links setting must be a JSON array.' };
  }
  if (value.length > 100) {
    return { error: 'The links setting can contain at most 100 links.' };
  }
  try {
    return {
      links: value.map((item: unknown, index: number) => {
        const number = index + 1;
        if (!item || typeof item !== 'object' || Array.isArray(item)) {
          throw new Error(`Link ${number} must be a JSON object.`);
        }
        const record = item as Record<string, unknown>;
        const title = optionalString(record, 'title', 120, number);
        if (!title) {
          throw new Error(`Link ${number} needs a title.`);
        }
        return {
          title,
          url: safeHttpsUrl(record.url, number),
          description: optionalString(record, 'description', 240, number),
          iconName: optionalString(record, 'iconName', 64, number),
          badge: optionalString(record, 'badge', 40, number),
          tone: readTone(optionalString(record, 'tone', 10, number), number)
        };
      })
    };
  } catch (error) {
    return { error: error instanceof Error ? error.message : 'The links setting is invalid.' };
  }
}
