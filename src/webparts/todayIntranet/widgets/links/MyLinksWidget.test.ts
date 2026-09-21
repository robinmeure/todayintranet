import { DEFAULT_LINKS_SETTING, parseLinksSetting } from './linksSettings';

describe('My links settings', () => {
  it('provides valid, explicitly labelled samples with matching public destinations and no badges', () => {
    const parsed = parseLinksSetting(DEFAULT_LINKS_SETTING);
    expect(parsed.error).toBeUndefined();
    expect(parsed.links?.map(({ title, url }) => ({ title, url }))).toEqual([
      { title: 'Microsoft Learn', url: 'https://learn.microsoft.com/' },
      { title: 'Microsoft Support', url: 'https://support.microsoft.com/' },
      { title: 'MDN Web Docs', url: 'https://developer.mozilla.org/' },
      { title: 'GitHub Docs', url: 'https://docs.github.com/' },
      { title: 'Stack Overflow', url: 'https://stackoverflow.com/' },
      { title: 'Wikipedia', url: 'https://www.wikipedia.org/' }
    ]);
    parsed.links?.forEach((link) => {
      expect(link.description).toMatch(/^Sample: /);
      expect(link.badge).toBeUndefined();
    });
  });

  it('accepts valid HTTPS links and trims optional text', () => {
    expect(parseLinksSetting(JSON.stringify([{
      title: '  Learning Hub  ',
      url: 'https://learn.microsoft.com/training/',
      description: '  Training  ',
      iconName: 'Education',
      badge: 'VPN',
      tone: 'accent'
    }]))).toEqual({
      links: [{
        title: 'Learning Hub',
        url: 'https://learn.microsoft.com/training/',
        description: 'Training',
        iconName: 'Education',
        badge: 'VPN',
        tone: 'accent'
      }]
    });
  });

  it.each([
    ['not JSON', 'The links setting is not valid JSON.'],
    ['{}', 'The links setting must be a JSON array.'],
    ['[{"title":"Unsafe","url":"http://example.com"}]', 'Link 1 must use an absolute HTTPS URL without credentials.'],
    ['[{"title":"","url":"https://example.com"}]', 'Link 1 needs a title.'],
    ['[{"title":"Link","url":"https://user:secret@example.com"}]', 'Link 1 must use an absolute HTTPS URL without credentials.'],
    ['[{"title":"Link","url":"https://example.com","tone":"pink"}]', 'Link 1 has an invalid tone.']
  ])('rejects invalid configuration %#', (json, message) => {
    expect(parseLinksSetting(json)).toEqual({ error: message });
  });

  it('accepts an intentionally empty configuration', () => {
    expect(parseLinksSetting('[]')).toEqual({ links: [] });
  });

  // eslint-disable-next-line no-script-url -- The validator must reject script URLs.
  it.each(['javascript:alert(1)', 'data:text/html,test', '/relative/path', '//example.test', 'mailto:test@example.test'])(
    'rejects non-HTTPS destinations (%s)', (url) => {
      expect(parseLinksSetting(JSON.stringify([{ title: 'Link', url }])).error).toBeDefined();
    }
  );

  it.each(['neutral', 'accent', 'success', 'warning', 'danger'])('accepts the shared %s tone', (tone) => {
    expect(parseLinksSetting(JSON.stringify([{ title: 'Link', url: 'https://example.test', tone }])).links?.[0].tone)
      .toBe(tone);
  });

  it.each([['blue', 'accent'], ['purple', 'accent'], ['green', 'success'], ['orange', 'warning'], ['red', 'danger']])(
    'migrates legacy %s tones to %s', (tone, expected) => {
      expect(parseLinksSetting(JSON.stringify([{ title: 'Link', url: 'https://example.test', tone }])).links?.[0].tone)
        .toBe(expected);
    }
  );

  it('enforces the configured item limit without truncating silently', () => {
    const links = Array.from({ length: 100 }, () => ({ title: 'Link', url: 'https://example.test' }));
    expect(parseLinksSetting(JSON.stringify(links)).links).toHaveLength(100);
    links.push(links[0]);
    expect(parseLinksSetting(JSON.stringify(links)).error).toContain('at most 100');
  });

  it.each([null, 123, [], { title: 'Link', url: 'https://example.test', badge: true }])(
    'reports malformed entries (%p)', (entry) => {
      expect(parseLinksSetting(JSON.stringify([entry])).error).toBeDefined();
    }
  );
});
