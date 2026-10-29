export {};

const { readFileSync } = jest.requireActual<{ readFileSync(path: string, encoding: 'utf8'): string }>('fs');
const { join } = jest.requireActual<{ join(...paths: string[]): string }>('path');
declare const __dirname: string;

function declarations(file: string, selector: string, property: string): string[] {
  const css = readFileSync(join(__dirname, file), 'utf8');
  const values: string[] = [];
  const rules = /([^{}]+)\{([^{}]*)\}/g;
  let rule: RegExpExecArray | null;
  while ((rule = rules.exec(css))) {
    if (rule[1].split(',').some((part) => part.trim() === selector)) {
      rule[2].split(';').forEach((declaration) => {
        const colon = declaration.indexOf(':');
        if (declaration.slice(0, colon).trim() === property) {
          values.push(declaration.slice(colon + 1).trim());
        }
      });
    }
  }
  return values;
}

describe('compiled widget typography', () => {
  it('hides the entire Tags column only in the small task layout', () => {
    const file = '../graph/TasksWidget.module.scss.css';
    const css = readFileSync(join(__dirname, file), 'utf8');
    const smallLayout = css.split('@container (max-width: 540px) {')[1]?.split('@container')[0];
    expect(smallLayout).toBeDefined();
    expect(smallLayout).toMatch(/\.table \.tags,\s*\.table \.due\s*\{\s*display: none;/);
    expect(declarations(file, '.table .tags', 'display')).toEqual(['none']);
  });

  it.each([
    ['WidgetCollections.module.scss.css', '.eventTitle'],
    ['WidgetCollections.module.scss.css', '.linkTitle'],
    ['WidgetContent.module.scss.css', '.itemTitle'],
    ['WidgetContent.module.scss.css', '.cardTitle'],
    ['../graph/TasksWidget.module.scss.css', '.title'],
    ['../graph/VivaEngageWidget.module.scss.css', '.name']
  ])('keeps %s %s at the shared title size, including responsive overrides', (file, selector) => {
    expect(declarations(file, selector, 'font-size')).toEqual(['16px']);
    expect(declarations(file, selector, 'line-height')).toEqual(['24px']);
    expect(declarations(file, selector, 'font-weight')).toEqual(['700']);
    expect(declarations(file, selector, 'font-family')[0]).toContain('"Segoe UI"');
  });

  it.each(['.title', '.settingsTitle'])('uses the reference hierarchy for %s', (selector) => {
    const file = '../../components/WidgetFrame.module.scss.css';
    expect(declarations(file, selector, 'font-size').slice(-1)).toEqual(['28px']);
    expect(declarations(file, selector, 'line-height').slice(-1)).toEqual(['28.5px']);
    expect(declarations(file, selector, 'font-weight')).toEqual(['400']);
  });

  it.each([
    ['WidgetContent.module.scss.css', '.itemMeta', '14px', '20px'],
    ['WidgetContent.module.scss.css', '.cardMeta', '14px', '20px'],
    ['WidgetCollections.module.scss.css', '.eventMeta', '14px', '20px'],
    ['WidgetContent.module.scss.css', '.itemDescription', '16px', '24px'],
    ['WidgetContent.module.scss.css', '.cardDescription', '16px', '24px'],
    ['WidgetCollections.module.scss.css', '.linkDescription', '16px', '24px'],
    ['WidgetCollections.module.scss.css', '.eventDescription', '16px', '24px'],
    ['WidgetContent.module.scss.css', '.settingsDescription', '12px', '16px']
  ])('keeps %s %s in its text role', (file, selector, size, height) => {
    expect(declarations(file, selector, 'font-size')).toEqual([size]);
    expect(declarations(file, selector, 'line-height')).toEqual([height]);
  });

  it.each(['.itemTitleEmphasized', '.cardTitleEmphasized'])('preserves emphasis without changing title weight (%s)', (selector) => {
    expect(declarations('WidgetContent.module.scss.css', selector, 'font-weight')).toEqual([]);
    expect(declarations('WidgetContent.module.scss.css', selector, 'color')).toEqual(['currentColor']);
  });

  it('uses wrapping, stretchable link cards with the approved dimensions', () => {
    const file = 'WidgetCollections.module.scss.css';
    expect(declarations(file, '.linksGrid', 'display')).toEqual(['flex']);
    expect(declarations(file, '.linksGrid', 'flex-wrap')).toEqual(['wrap']);
    expect(declarations(file, '.linksGrid', 'max-width')).toEqual(['1164px']);
    expect(declarations(file, '.linksGrid > li', 'flex')).toEqual(['1 0 280px']);
    expect(declarations(file, '.linkCard', 'display')).toEqual(['flex']);
    expect(declarations(file, '.linkCard', 'flex')).toEqual(['1 0 0']);
    expect(declarations(file, '.linkCard', 'min-width')).toEqual(['min(280px, 100%)']);
    expect(declarations(file, '.linkCard', 'height')).toEqual(['112px']);
    expect(declarations(file, '.linkCard', 'min-height')).toEqual(['112px']);
    expect(declarations(file, '.linkCard', 'padding')).toEqual(['16px 24px 16px 16px']);
    expect(declarations(file, '.linkCard', 'gap')).toEqual(['16px']);
  });

  it.each(['.linkTitle', '.linkDescription'])('clamps %s visually to one line only in link tiles', (selector) => {
    const file = 'WidgetCollections.module.scss.css';
    expect(declarations(file, selector, 'display')).toEqual(['-webkit-box']);
    expect(declarations(file, selector, '-webkit-box-orient')).toEqual(['vertical']);
    expect(declarations(file, selector, '-webkit-line-clamp')).toEqual(['1']);
    expect(declarations(file, selector, 'align-self')).toEqual(['stretch']);
    expect(declarations(file, selector, 'overflow')).toEqual(['hidden']);
    expect(declarations(file, '.eventTitle', '-webkit-line-clamp')).toEqual([]);
    expect(declarations(file, '.eventDescription', '-webkit-line-clamp')).toEqual([]);
  });
});
