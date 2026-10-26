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
});
