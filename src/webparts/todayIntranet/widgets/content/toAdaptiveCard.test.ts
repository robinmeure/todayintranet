import { itemsToAdaptiveCard } from './toAdaptiveCard';

describe('collection Adaptive Card typography', () => {
  it.each([false, true])('uses the same bold title for emphasized=%s', (isEmphasized) => {
    expect(itemsToAdaptiveCard([{
      title: 'Meeting title', meta: ['Tomorrow'], description: 'Meeting description',
      isEmphasized, tone: 'accent', href: 'https://example.com/meeting'
    }])).toMatchObject({
      body: [{
        selectAction: { type: 'Action.OpenUrl', url: 'https://example.com/meeting' },
        items: [{
          columns: [{
            items: [
              { text: 'Meeting title', weight: 'Bolder', color: 'Accent', wrap: true },
              { text: 'Tomorrow', size: 'Small', wrap: true },
              { text: 'Meeting description', size: 'Default', wrap: true }
            ]
          }]
        }]
      }]
    });
  });
});
