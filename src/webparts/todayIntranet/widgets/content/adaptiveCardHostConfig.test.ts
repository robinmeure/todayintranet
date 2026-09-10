import { createTheme } from '@fluentui/react/lib/Theme';
import { toAdaptiveCardHostConfig } from './adaptiveCardHostConfig';

describe('toAdaptiveCardHostConfig', () => {
  it('uses the shared text hierarchy without adding a second layer of padding', () => {
    const theme = createTheme();

    expect(toAdaptiveCardHostConfig(theme)).toMatchObject({
      fontFamily: theme.fonts.medium.fontFamily,
      fontSizes: { small: 12, default: 14 },
      lineHeights: { small: 16, default: 20 },
      spacing: { small: 4, default: 8, medium: 12, large: 16, padding: 0 }
    });
  });

  it('uses semantic status foregrounds rather than fixed light-theme palette colors', () => {
    const theme = createTheme({
      semanticColors: {
        successText: '#185c18',
        warningText: '#5a4300',
        errorText: '#8e1b22'
      }
    });

    expect(toAdaptiveCardHostConfig(theme)).toMatchObject({
      containerStyles: {
        default: {
          foregroundColors: {
            accent: { default: theme.palette.themeDark, subtle: theme.palette.themeDark },
            good: { default: '#185c18', subtle: '#185c18' },
            warning: { default: '#5a4300', subtle: '#5a4300' },
            attention: { default: '#8e1b22', subtle: '#8e1b22' }
          }
        }
      }
    });
  });

  it('rebuilds surfaces and status colors for an inverted theme without changing the earlier config', () => {
    const light = createTheme();
    const dark = createTheme({
      isInverted: true,
      palette: {
        white: '#201f1e',
        neutralPrimary: '#f3f2f1',
        neutralSecondary: '#c8c6c4',
        neutralLight: '#3b3a39'
      }
    });
    const lightConfig = toAdaptiveCardHostConfig(light);
    const darkConfig = toAdaptiveCardHostConfig(dark);

    expect(darkConfig).toMatchObject({
      separator: { lineColor: dark.palette.neutralLight },
      containerStyles: {
        default: {
          backgroundColor: '#201f1e',
          foregroundColors: {
            default: { default: '#f3f2f1', subtle: '#c8c6c4' },
            good: { default: dark.semanticColors.successText },
            attention: { default: dark.semanticColors.errorText }
          }
        },
        warning: { backgroundColor: dark.semanticColors.warningBackground }
      }
    });
    expect(lightConfig).toMatchObject({
      containerStyles: { default: { backgroundColor: light.palette.white } }
    });
  });

  it('preserves the existing action and authored-card host settings', () => {
    expect(toAdaptiveCardHostConfig(createTheme())).toMatchObject({
      actions: {
        actionsOrientation: 'Horizontal',
        actionAlignment: 'Left',
        maxActions: 5,
        showCard: { actionMode: 'Inline', inlineTopMargin: 8 }
      },
      adaptiveCard: { allowCustomStyle: false },
      imageSet: { imageSize: 'Medium', maxImageHeight: 100 }
    });
  });
});
