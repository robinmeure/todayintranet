import { ITheme } from '@fluentui/react/lib/Theme';

export function toAdaptiveCardHostConfig(theme: ITheme): Record<string, unknown> {
  const { palette, semanticColors } = theme;
  const foregroundColors = {
    default: { default: palette.neutralPrimary, subtle: palette.neutralSecondary },
    accent: { default: palette.themeDark, subtle: palette.themeDark },
    good: { default: semanticColors.successText, subtle: semanticColors.successText },
    warning: { default: semanticColors.warningText, subtle: semanticColors.warningText },
    attention: { default: semanticColors.errorText, subtle: semanticColors.errorText },
    light: { default: palette.neutralTertiary, subtle: palette.neutralQuaternary },
    dark: { default: palette.neutralDark, subtle: palette.neutralSecondary }
  };

  return {
    fontFamily: theme.fonts.medium.fontFamily,
    // The widget body already supplies the outer padding.
    spacing: { none: 0, small: 4, default: 8, medium: 12, large: 16, extraLarge: 24, padding: 0 },
    separator: { lineThickness: 1, lineColor: palette.neutralLight },
    fontSizes: { small: 12, default: 14, medium: 16, large: 18, extraLarge: 22 },
    lineHeights: { small: 16, default: 20, medium: 24, large: 24, extraLarge: 28 },
    fontWeights: { lighter: 300, default: 400, bolder: 600 },
    containerStyles: {
      default: { backgroundColor: palette.white, foregroundColors },
      emphasis: { backgroundColor: palette.neutralLighter, foregroundColors },
      accent: { backgroundColor: palette.themeLighterAlt, foregroundColors },
      good: { backgroundColor: semanticColors.successBackground, foregroundColors },
      warning: { backgroundColor: semanticColors.warningBackground, foregroundColors },
      attention: { backgroundColor: semanticColors.errorBackground, foregroundColors }
    },
    actions: {
      actionsOrientation: 'Horizontal',
      actionAlignment: 'Left',
      buttonSpacing: 8,
      maxActions: 5,
      spacing: 'Default',
      showCard: { actionMode: 'Inline', inlineTopMargin: 8 }
    },
    factSet: {
      title: { weight: 'Bolder', wrap: true, maxWidth: 150 },
      value: { wrap: true },
      spacing: 8
    },
    adaptiveCard: { allowCustomStyle: false },
    imageSet: { imageSize: 'Medium', maxImageHeight: 100 }
  };
}
