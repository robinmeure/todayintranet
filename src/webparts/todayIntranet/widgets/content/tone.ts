import { WidgetTone } from './IWidgetContent';
import styles from './WidgetContent.module.scss';

const TONE_CLASS: Record<WidgetTone, string> = {
  neutral: styles.toneNeutral,
  accent: styles.toneAccent,
  success: styles.toneSuccess,
  warning: styles.toneWarning,
  danger: styles.toneDanger
};

/** Class carrying the tone colour, applied to whatever should take it as `currentColor`. */
export function toneClass(tone: WidgetTone | undefined, fallback: WidgetTone = 'accent'): string {
  return TONE_CLASS[tone ?? fallback];
}
