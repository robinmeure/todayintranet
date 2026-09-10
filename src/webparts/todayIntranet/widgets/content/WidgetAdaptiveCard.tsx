import * as React from 'react';
import { useTheme } from '@fluentui/react/lib/Theme';
import { IWidgetError } from './IWidgetContent';
import { AdaptiveCardPayload } from './toAdaptiveCard';
import { WidgetErrorMessage, WidgetLoading } from './WidgetStates';
import { toAdaptiveCardHostConfig } from './adaptiveCardHostConfig';
import styles from './WidgetContent.module.scss';

type AdaptiveCardsLibrary = typeof import('adaptivecards');

let pending: Promise<AdaptiveCardsLibrary> | undefined;

/**
 * Loaded on demand. The renderer is a sizeable dependency, so only a dashboard that
 * actually shows an Adaptive Card pays for downloading it.
 */
function loadAdaptiveCards(): Promise<AdaptiveCardsLibrary> {
  if (!pending) {
    pending = import(/* webpackChunkName: 'adaptivecards' */ 'adaptivecards');
  }
  return pending;
}

export interface IWidgetAdaptiveCardProps {
  /** Adaptive Card payload, schema 1.5 or lower. */
  card: AdaptiveCardPayload;
  ariaLabel?: string;
  /** Handles `Action.Submit`. `Action.OpenUrl` is handled here. */
  onSubmit?(data: unknown): void;
}

/**
 * Renders an Adaptive Card inside a tile, themed from the site. Any widget can use
 * it, and `WidgetItems` uses it to draw ordinary widget items as a card.
 */
export const WidgetAdaptiveCard: React.FunctionComponent<IWidgetAdaptiveCardProps> = (props) => {
  const { card, ariaLabel, onSubmit } = props;
  const theme = useTheme();
  const host = React.useRef<HTMLDivElement>(null);
  const [library, setLibrary] = React.useState<AdaptiveCardsLibrary | undefined>(undefined);
  const [error, setError] = React.useState<IWidgetError | undefined>(undefined);

  const onSubmitRef = React.useRef(onSubmit);
  onSubmitRef.current = onSubmit;

  // Serialised so a caller that builds the payload inline does not re-render the
  // card on every single React render.
  const payload = JSON.stringify(card);

  React.useEffect(() => {
    let cancelled = false;
    loadAdaptiveCards()
      .then((loaded) => {
        if (!cancelled) {
          setLibrary(loaded);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setError({ message: 'The Adaptive Card renderer could not be loaded.' });
        }
      });
    return () => {
      cancelled = true;
    };
  }, []);

  React.useEffect(() => {
    const container = host.current;
    if (!library || !container) {
      return undefined;
    }

    try {
      // Without this, Adaptive Cards falls back to any markdown-it it happens to find
      // on the page and writes the result as HTML. Leaving it unprocessed keeps every
      // text block plain text, whoever authored the payload.
      library.AdaptiveCard.onProcessMarkdown = (text, result) => {
        result.didProcess = false;
        result.outputHtml = text;
      };

      const adaptiveCard = new library.AdaptiveCard();
      adaptiveCard.hostConfig = new library.HostConfig(toAdaptiveCardHostConfig(theme));
      adaptiveCard.onExecuteAction = (action) => {
        if (action instanceof library.OpenUrlAction) {
          if (action.url) {
            window.open(action.url, '_blank', 'noopener,noreferrer');
          }
        } else if (action instanceof library.SubmitAction && onSubmitRef.current) {
          onSubmitRef.current(action.data);
        }
      };
      adaptiveCard.parse(JSON.parse(payload));

      const rendered = adaptiveCard.render();
      container.innerHTML = '';
      if (rendered) {
        container.appendChild(rendered);
      }
      setError(undefined);
    } catch {
      setError({
        message: 'That Adaptive Card could not be rendered. Check the payload in the tile settings.',
        isActionRequired: true
      });
    }

    return () => {
      container.innerHTML = '';
    };
  }, [library, payload, theme]);

  return (
    <>
      {error && <WidgetErrorMessage error={error} />}
      {!library && !error && <WidgetLoading label="Loading the card…" rows={2} />}
      {/* Kept mounted even while an error shows, so a corrected payload can render. */}
      <div className={styles.adaptiveCard} ref={host} aria-label={ariaLabel} hidden={!!error} />
    </>
  );
};
