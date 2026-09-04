import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import {
  LAYOUT_PRESETS,
  ROW_SIZES,
  ILayoutPreset,
  GRID_COLUMNS,
  getRowSize
} from '../model/LayoutPresets';
import styles from './LayoutPresetPanel.module.scss';

export interface ILayoutPresetPanelProps {
  isOpen: boolean;
  /** Preset currently applied, or undefined when the layout has been customised. */
  presetId?: string;
  rowSizeId?: string;
  /** Number of widgets on the dashboard, used to describe what applying will do. */
  widgetCount: number;
  onDismiss(): void;
  onApply(presetId: string, rowSizeId: string): void;
}

/** Miniature of the slot bands a preset produces, at the chosen tile height. */
const PresetPreview: React.FunctionComponent<{ preset: ILayoutPreset; rows: number }> = ({
  preset,
  rows
}) => {
  // One repeat is enough to show a banner preset; two for the others.
  const bands = preset.banner ? [0] : [0, 1];
  const height = Math.round(10 + rows * 1.6);

  return (
    <div className={styles.preview} aria-hidden={true}>
      {preset.banner && <div className={styles.previewRow} style={{ height }}>
        <div className={styles.previewTile} style={{ width: '100%' }} />
      </div>}
      {bands.map((band) => (
        <div key={band} className={styles.previewRow} style={{ height }}>
          {preset.columns.map((span, index) => (
            <div
              key={index}
              className={styles.previewTile}
              style={{ width: `${(span / GRID_COLUMNS) * 100}%` }}
            />
          ))}
        </div>
      ))}
    </div>
  );
};

/**
 * Lets the user pick how the dashboard is arranged: a column preset, and how tall
 * a row of tiles is. Applying pours the existing widgets into the new slots.
 */
export const LayoutPresetPanel: React.FunctionComponent<ILayoutPresetPanelProps> = (props) => {
  const { isOpen, presetId, rowSizeId, widgetCount, onDismiss, onApply } = props;

  const [draftPresetId, setDraftPresetId] = React.useState<string>(presetId ?? LAYOUT_PRESETS[0].id);
  const [draftRowSizeId, setDraftRowSizeId] = React.useState<string>(getRowSize(rowSizeId).id);

  // Re-seed from the applied layout every time the panel is opened.
  React.useEffect(() => {
    if (isOpen) {
      setDraftPresetId(presetId ?? LAYOUT_PRESETS[0].id);
      setDraftRowSizeId(getRowSize(rowSizeId).id);
    }
  }, [isOpen, presetId, rowSizeId]);

  const rowOptions: IChoiceGroupOption[] = ROW_SIZES.map((size) => ({
    key: size.id,
    text: `${size.name} — ${size.description}`
  }));

  const rows = getRowSize(draftRowSizeId).rows;

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Choose a layout"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={() => (
        <div className={styles.footer}>
          <PrimaryButton
            text="Apply layout"
            disabled={widgetCount === 0}
            onClick={() => onApply(draftPresetId, draftRowSizeId)}
          />
          <DefaultButton text="Cancel" onClick={onDismiss} />
        </div>
      )}
      isFooterAtBottom={true}
    >
      <p className={styles.intro}>
        {widgetCount === 0
          ? 'Add a widget first — a layout arranges the widgets already on your dashboard.'
          : `Rearranges your ${widgetCount} widget${widgetCount === 1 ? '' : 's'} into the slots below, in their current order. You can still drag anything afterwards.`}
      </p>

      <ChoiceGroup
        label="Tile height"
        selectedKey={draftRowSizeId}
        options={rowOptions}
        onChange={(_, option) => option && setDraftRowSizeId(option.key)}
      />

      <div className={styles.gallery} role="group" aria-label="Column layout">
        {LAYOUT_PRESETS.map((preset) => {
          const isSelected = preset.id === draftPresetId;
          return (
            <button
              key={preset.id}
              type="button"
              aria-pressed={isSelected}
              className={isSelected ? `${styles.card} ${styles.cardSelected}` : styles.card}
              onClick={() => setDraftPresetId(preset.id)}
            >
              <PresetPreview preset={preset} rows={rows} />
              <div className={styles.cardTitle}>{preset.name}</div>
              <div className={styles.cardDescription}>{preset.description}</div>
            </button>
          );
        })}
      </div>
    </Panel>
  );
};
