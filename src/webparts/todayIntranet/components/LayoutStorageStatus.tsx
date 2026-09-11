import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ILayoutStoreStatus, LayoutStoreAction, LayoutSaveState } from '../services/ILayoutStore';

const STATE_LABELS: Record<LayoutSaveState, string> = {
  loading: 'Loading layout',
  saved: 'Layout saved',
  saving: 'Saving layout',
  pending: 'Layout sync pending',
  conflict: 'Layout conflict',
  error: 'Layout storage error'
};

const ACTION_LABELS: Record<LayoutStoreAction, string> = {
  retry: 'Retry',
  'use-remote': 'Use saved version',
  'keep-local': 'Keep my version',
  'import-legacy': 'Import legacy layout',
  reset: 'Reset layout',
  export: 'Export recovery data'
};

export interface ILayoutStorageStatusProps {
  status: ILayoutStoreStatus;
  error?: string;
  busy: boolean;
  confirmation?: 'reset' | 'import-legacy';
  onAction(action: LayoutStoreAction): void;
  onConfirm(): void;
  onCancel(): void;
}

export const LayoutStorageStatus: React.FunctionComponent<ILayoutStorageStatusProps> = ({
  status, error, busy, confirmation, onAction, onConfirm, onCancel
}) => {
  const showBanner = !!error || status.state === 'error' || status.state === 'conflict';
  const failed = !!error || status.state === 'error';
  const warning = status.state === 'conflict';
  const messageType = failed ? MessageBarType.error : warning ? MessageBarType.warning : MessageBarType.info;
  const importing = confirmation === 'import-legacy';

  return (
    <>
      {showBanner && (
        <MessageBar
          messageBarType={messageType}
          isMultiline={true}
          role={failed || warning ? 'alert' : 'status'}
          aria-live={failed ? 'assertive' : 'polite'}
        >
          <strong>{STATE_LABELS[error ? 'error' : status.state]}.</strong> {status.message}
          {status.mode === 'local' && (
            <span> Browser-only storage: this layout is not synced to SharePoint or other browsers.</span>
          )}
          {error && <span> {error}</span>}
          {status.state === 'conflict' && (
            <p>Use the saved version to replace your changes, or keep your version to replace the other saved version.</p>
          )}
          {status.actions.map((action) => (
            <DefaultButton
              key={action}
              text={ACTION_LABELS[action]}
              disabled={busy}
              onClick={() => onAction(action)}
            />
          ))}
        </MessageBar>
      )}
      <Dialog
        hidden={!confirmation}
        onDismiss={onCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: importing ? 'Import legacy layout into this dashboard?' : 'Reset this dashboard?',
          subText: importing
            ? 'The legacy layout has an unknown source site. Importing it replaces the layout for this dashboard, not any other dashboard. On SharePoint, the imported layout may sync to this dashboard. Export recovery data first if you want a backup.'
            : 'This replaces this dashboard with its starter layout and may discard recoverable changes. Export recovery data first if you want a backup. On SharePoint, the reset layout may sync to this dashboard.'
        }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton text={importing ? 'Import into this dashboard' : 'Reset this dashboard'} disabled={busy} onClick={onConfirm} />
          <DefaultButton text="Cancel" disabled={busy} onClick={onCancel} />
        </DialogFooter>
      </Dialog>
    </>
  );
};
