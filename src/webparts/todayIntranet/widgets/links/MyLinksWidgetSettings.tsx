import * as React from 'react';
import { DefaultButton, IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { TextField } from '@fluentui/react/lib/TextField';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { IWidgetContext } from '../IWidget';
import { NumberSetting, SettingsSurface, TextSetting, textSetting } from '../content';
import { DEFAULT_ITEMS_PER_PAGE, DEFAULT_LINKS_SETTING, IMyLink, ITEMS_PER_PAGE_BOUNDS, parseLinksSetting } from './linksSettings';
import styles from './MyLinksWidgetSettings.module.scss';

interface ILinkDraft {
  index?: number;
  link: IMyLink;
}

export const MyLinksWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const stored = textSetting(context, 'links', DEFAULT_LINKS_SETTING);
  const parsed = parseLinksSetting(stored);
  const links = parsed.links ?? [];
  const [draft, setDraft] = React.useState<ILinkDraft | undefined>(() =>
    context.settingsAction === 'add' ? { link: { title: '', url: '' } } : undefined);
  const [query, setQuery] = React.useState('');
  const [error, setError] = React.useState<string>();
  const [deleting, setDeleting] = React.useState<number>();
  const [announcement, setAnnouncement] = React.useState('');
  const [dragging, setDragging] = React.useState<number>();
  const [dropTarget, setDropTarget] = React.useState<number>();
  const drag = React.useRef<{ index: number; x: number; y: number; active: boolean }>();
  const list = React.useRef<HTMLOListElement>(null);
  const previousStored = React.useRef(stored);

  React.useEffect(() => {
    if (previousStored.current === stored) { return; }
    previousStored.current = stored;
    setDraft(undefined);
    setDeleting(undefined);
    setError(undefined);
    drag.current = undefined;
    setDragging(undefined);
    setDropTarget(undefined);
  }, [stored]);

  const persist = (next: IMyLink[], message: string): boolean => {
    const validated = parseLinksSetting(JSON.stringify(next));
    if (validated.error) {
      setError(validated.error);
      return false;
    }
    context.updateSettings({ ...context.settings, links: JSON.stringify(validated.links) });
    setAnnouncement(message);
    setError(undefined);
    return true;
  };
  const back = (): void => { setDraft(undefined); setError(undefined); };
  const changeDraft = (values: Partial<IMyLink>): void => {
    if (draft) {
      setDraft({ ...draft, link: { ...draft.link, ...values } });
      setError(undefined);
    }
  };
  const save = (): void => {
    if (!draft) { return; }
    const next = links.slice();
    if (draft.index === undefined) { next.push(draft.link); }
    else { next[draft.index] = draft.link; }
    if (persist(next, draft.index === undefined ? 'Link added.' : 'Link updated.')) { back(); }
  };
  const reorder = (from: number, to: number): void => {
    if (from === to) { return; }
    const next = links.slice();
    const moved = next.splice(from, 1)[0];
    next.splice(to, 0, moved);
    persist(next, `${moved.title} moved to position ${to + 1}.`);
  };
  const hitTarget = (x: number, y: number): number | undefined => {
    const row = document.elementFromPoint(x, y)?.closest<HTMLLIElement>('li[data-link-index]');
    if (!row || !list.current?.contains(row)) { return undefined; }
    return Number(row.dataset.linkIndex);
  };
  const cancelDrag = (): void => {
    drag.current = undefined;
    setDragging(undefined);
    setDropTarget(undefined);
  };
  const sort = (descending: boolean): void => {
    persist(links.slice().sort((left, right) =>
      left.title.localeCompare(right.title, undefined, { sensitivity: 'base' }) * (descending ? -1 : 1)),
      `Links sorted ${descending ? 'Z to A' : 'A to Z'}.`);
  };
  const needle = query.trim().toLowerCase();
  const filtered = links.map((link, index) => ({ link, index }))
    .filter(({ link }) => `${link.title} ${link.url}`.toLowerCase().indexOf(needle) >= 0);

  return (
    <div className={styles.authoring}>
      {draft && <div className={styles.header}>
        <IconButton iconProps={{ iconName: 'Back' }} ariaLabel="Back to My links" onClick={back} />
        <h2>{draft.index === undefined ? 'Add a new link' : 'Edit link'}</h2>
      </div>}
      <div className={styles.body}>
        {draft ? (
          <form className={styles.form} onSubmit={(event) => { event.preventDefault(); save(); }}>
            <TextField label="Name" placeholder="Type the name of the link" required autoFocus
              value={draft.link.title} maxLength={120}
              onChange={(_, value) => changeDraft({ title: value ?? '' })} />
            <Toggle label="VPN OFF/ON" inlineLabel checked={draft.link.badge === 'VPN'}
              styles={{
                root: { flexDirection: 'row', justifyContent: 'flex-start' },
                label: { margin: '0 8px 0 0', order: -1 }
              }}
              onChange={(_, checked) => changeDraft({ badge: checked ? 'VPN' : undefined })} />
            <TextField label="Link" placeholder="Add the URL" required value={draft.link.url} maxLength={2048}
              onChange={(_, value) => changeDraft({ url: value ?? '' })} />
            <TextField label="Description" placeholder="Type the description of your link" multiline rows={5}
              value={draft.link.description ?? ''} maxLength={240}
              onChange={(_, value) => changeDraft({ description: value ?? '' })} />
            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
            <div className={styles.formActions}>
              {draft.index !== undefined && <DefaultButton text="Delete" className={styles.textButton}
                onClick={() => setDeleting(draft.index)} />}
              <DefaultButton text="Cancel" className={styles.textButton} onClick={back} />
              <PrimaryButton text={draft.index === undefined ? 'Add link' : 'Update'} type="submit" />
            </div>
          </form>
        ) : (
          <SettingsSurface description={
            'Access your most important tools from one place and customize the shortcuts that help you start your work quickly.'
          }>
            <PrimaryButton text="Add a new link" className={styles.addButton}
              disabled={!!parsed.error || links.length >= 100}
              onClick={() => { setError(undefined); setDraft({ link: { title: '', url: '' } }); }} />
            {links.length >= 100 && <MessageBar>You can save at most 100 links. Delete a link before adding another.</MessageBar>}
            {parsed.error && <MessageBar messageBarType={MessageBarType.error}>
              {parsed.error} Correct Links (JSON) in Advanced settings below.
            </MessageBar>}
            <h3 className={styles.label}>The list of your favorite links</h3>
            <SearchBox placeholder="Find your links" ariaLabel="Find your links" value={query}
              onChange={(_, value) => setQuery(value ?? '')} />
            <h3 className={styles.label}>Drag the links or categories to reorder</h3>
            <span className={styles.srOnly}>You can also focus a link and use Alt+Up or Alt+Down to reorder it.</span>
            {!parsed.error && !filtered.length && <span>{links.length ? 'No links found' : 'No links added'}</span>}
            <ol ref={list} className={styles.links} aria-label="Manage My links">
              {filtered.map(({ link, index }) => (
                <li key={index} tabIndex={0} data-link-index={index}
                  aria-label={`${link.title}. Position ${index + 1} of ${links.length}. Alt+Up or Alt+Down to reorder.`}
                  className={`${styles.row} ${dragging === index ? styles.dragging : ''} ${dropTarget === index ? styles.dropTarget : ''}`}
                  onPointerDown={(event) => {
                    if (event.button !== 0 || !event.isPrimary ||
                      (event.target instanceof Element && event.target.closest('button'))) { return; }
                    event.stopPropagation();
                    event.currentTarget.focus();
                    event.currentTarget.setPointerCapture(event.pointerId);
                    drag.current = { index, x: event.clientX, y: event.clientY, active: false };
                  }}
                  onPointerMove={(event) => {
                    const current = drag.current;
                    if (!current || current.index !== index) { return; }
                    if (!current.active && Math.hypot(event.clientX - current.x, event.clientY - current.y) < 5) { return; }
                    current.active = true;
                    setDragging(index);
                    setDropTarget(hitTarget(event.clientX, event.clientY));
                  }}
                  onPointerUp={(event) => {
                    if (drag.current?.active) {
                      const target = hitTarget(event.clientX, event.clientY);
                      if (target !== undefined) { reorder(drag.current.index, target); }
                    }
                    cancelDrag();
                  }}
                  onPointerCancel={cancelDrag} onLostPointerCapture={cancelDrag}
                  onKeyDown={(event) => {
                    if (event.target !== event.currentTarget || !event.altKey) { return; }
                    const offset = event.key === 'ArrowUp' ? -1 : event.key === 'ArrowDown' ? 1 : 0;
                    if (!offset) { return; }
                    event.preventDefault();
                    const to = index + offset;
                    if (to >= 0 && to < links.length) {
                      reorder(index, to);
                      const rows = event.currentTarget.parentElement?.querySelectorAll<HTMLLIElement>('li');
                      // Restore focus to the moved row when the complete list is visible.
                      if (!needle) { rows?.[to]?.focus(); }
                    }
                  }}>
                  <span className={styles.rowTitle} title={link.title}>{link.title}</span>
                  <IconButton iconProps={{ iconName: 'Heart' }} title="Remove link"
                    ariaLabel={`Delete ${link.title}`} onClick={() => setDeleting(index)} />
                  <IconButton iconProps={{ iconName: 'Edit' }} title="Edit link"
                    ariaLabel={`Edit ${link.title}`} onClick={() => setDraft({ index, link: { ...link } })} />
                </li>
              ))}
            </ol>
            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
            <details className={styles.advanced}>
              <summary>Advanced settings</summary>
              <SettingsSurface>
                <DefaultButton text="Sort A-Z" disabled={!!parsed.error || links.length < 2} onClick={() => sort(false)} />
                <DefaultButton text="Sort Z-A" disabled={!!parsed.error || links.length < 2} onClick={() => sort(true)} />
                <NumberSetting context={context} settingKey="itemsPerPage" label="Links per page"
                  fallback={DEFAULT_ITEMS_PER_PAGE} {...ITEMS_PER_PAGE_BOUNDS} />
                <TextSetting context={context} settingKey="links" label="Links (JSON)" fallback={DEFAULT_LINKS_SETTING}
                  multiline rows={8} commitOn="blur" validate={(value) => parseLinksSetting(value).error} />
              </SettingsSurface>
            </details>
          </SettingsSurface>
        )}
        <span className={styles.srOnly} role="status" aria-live="polite">{announcement}</span>
      </div>
      <Dialog hidden={deleting === undefined} onDismiss={() => setDeleting(undefined)}
        dialogContentProps={{
          type: DialogType.normal, title: 'Delete link',
          subText: deleting === undefined ? '' : `Delete "${links[deleting]?.title}"?`
        }} modalProps={{ isBlocking: true }}>
        <DialogFooter>
          <PrimaryButton text="Confirm delete" onClick={() => {
            if (persist(links.filter((_, index) => index !== deleting), 'Link deleted.')) {
              setDeleting(undefined);
              back();
            }
          }} />
          <DefaultButton text="Cancel delete" onClick={() => setDeleting(undefined)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
