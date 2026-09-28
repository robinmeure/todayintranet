/**
 * Some browsers cancel a Blob download whose URL is revoked in the same task as the
 * click, so the URL outlives the click briefly. This does not confirm completion.
 */
export const DOWNLOAD_URL_REVOKE_DELAY_MS: number = 1000;

/** Offers `json` as a file download. Throws if the browser refuses to start it. */
export function downloadJson(json: string, fileName: string): void {
  const url = URL.createObjectURL(new Blob([json], { type: 'application/json' }));
  const link = document.createElement('a');
  try {
    link.href = url;
    link.download = fileName;
    link.style.display = 'none';
    // Firefox only starts downloads for links that are attached to the document.
    document.body.appendChild(link);
    link.click();
  } catch (error) {
    URL.revokeObjectURL(url);
    throw error;
  } finally {
    link.remove();
  }
  window.setTimeout(() => URL.revokeObjectURL(url), DOWNLOAD_URL_REVOKE_DELAY_MS);
}
