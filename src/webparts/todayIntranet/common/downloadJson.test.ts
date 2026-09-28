import { DOWNLOAD_URL_REVOKE_DELAY_MS, downloadJson } from './downloadJson';

describe('downloadJson', () => {
  let create: jest.Mock<string, [Blob]>;
  let revoke: jest.Mock<void, [string]>;
  let oldCreate: PropertyDescriptor | undefined;
  let oldRevoke: PropertyDescriptor | undefined;

  beforeEach(() => {
    jest.useFakeTimers();
    create = jest.fn<string, [Blob]>(() => 'blob:download');
    revoke = jest.fn<void, [string]>();
    oldCreate = Object.getOwnPropertyDescriptor(URL, 'createObjectURL');
    oldRevoke = Object.getOwnPropertyDescriptor(URL, 'revokeObjectURL');
    Object.defineProperty(URL, 'createObjectURL', { configurable: true, value: create });
    Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, value: revoke });
  });

  afterEach(() => {
    if (oldCreate) { Object.defineProperty(URL, 'createObjectURL', oldCreate); }
    else { delete (URL as Partial<typeof URL>).createObjectURL; }
    if (oldRevoke) { Object.defineProperty(URL, 'revokeObjectURL', oldRevoke); }
    else { delete (URL as Partial<typeof URL>).revokeObjectURL; }
    jest.restoreAllMocks();
    jest.useRealTimers();
  });

  it('clicks an attached link with the file name and JSON content, then defers revocation', () => {
    let clicked: { href: string; download: string; attached: boolean } | undefined;
    jest.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      clicked = { href: this.getAttribute('href') ?? '', download: this.download, attached: document.body.contains(this) };
    });

    downloadJson('{"a":1}', 'recovery.json');

    expect(clicked).toEqual({ href: 'blob:download', download: 'recovery.json', attached: true });
    const blob = create.mock.calls[0][0];
    expect(blob.type).toBe('application/json');
    expect(blob.size).toBe('{"a":1}'.length);
    expect(document.querySelector('a[download]')).toBeNull();
    expect(revoke).not.toHaveBeenCalled();

    jest.advanceTimersByTime(DOWNLOAD_URL_REVOKE_DELAY_MS);
    expect(revoke).toHaveBeenCalledWith('blob:download');
  });

  it('revokes immediately and rethrows when the download cannot start', () => {
    jest.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {
      throw new Error('Blocked');
    });

    expect(() => downloadJson('{}', 'recovery.json')).toThrow('Blocked');
    expect(revoke).toHaveBeenCalledWith('blob:download');
    expect(document.querySelector('a[download]')).toBeNull();
    jest.runAllTimers();
    expect(revoke).toHaveBeenCalledTimes(1);
  });
});
