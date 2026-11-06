import { beforeEach, describe, expect, it, vi } from 'vitest';

const captureRequestHeader = vi.fn();
const start = vi.fn(async () => {});
const close = vi.fn(async () => {});
const createFetchproxyTransport = vi.fn((_opts: Record<string, unknown>) => ({
  start,
  close,
  server: { captureRequestHeader },
}));

vi.mock('@chrischall/mcp-utils/fetchproxy', async (importOriginal) => {
  const real = await importOriginal<typeof import('@chrischall/mcp-utils/fetchproxy')>();
  return { ...real, createFetchproxyTransport };
});

const { captureTokenViaFetchproxy, captureHost, OUTLOOK_WEB_HOSTS } = await import('../src/auth-fetchproxy.js');

// Opaque filler: long enough to pass the length guard, never a real token.
const FAKE_TOKEN = 'x'.repeat(200);

beforeEach(() => {
  captureRequestHeader.mockReset();
  start.mockClear();
  close.mockClear();
  createFetchproxyTransport.mockClear();
});

interface HelloOpts {
  domains: string[];
  captureHeaders: { host: string; path?: string; headerName: string }[];
  port: number;
}
const helloOpts = () => createFetchproxyTransport.mock.calls[0][0] as unknown as HelloOpts;

describe('captureHost', () => {
  it('defaults to outlook.cloud.microsoft, where a current tenant’s tab lives', () => {
    expect(captureHost({})).toBe('outlook.cloud.microsoft');
  });

  it('takes the legacy host from OUTLOOK_WEB_HOST, in any case', () => {
    expect(captureHost({ OUTLOOK_WEB_HOST: 'Outlook.Office.com ' })).toBe('outlook.office.com');
  });

  it('refuses a host that is not an Outlook Web host, naming the ones it takes', () => {
    expect(() => captureHost({ OUTLOOK_WEB_HOST: 'office.com' })).toThrow(
      /OUTLOOK_WEB_HOST.*outlook\.cloud\.microsoft.*outlook\.office\.com/,
    );
  });

  it('treats an empty or unsubstituted value as unset', () => {
    expect(captureHost({ OUTLOOK_WEB_HOST: '' })).toBe('outlook.cloud.microsoft');
    expect(captureHost({ OUTLOOK_WEB_HOST: '${OUTLOOK_WEB_HOST}' })).toBe('outlook.cloud.microsoft');
  });

  it('only ever names one of the Outlook Web hosts', () => {
    expect(OUTLOOK_WEB_HOSTS).toEqual(['outlook.cloud.microsoft', 'outlook.office.com']);
  });
});

describe('captureTokenViaFetchproxy', () => {
  it('declares ONE domain and ONE capture, so the bridge opens no second tab', async () => {
    // Declaring both hosts made the bridge open a background relay tab on
    // outlook.office.com after every hello. Signed in, that tab redirects to
    // outlook.cloud.microsoft, so no tab ever matched it and every capture
    // opened another one.
    captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);

    await expect(captureTokenViaFetchproxy({})).resolves.toBe(FAKE_TOKEN);

    expect(createFetchproxyTransport).toHaveBeenCalledTimes(1);
    expect(helloOpts().domains).toEqual(['outlook.cloud.microsoft']);
    expect(helloOpts().captureHeaders).toEqual([
      { host: 'outlook.cloud.microsoft', path: '/*', headerName: 'authorization' },
    ]);
    expect(helloOpts().port).toBe(37_149);
    expect(captureRequestHeader).toHaveBeenCalledTimes(1);
    expect(captureRequestHeader).toHaveBeenCalledWith(
      expect.objectContaining({ host: 'outlook.cloud.microsoft', headerName: 'authorization', timeoutMs: 30_000 }),
    );
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('declares the same scope on every capture', async () => {
    // The bridge re-pairs on any change to the domain set and offers a grant
    // on any growth in capture headers, so two hellos must be identical.
    captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);
    await captureTokenViaFetchproxy({});
    await captureTokenViaFetchproxy({});
    const [a, b] = createFetchproxyTransport.mock.calls.map((c) => c[0] as unknown as HelloOpts);
    expect(b.domains).toEqual(a.domains);
    expect(b.captureHeaders).toEqual(a.captureHeaders);
  });

  it('captures from the legacy host alone when OUTLOOK_WEB_HOST names it', async () => {
    captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);
    await captureTokenViaFetchproxy({ OUTLOOK_WEB_HOST: 'outlook.office.com' });
    expect(helloOpts().domains).toEqual(['outlook.office.com']);
    expect(helloOpts().captureHeaders.map((d) => d.host)).toEqual(['outlook.office.com']);
  });

  it('reads the capture window per capture, not at import', async () => {
    captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);
    await captureTokenViaFetchproxy({ OUTLOOK_CAPTURE_TIMEOUT: '12' });
    expect(captureRequestHeader).toHaveBeenCalledWith(expect.objectContaining({ timeoutMs: 12_000 }));
  });

  it('names the host and what to do when the tab makes no request', async () => {
    captureRequestHeader.mockRejectedValue(new Error('timed out'));
    await expect(captureTokenViaFetchproxy({})).rejects.toThrow(
      /could not capture an Outlook token.*outlook\.cloud\.microsoft.*OUTLOOK_WEB_HOST/s,
    );
  });
});
