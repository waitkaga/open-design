// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { CLIENT_PPTX_VERSION, isClientPptxRequest, isClientPptxResult, pptxExportRoute, redactPptxAssetUrl } from '../../src/runtime/client-pptx-protocol';
import { cjkPromotedFontFamily, measureAuthoredSlideBox, prepareDeckStage, runDomToPptx, showAllSlides, SLIDE_SELECTOR, DECK_STAGE_SELECTOR, HIDE_CHROME_SELECTOR } from '../../src/runtime/pptx-export-normalizer';
import { buildSrcdoc, PREVIEW_REDIRECT_LOOP_MESSAGE } from '../../src/runtime/srcdoc';

describe('browser PPTX protocol and routing', () => {
  it('validates correlated success and error results and rejects malformed payloads', () => {
    const requestId = '1234567890123456';
    const result = { type: 'od:pptx-export-result', requestId, ok: true, buffer: new ArrayBuffer(4), warnings: [] };
    expect(isClientPptxResult(result, requestId)).toBe(true);
    expect(isClientPptxResult(result, 'other')).toBe(false);
    expect(isClientPptxResult({ ...result, buffer: new Uint8Array(4) }, requestId)).toBe(false);
    expect(isClientPptxResult({ ...result, buffer: new ArrayBuffer(0) }, requestId)).toBe(false);
    expect(isClientPptxResult({ ...result, warnings: [123] }, requestId)).toBe(false);
    expect(isClientPptxResult({ ...result, ok: false, error: 'Missing image' }, requestId)).toBe(true);
    expect(isClientPptxRequest({ type: 'od:pptx-export', requestId, fileName: 'deck.pptx' })).toBe(true);
    expect(isClientPptxRequest({ type: 'od:pptx-export', requestId: '', fileName: 'deck' })).toBe(false);
  });
  it('keeps the existing server path and selects browser export only for an explicit missing renderer', () => {
    expect(pptxExportRoute(true, true)).toBe('server');
    expect(pptxExportRoute(null, true)).toBe('server');
    expect(pptxExportRoute(false, true)).toBe('browser');
    expect(pptxExportRoute(false, false)).toBeNull();
  });
  it('does not expose credentials or scoped path tokens in asset errors', () => {
    expect(redactPptxAssetUrl('https://user:pass@example.com/secret-token/assets/image.png?token=secret#secret'))
      .toBe('https://example.com/.../image.png');
  });
  it('injects UMD then bridge before author CSP while preserving sandbox redirect guards', () => {
    const html = buildSrcdoc('<!doctype html><html><head><meta http-equiv="Content-Security-Policy" content="script-src none"></head><body><section class="slide">Hello</section></body></html>', {
      baseHref: 'https://preview.example.com/scoped/',
      exportPptxScripts: { bundle: '/* UMD */', bridge: '/* BRIDGE */' },
    });
    expect(html.indexOf('/* UMD */')).toBeLessThan(html.indexOf('/* BRIDGE */'));
    expect(html.indexOf('/* BRIDGE */')).toBeLessThan(html.indexOf('Content-Security-Policy'));
    expect(html).toContain('https://preview.example.com/scoped/');
    expect(html).toContain(PREVIEW_REDIRECT_LOOP_MESSAGE);
  });
});

describe('ported deck normalization', () => {
  afterEach(() => { document.body.innerHTML = ''; vi.unstubAllGlobals(); vi.restoreAllMocks(); });
  it('selects every supported slide family and excludes thumbnails, overview and notes', () => {
    document.body.innerHTML = '<div class="slide"></div><div data-screen-label="2"></div><div class="deck-slide"></div><div class="ppt-slide"></div><div class="overview"><div class="slide"></div></div><div class="mini-slide slide"></div><div class="thumb"><div class="ppt-slide"></div></div>';
    expect(showAllSlides(SLIDE_SELECTOR)).toBe(4);
    expect(document.querySelector('.overview .slide')?.getAttribute('style')).toBeNull();
    expect(document.querySelector('.ppt-slide')?.classList.contains('active')).toBe(true);
  });
  it('uses authored non-widescreen dimensions and removes deck scaling and chrome', () => {
    document.body.innerHTML = '<deck-stage width="1200" height="900" style="transform:scale(.5)"><div class="slide"></div><div class="deck-nav">Next</div></deck-stage>';
    const slide = document.querySelector<HTMLElement>('.slide')!;
    expect(measureAuthoredSlideBox(slide)).toEqual({ w: 1200, h: 900 });
    prepareDeckStage(HIDE_CHROME_SELECTOR, DECK_STAGE_SELECTOR);
    expect(document.querySelector('deck-stage')?.hasAttribute('noscale')).toBe(true);
    expect(document.querySelector<HTMLElement>('.deck-nav')!.style.display).toBe('none');
  });
  it('promotes the CJK family without changing Latin-only text', () => {
    expect(cjkPromotedFontFamily('Inter, "Noto Sans SC", sans-serif', '\u4e2d\u6587')).toBe('"Noto Sans SC", Inter, sans-serif');
    expect(cjkPromotedFontFamily('Inter, "Noto Sans SC", sans-serif', 'Hello')).toBeNull();
  });
  it('stabilizes headings and SVG className before passing native DOM to the converter', async () => {
    document.body.innerHTML = '<div class="slide" style="background:red"><h1>Hello<br>World</h1><svg class="vector"></svg></div>';
    const buffer = new ArrayBuffer(8);
    const blob = new Blob(['zip']);
    Object.defineProperty(blob, 'arrayBuffer', { value: async () => buffer });
    const exportToPptx = vi.fn(async (_slides: unknown, _options: unknown) => blob);
    vi.stubGlobal('domToPptx', { exportToPptx });
    const result = await runDomToPptx(SLIDE_SELECTOR);
    expect(result.buffer).toBe(buffer);
    expect(document.querySelector('h1')!.style.whiteSpace).toBe('nowrap');
    expect(typeof document.querySelector('svg')!.className).toBe('string');
    expect(exportToPptx.mock.calls[0]?.[0]).toBeDefined();
  });
});

describe('browser PPTX coordinator lifecycle', () => {
  let ports: Array<{ onmessage: ((event: { data: unknown }) => void) | null; start: ReturnType<typeof vi.fn>; close: ReturnType<typeof vi.fn> }>;
  let exporter: typeof import('../../src/runtime/clientPptxExport');
  const options = { sourceHtml: '<div class="slide">Hello</div>', baseHref: 'https://preview.example.com/scoped/', fileName: 'deck' };
  beforeEach(async () => {
    vi.resetModules();
    vi.useFakeTimers();
    ports = [];
    vi.spyOn(navigator, 'userAgent', 'get').mockReturnValue('Chrome/130.0');
    vi.stubGlobal('MessageChannel', class {
      port1 = { onmessage: null, start: vi.fn(), close: vi.fn() };
      port2 = { close: vi.fn() };
      constructor() { ports.push(this.port1); }
    });
    vi.stubGlobal('fetch', vi.fn(async (url: string) => url.endsWith('.json')
      ? { ok: true, json: async () => ({ enabled: true, version: CLIENT_PPTX_VERSION }) }
      : { ok: true, text: async () => `// ${CLIENT_PPTX_VERSION}\n/* test */` }));
    Object.defineProperty(URL, 'createObjectURL', { configurable: true, value: vi.fn(() => 'blob:export') });
    Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, value: vi.fn() });
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
    exporter = await import('../../src/runtime/clientPptxExport');
  });
  afterEach(() => {
    document.body.innerHTML = '';
    vi.useRealTimers(); vi.unstubAllGlobals(); vi.restoreAllMocks();
  });
  const flush = async () => { await vi.advanceTimersByTimeAsync(0); };
  const ready = (frame: HTMLIFrameElement) => window.dispatchEvent(new MessageEvent('message', {
    source: frame.contentWindow, data: { type: 'od:pptx-export-ready' },
  }));
  it('rejects other windows, correlates port replies, downloads once and leaves the live preview intact', async () => {
    const live = document.createElement('iframe');
    live.srcdoc = '<p>Live</p>';
    document.body.appendChild(live);
    const promise = exporter.exportDeckAsPptxInBrowser(options);
    await flush();
    const frame = document.querySelector<HTMLIFrameElement>('[data-od-pptx-export]')!;
    expect(frame.getAttribute('sandbox')).toBe('allow-scripts allow-downloads');
    window.dispatchEvent(new MessageEvent('message', { source: window, data: { type: 'od:pptx-export-ready' } }));
    expect(ports).toHaveLength(0);
    const post = vi.spyOn(frame.contentWindow!, 'postMessage').mockImplementation(() => {});
    ready(frame);
    expect(post).toHaveBeenCalledOnce();
    const requestId = frame.dataset.odPptxExport;
    ports[0]!.onmessage!({ data: { type: 'od:pptx-export-result', requestId: 'wrong', ok: false, error: 'spoof' } });
    ports[0]!.onmessage!({ data: { type: 'od:pptx-export-result', requestId, ok: true, buffer: new Uint8Array([80, 75, 3, 4]).buffer, warnings: [] } });
    await vi.advanceTimersByTimeAsync(1000);
    await expect(promise).resolves.toEqual({ warnings: [] });
    expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:export');
    expect(ports[0]!.close).toHaveBeenCalled();
    expect(document.querySelector('[data-od-pptx-export]')).toBeNull();
    expect(live.isConnected).toBe(true);
    expect(live.srcdoc).toBe('<p>Live</p>');
  });
  it('enforces mutual exclusion and cleans up at the timeout boundary', async () => {
    const removeListener = vi.spyOn(window, 'removeEventListener');
    const promise = exporter.exportDeckAsPptxInBrowser({ ...options, timeoutMs: 100 });
    const failure = expect(promise).rejects.toThrow('timed out');
    await flush();
    await expect(exporter.exportDeckAsPptxInBrowser(options)).rejects.toThrow('already running');
    await vi.advanceTimersByTimeAsync(99);
    expect(document.querySelector('[data-od-pptx-export]')).not.toBeNull();
    await vi.advanceTimersByTimeAsync(1);
    await failure;
    expect(document.querySelector('[data-od-pptx-export]')).toBeNull();
    expect(removeListener).toHaveBeenCalledWith('message', expect.any(Function));
  });
  it('cleans up after cancellation and malformed or failed replies', async () => {
    for (const kind of ['cancel', 'malformed', 'failure']) {
      const controller = new AbortController();
      const promise = exporter.exportDeckAsPptxInBrowser({ ...options, signal: controller.signal });
      const failure = expect(promise).rejects.toThrow();
      await flush();
      const frame = document.querySelector<HTMLIFrameElement>('[data-od-pptx-export]')!;
      vi.spyOn(frame.contentWindow!, 'postMessage').mockImplementation(() => {});
      if (kind === 'cancel') controller.abort(new Error('Cancelled'));
      else {
        ready(frame);
        ports.at(-1)!.onmessage!({ data: { type: 'od:pptx-export-result', requestId: frame.dataset.odPptxExport,
          ok: kind === 'malformed', buffer: 'invalid', error: 'Missing image' } });
      }
      await failure;
      expect(document.querySelector('[data-od-pptx-export]')).toBeNull();
    }
  });
  it('fails closed for missing bundles, wrong versions and runtime disable', async () => {
    vi.mocked(fetch).mockResolvedValue({ ok: false } as Response);
    expect(await exporter.clientPptxAvailable()).toBe(false);
    await expect(exporter.exportDeckAsPptxInBrowser(options)).rejects.toThrow('disabled');
    vi.mocked(fetch).mockImplementation(async (url) => String(url).endsWith('.json')
      ? { ok: true, json: async () => ({ enabled: true, version: CLIENT_PPTX_VERSION }) } as Response
      : { ok: true, text: async () => '<html>not a bundle</html>' } as Response);
    expect(await exporter.clientPptxAvailable()).toBe(false);
    await expect(exporter.exportDeckAsPptxInBrowser(options)).rejects.toThrow('version mismatch');
  });
  it('supports HTTP deployments without the secure-context-only randomUUID API', async () => {
    vi.stubGlobal('crypto', { getRandomValues: (bytes: Uint8Array) => bytes.fill(42) });
    expect(await exporter.clientPptxAvailable()).toBe(true);
    const controller = new AbortController();
    const promise = exporter.exportDeckAsPptxInBrowser({ ...options, signal: controller.signal });
    const failure = expect(promise).rejects.toThrow('Cancelled');
    await flush();
    const frame = document.querySelector<HTMLIFrameElement>('[data-od-pptx-export]')!;
    expect(frame.dataset.odPptxExport).toBe('2a'.repeat(16));
    controller.abort(new Error('Cancelled'));
    await failure;
  });
});
