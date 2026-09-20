import { buildSrcdoc, PREVIEW_REDIRECT_LOOP_MESSAGE } from './srcdoc';
import { CLIENT_PPTX_TIMEOUT_MS, CLIENT_PPTX_VERSION, isClientPptxResult } from './client-pptx-protocol';

type ExportScripts = { bundle: string; bridge: string };
let cachedScripts: Promise<ExportScripts> | null = null;
let exporting = false;

function browserSupported(): boolean {
  return typeof window !== 'undefined' && /(?:Chrome|Chromium|Edg)\//.test(navigator.userAgent)
    && typeof MessageChannel === 'function' && typeof crypto.getRandomValues === 'function'
    && typeof AbortSignal.timeout === 'function';
}

async function loadScripts(): Promise<ExportScripts> {
  if (!cachedScripts) {
    cachedScripts = Promise.all(['dom-to-pptx.bundle.js', 'pptx-export-bridge.js'].map(async (file) => {
      const response = await fetch(`/vendor/${file}`, { credentials: 'same-origin', signal: AbortSignal.timeout(15_000) });
      if (!response.ok) throw new Error('Browser PPTX bundle unavailable');
      const text = await response.text();
      if (!text.startsWith(`// ${CLIENT_PPTX_VERSION}\n`)) throw new Error('Browser PPTX bundle version mismatch');
      return text;
    })).then(([bundle, bridge]) => ({ bundle: bundle!, bridge: bridge! })).catch((error: unknown) => {
      cachedScripts = null;
      throw error;
    });
  }
  return cachedScripts;
}

async function featureEnabled(): Promise<boolean> {
  const response = await fetch('/vendor/client-pptx.json', {
    credentials: 'same-origin', cache: 'no-store', signal: AbortSignal.timeout(10_000),
  });
  if (!response.ok) return false;
  const config: unknown = await response.json();
  return !!config && typeof config === 'object' && 'enabled' in config && config.enabled === true
    && 'version' in config && config.version === CLIENT_PPTX_VERSION;
}

export async function clientPptxAvailable(): Promise<boolean> {
  if (!browserSupported()) return false;
  try {
    if (!await featureEnabled()) return false;
    await loadScripts();
    return true;
  } catch { return false; }
}

export async function exportDeckAsPptxInBrowser(options: {
  sourceHtml: string;
  baseHref: string;
  fileName: string;
  signal?: AbortSignal;
  timeoutMs?: number;
}): Promise<{ warnings: string[] }> {
  if (exporting) throw new Error('A browser PPTX export is already running');
  if (!browserSupported()) throw new Error('Browser PPTX export requires Chrome or Edge');
  const base = new URL(options.baseHref, location.href);
  if (!/^https?:$/.test(base.protocol)) throw new Error('Invalid preview asset base URL');
  if (!options.sourceHtml || options.sourceHtml.length > 20 * 1024 * 1024) throw new Error('Invalid deck source size');
  exporting = true;
  const controller = new AbortController();
  const cancel = () => controller.abort(options.signal?.reason ?? new Error('Browser PPTX export cancelled'));
  const timer = setTimeout(() => controller.abort(new Error('Browser PPTX export timed out')), options.timeoutMs ?? CLIENT_PPTX_TIMEOUT_MS);
  options.signal?.addEventListener('abort', cancel, { once: true });
  if (options.signal?.aborted) cancel();
  let iframe: HTMLIFrameElement | null = null;
  let downloadUrl: string | null = null;
  let channel: MessageChannel | null = null;
  let onMessage: ((event: MessageEvent) => void) | null = null;
  let onAbort: (() => void) | null = null;
  try {
    const aborted = new Promise<never>((_, reject) => {
      onAbort = () => reject(controller.signal.reason);
      controller.signal.addEventListener('abort', onAbort, { once: true });
      if (controller.signal.aborted) onAbort();
    });
    const scripts = await Promise.race([(async () => {
      if (!await featureEnabled()) throw new Error('Browser PPTX export is disabled');
      return loadScripts();
    })(), aborted]);
    controller.signal.throwIfAborted();
    const requestId = Array.from(crypto.getRandomValues(new Uint8Array(16)), (byte) => byte.toString(16).padStart(2, '0')).join('');
    const fileName = (options.fileName.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').replace(/\.(html?|pptx)$/i, '').slice(0, 200) || 'deck') + '.pptx';
    iframe = document.createElement('iframe');
    iframe.dataset.odPptxExport = requestId;
    iframe.setAttribute('sandbox', 'allow-scripts allow-downloads');
    iframe.setAttribute('aria-hidden', 'true');
    iframe.tabIndex = -1;
    iframe.style.cssText = 'position:fixed;left:-10000px;top:0;width:1920px;height:1080px;border:0;pointer-events:none;';
    const frame = iframe;
    const result = new Promise<{ buffer: ArrayBuffer; warnings: string[] }>((resolve, reject) => {
      let connected = false;
      onMessage = (event) => {
        if (event.source !== frame.contentWindow) return;
        if (event.data?.type === PREVIEW_REDIRECT_LOOP_MESSAGE) {
          reject(new Error('Deck navigation blocked during export'));
          return;
        }
        if (connected || event.data?.type !== 'od:pptx-export-ready') return;
        connected = true;
        channel = new MessageChannel();
        channel.port1.onmessageerror = () => reject(new Error('Invalid browser PPTX message'));
        channel.port1.onmessage = (message) => {
          if (message.data?.requestId !== requestId) return;
          if (!isClientPptxResult(message.data, requestId)) {
            reject(new Error('Invalid browser PPTX result'));
          } else if (!message.data.ok) {
            reject(new Error(message.data.error));
          } else {
            const header = new Uint8Array(message.data.buffer, 0, Math.min(4, message.data.buffer.byteLength));
            if (header[0] !== 0x50 || header[1] !== 0x4b || header[2] !== 3 || header[3] !== 4) {
              reject(new Error('Browser PPTX result is not a ZIP archive'));
              return;
            }
            resolve({ buffer: message.data.buffer, warnings: message.data.warnings });
          }
        };
        channel.port1.start();
        frame.contentWindow?.postMessage({ type: 'od:pptx-export', requestId, fileName }, '*', [channel.port2]);
      };
      window.addEventListener('message', onMessage);
      frame.srcdoc = buildSrcdoc(options.sourceHtml, {
        baseHref: base.href, deck: true, hideDeckChrome: true, exportPptxScripts: scripts,
      });
      document.body.appendChild(frame);
    });
    const { buffer, warnings } = await Promise.race([result, aborted]);
    controller.signal.throwIfAborted();
    downloadUrl = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
    const anchor = document.createElement('a');
    anchor.href = downloadUrl;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    try { anchor.click(); } finally { anchor.remove(); }
    // 让浏览器的下载任务取得 URL，再统一释放所有导出资源。
    await new Promise<void>((resolve) => setTimeout(resolve, 1000));
    return { warnings };
  } finally {
    clearTimeout(timer);
    options.signal?.removeEventListener('abort', cancel);
    if (onAbort) controller.signal.removeEventListener('abort', onAbort);
    if (onMessage) window.removeEventListener('message', onMessage);
    const ports = channel as MessageChannel | null;
    ports?.port1.close();
    ports?.port2.close();
    iframe?.remove();
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    exporting = false;
  }
}
