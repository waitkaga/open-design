export const CLIENT_PPTX_VERSION = 'od-client-pptx:1;dom-to-pptx:2.0.1';
export const CLIENT_PPTX_MAX_BYTES = 128 * 1024 * 1024;
export const CLIENT_PPTX_MAX_SLIDES = 100;
export const CLIENT_PPTX_TIMEOUT_MS = 120_000;

export type ClientPptxExportRequest = {
  type: 'od:pptx-export';
  requestId: string;
  fileName: string;
};

export type ClientPptxExportResult =
  | { type: 'od:pptx-export-result'; requestId: string; ok: true; buffer: ArrayBuffer; warnings: string[] }
  | { type: 'od:pptx-export-result'; requestId: string; ok: false; error: string };

export function isClientPptxRequest(value: unknown): value is ClientPptxExportRequest {
  if (!value || typeof value !== 'object') return false;
  const data = value as Record<string, unknown>;
  return data.type === 'od:pptx-export' && typeof data.requestId === 'string'
    && data.requestId.length >= 16 && data.requestId.length <= 100
    && typeof data.fileName === 'string' && data.fileName.length <= 240;
}

export function isClientPptxResult(value: unknown, requestId: string): value is ClientPptxExportResult {
  if (!value || typeof value !== 'object') return false;
  const data = value as Record<string, unknown>;
  if (data.type !== 'od:pptx-export-result' || data.requestId !== requestId) return false;
  if (data.ok === false) return typeof data.error === 'string' && data.error.length <= 2000;
  return data.ok === true && data.buffer instanceof ArrayBuffer
    && data.buffer.byteLength > 0 && data.buffer.byteLength <= CLIENT_PPTX_MAX_BYTES
    && Array.isArray(data.warnings) && data.warnings.length <= 100
    && data.warnings.every((warning) => typeof warning === 'string' && warning.length <= 500);
}

export function pptxExportRoute(renderer: boolean | null, clientAvailable: boolean): 'server' | 'browser' | null {
  return renderer !== false ? 'server' : clientAvailable ? 'browser' : null;
}

export function redactPptxAssetUrl(value: string): string {
  try {
    const url = new URL(value);
    if (url.protocol === 'data:' || url.protocol === 'blob:') return `[${url.protocol} asset]`;
    // Scoped preview credentials may also appear in path segments.
    const name = url.pathname.split('/').pop() || 'asset';
    return `${url.origin}/.../${name.slice(0, 100)}`;
  } catch {
    return '[asset]';
  }
}
