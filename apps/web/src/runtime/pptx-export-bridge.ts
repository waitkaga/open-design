import { CLIENT_PPTX_MAX_BYTES, CLIENT_PPTX_MAX_SLIDES, isClientPptxRequest, redactPptxAssetUrl } from './client-pptx-protocol';
import { DECK_STAGE_SELECTOR, HIDE_CHROME_SELECTOR, SLIDE_SELECTOR, measureAuthoredSlideBox, prepareDeckStage, runDomToPptx, showAllSlides } from './pptx-export-normalizer';

export function installPptxExportBridge(): void {
  let started = false;
  let imageBytes = 0;
  const assetErrors = new Set<string>();
  window.addEventListener('error', (event) => {
    const target = event.target;
    if (target instanceof HTMLScriptElement || target instanceof HTMLLinkElement) {
      assetErrors.add(redactPptxAssetUrl(target instanceof HTMLScriptElement ? target.src : target.href));
    }
  }, true);

  const frames = () => new Promise<void>((resolve) => {
    // Chromium 会节流离屏 iframe 的动画帧；定时兜底避免静态布局永久等待。
    const timer = setTimeout(resolve, 100);
    requestAnimationFrame(() => requestAnimationFrame(() => { clearTimeout(timer); resolve(); }));
  });
  const loaded = new Promise<void>((resolve) => {
    if (document.readyState === 'complete') resolve();
    else window.addEventListener('load', () => resolve(), { once: true });
  });

  async function embedImage(image: HTMLImageElement): Promise<void> {
    image.loading = 'eager';
    const url = image.currentSrc || image.src;
    if (!url) return;
    try {
      if (!url.startsWith('data:')) {
        const response = await fetch(url, { signal: AbortSignal.timeout(15_000) });
        if (!response.ok) throw new Error('Image request failed');
        const blob = await response.blob();
        if (!blob.size || blob.size > 20 * 1024 * 1024) throw new Error('Image size limit');
        imageBytes += blob.size;
        if (imageBytes > 96 * 1024 * 1024) throw new Error('Deck asset size limit');
        const dataUrl = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result));
          reader.onerror = () => reject(new Error('Image read failed'));
          reader.readAsDataURL(blob);
        });
        image.removeAttribute('srcset');
        image.src = dataUrl;
      }
      await image.decode();
      if (!image.naturalWidth) throw new Error('Image decode failed');
    } catch {
      throw new Error(`Image unavailable: ${redactPptxAssetUrl(url)}`);
    }
  }

  async function exportDeck(fileName: string): Promise<{ buffer: ArrayBuffer; warnings: string[] }> {
    await loaded;
    await document.fonts.ready;
    if (assetErrors.size) throw new Error(`Stylesheet or script unavailable: ${[...assetErrors].join(', ')}`);
    const slides = [...document.querySelectorAll<HTMLElement>(SLIDE_SELECTOR)]
      .filter((slide) => !slide.closest('.mini-slide, .overview, .notes-overlay, .thumb'));
    if (!slides.length || slides.length > CLIENT_PPTX_MAX_SLIDES) {
      throw new Error(`Deck must contain 1-${CLIENT_PPTX_MAX_SLIDES} slides`);
    }
    if (slides.reduce((total, slide) => total + slide.querySelectorAll('*').length, 0) > 50_000) {
      throw new Error('Deck exceeds the 50,000 element limit');
    }
    prepareDeckStage(HIDE_CHROME_SELECTOR, DECK_STAGE_SELECTOR);
    showAllSlides(SLIDE_SELECTOR);
    const size = slides.map(measureAuthoredSlideBox).find(Boolean);
    if (!size || size.w < 100 || size.h < 100 || size.w > 8192 || size.h > 8192) {
      throw new Error('Invalid authored slide dimensions');
    }
    const warnings = new Set<string>();
    if (slides.length > 40) warnings.add('Large deck: browser export may require substantial memory.');
    for (const stage of document.querySelectorAll<HTMLElement>(DECK_STAGE_SELECTOR)) {
      stage.style.setProperty('width', `${size.w}px`, 'important');
      stage.style.setProperty('height', `${size.h}px`, 'important');
    }
    for (const slide of slides) {
      const authored = measureAuthoredSlideBox(slide);
      if (authored && (Math.abs(authored.w - size.w) > 1 || Math.abs(authored.h - size.h) > 1)) {
        throw new Error('Mixed slide dimensions are unsupported');
      }
      if (getComputedStyle(slide).display === 'none') slide.style.setProperty('display', 'block', 'important');
      slide.style.setProperty('width', `${size.w}px`, 'important');
      slide.style.setProperty('height', `${size.h}px`, 'important');
      slide.style.setProperty('transform', 'none', 'important');
      for (const image of slide.querySelectorAll<HTMLImageElement>('img')) await embedImage(image);
      for (const element of [slide, ...slide.querySelectorAll<HTMLElement>('*')]) {
        const style = getComputedStyle(element);
        if (style.filter !== 'none' || style.backdropFilter !== 'none'
          || style.mixBlendMode !== 'normal' || /gradient\(.*gradient\(/i.test(style.backgroundImage)) {
          warnings.add('Complex CSS effects may be rasterized or approximated.');
        }
        if (element instanceof HTMLIFrameElement) {
          warnings.add('Embedded iframe content is unsupported and is omitted.');
          element.style.display = 'none';
        }
        const background = style.backgroundImage;
        if (background.includes('url(')) {
          let rewritten = background;
          for (const match of background.matchAll(/url\(["']?([^"')]+)["']?\)/g)) {
            const image = new Image();
            image.src = new URL(match[1]!, document.baseURI).href;
            await embedImage(image);
            rewritten = rewritten.replace(match[0], `url("${image.src}")`);
          }
          element.style.setProperty('background-image', rewritten, 'important');
        }
      }
    }
    for (const animation of document.getAnimations()) {
      try { animation.finish(); } catch { animation.cancel(); }
    }
    await frames();
    const result = await runDomToPptx(SLIDE_SELECTOR, {}, 'export', [], fileName);
    if (result.error || !result.buffer?.byteLength || result.buffer.byteLength > CLIENT_PPTX_MAX_BYTES) {
      throw new Error(result.error || 'Invalid PPTX output');
    }
    if (result.buffer.byteLength > 32 * 1024 * 1024) warnings.add('Large PPTX file: over 32 MiB.');
    return { buffer: result.buffer, warnings: [...warnings] };
  }

  window.addEventListener('message', (event: MessageEvent) => {
    if (started || event.source !== parent || !isClientPptxRequest(event.data) || event.ports.length !== 1) return;
    started = true;
    const { requestId, fileName } = event.data;
    const port = event.ports[0]!;
    void exportDeck(fileName).then(({ buffer, warnings }) => {
      port.postMessage({ type: 'od:pptx-export-result', requestId, ok: true, buffer, warnings }, [buffer]);
    }).catch((error: unknown) => {
      port.postMessage({ type: 'od:pptx-export-result', requestId, ok: false,
        error: (error instanceof Error ? error.message : 'Browser PPTX export failed').slice(0, 2000) });
    }).finally(() => port.close());
  });
  parent.postMessage({ type: 'od:pptx-export-ready' }, '*');
}
