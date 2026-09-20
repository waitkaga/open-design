// DOM-only port of desktop/deck-capture.ts at 73953213a6fec2c8092e8e77d229a3074aa828a9.
// Keep the authored geometry, heading, CJK and background algorithms aligned with that revision.
import { CLIENT_PPTX_MAX_BYTES, redactPptxAssetUrl } from './client-pptx-protocol';
export const HIDE_CHROME_SELECTOR =
  ".progress-bar, .notes-overlay, aside.notes, .speaker-notes, .deck-nav, .deck-hint, .deck-counter";

export const SLIDE_SELECTOR = ".slide, [data-screen-label], .deck-slide, .ppt-slide";

export const DECK_STAGE_SELECTOR = "deck-stage, #deck-stage, .deck-stage";

export type LayeredPptxBackgroundCapture = {
  dataUrl: string;
  height: number;
  left: number;
  slideIndex: number;
  top: number;
  width: number;
};

export function showAllSlides(slideSelector: string): number {
  const slides = Array.prototype.slice
    .call(document.querySelectorAll(slideSelector))
    .filter((el) => !(el as HTMLElement).closest(".mini-slide, .overview, .notes-overlay, .thumb"));
  for (const node of slides) {
    const el = node as HTMLElement;
    el.style.setProperty("opacity", "1", "important");
    el.style.setProperty("visibility", "visible", "important");
    el.style.setProperty("position", "absolute", "important");
    el.style.setProperty("left", "0", "important");
    el.style.setProperty("top", "0", "important");
    ["active", "visible", "is-active", "current"].forEach((c) => el.classList.add(c));
  }
  return slides.length;
}

export function cjkPromotedFontFamily(fontFamily: string, text: string): string | null {
  // CJK symbols/punctuation, Hiragana, Katakana, CJK Unified Ideographs (+ Ext-A),
  // Yi, Hangul syllables, CJK compatibility ideographs, and half/fullwidth forms.
  const cjkText =
    /[\u2E80-\u2FDF\u3000-\u303F\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uA000-\uA4CF\uAC00-\uD7AF\uF900-\uFAFF\uFF00-\uFFEF]/;
  // Family names that carry CJK glyph coverage: the Noto SC/TC/JP/KR webfonts the
  // html-ppt templates ship, plus common system CJK faces an authored deck may
  // name, so a promoted typeface resolves to a real CJK font across the office
  // suites instead of each app's arbitrary fallback.
  const cjkFamily =
    /noto\s*(sans|serif)\s*(sc|tc|hk|jp|kr|cjk)|source\s*han|pingfang|hiragino|heiti|songti|kaiti|fangsong|microsoft\s*(yahei|jhenghei)|yahei|simsun|simhei|mingliu|meiryo|ms\s*(gothic|mincho)|malgun|nanum|gulim|batang|dotum|思源|苹方|黑体|宋体|楷体|仿宋|微软雅黑|明體|明朝|ゴシック/i;
  if (!fontFamily || !cjkText.test(text || "")) return null;
  const families = fontFamily
    .split(",")
    .map((f) => f.trim())
    .filter(Boolean);
  if (families.length < 2) return null;
  const firstCjk = families.findIndex((f) => cjkFamily.test(f.replace(/^["']|["']$/g, "").trim()));
  // No CJK family to promote, or one already leads the stack.
  if (firstCjk <= 0) return null;
  return [families[firstCjk], ...families.filter((_, i) => i !== firstCjk)].join(", ");
}

export async function runDomToPptx(
  slideSelector: string,
  layeredBackgrounds: Record<string, LayeredPptxBackgroundCapture> = {},
  phase: "export" | "prepare" | "export-prepared" = "export",
  importedStylesheetOverrides: Array<{ cssText: string; url: string }> = [],
  fileName = 'deck.pptx',
): Promise<{ buffer?: ArrayBuffer; error?: string; prepared?: boolean }> {
  // dom-to-pptx fixes native ::before content at -1,000,000. Reserve the two
  // preceding slots for its raster background and the slide background below
  // it so a slide-root pseudo remains visible over an opaque slide fill.
  const slideBackgroundSortSlot = "-1000002";
  const pseudoBeforeBackgroundSortSlot = "-1000001";
  const pseudoAfterBackgroundSortSlot = "0";
  function importedStylesheetUrls(cssText: string, baseHref: string): string[] {
    const urls: string[] = [];
    const importPattern =
      /@import\s+(?:url\(\s*)?(?:(["'])([\s\S]*?)\1|([^"')\s;]+))\s*\)?[^;]*;/giu;
    for (const match of cssText.matchAll(importPattern)) {
      const raw = match[2] || match[3];
      if (!raw) continue;
      try {
        urls.push(new URL(raw, baseHref).href);
      } catch {
        // Ignore malformed author CSS and let the existing font fallback apply.
      }
    }
    return urls;
  }

  function importedFontFaceCss(cssText: string, baseHref: string): string {
    const faces = (cssText.match(/@font-face\s*\{[\s\S]*?\}/giu) || []).map((rule) => {
      const value = (property: string): string =>
        rule.match(new RegExp(`${property}\\s*:\\s*([^;]+)`, "iu"))?.[1]?.trim() || "";
      return {
        family: value("font-family").replace(/^['"]|['"]$/g, ""),
        rule,
        style: value("font-style").toLowerCase() || "normal",
        unicodeRange: value("unicode-range"),
        weight: value("font-weight").toLowerCase() || "400",
      };
    });
    const preferredFace = new Map<string, { rank: number; style: string; weight: string }>();
    for (const face of faces) {
      const rank = face.style === "normal" ? (face.weight === "400" || face.weight === "normal" ? 0 : 1) : 2;
      const current = preferredFace.get(face.family);
      if (!current || rank < current.rank) {
        preferredFace.set(face.family, { rank, style: face.style, weight: face.weight });
      }
    }

    const preferredRule = new Map<string, { rank: number; rule: string }>();
    for (const face of faces) {
      const preferred = preferredFace.get(face.family);
      if (preferred?.style !== face.style || preferred.weight !== face.weight) continue;
      // Google Fonts commonly returns one @font-face per unicode subset. The
      // vendored converter can fail while merging some families' subsets, so
      // prefer the complete face when present, then its Latin core subset.
      const rank = face.unicodeRange === "" ? 0 : /U\+0000-00FF/iu.test(face.unicodeRange) ? 1 : 2;
      const current = preferredRule.get(face.family);
      if (!current || rank < current.rank) preferredRule.set(face.family, { rank, rule: face.rule });
    }

    return faces
      .filter((face) => preferredRule.get(face.family)?.rule === face.rule)
      .map((rule) =>
        rule.rule.replace(/url\(\s*(["']?)([^"')]+)\1\s*\)/giu, (_match, _quote, raw: string) => {
          try {
            return `url("${new URL(raw.trim(), baseHref).href}")`;
          } catch {
            return `url("${raw.trim()}")`;
          }
        }),
      )
      .join("\n");
  }

  // dom-to-pptx's autoEmbedFonts scanner sees top-level CSSFontFaceRule entries,
  // but many OpenDesign decks load Google Fonts through an inline `@import`.
  // Expand those imports into a throwaway top-level style so the vendored engine
  // can discover and embed the actual font files instead of only writing their
  // family names into the PPTX. The render window is destroyed after export, so
  // this never mutates the authored HTML or the live preview.
  async function exposeImportedFontFaces(): Promise<Array<{ name: string; urls: string[] }>> {
    const importedUrls = new Set<string>();
    document.querySelectorAll("style").forEach((style) => {
      for (const url of importedStylesheetUrls(style.textContent || "", document.baseURI)) {
        importedUrls.add(url);
      }
    });
    document.querySelectorAll<HTMLLinkElement>('link[rel="stylesheet"]').forEach((link) => {
      if (link.href) importedUrls.add(link.href);
    });

    const visited = new Set<string>();
    const fontFaceRules: string[] = [];
    const collect = async (url: string): Promise<void> => {
      if (visited.has(url)) return;
      visited.add(url);
      try {
        const override = importedStylesheetOverrides.find((entry) => entry.url === url);
        const response = override ? null : await fetch(url, { signal: AbortSignal.timeout(15_000) });
        if (response && !response.ok) throw new Error(`HTTP ${response.status}`);
        const cssText = override?.cssText ?? (await response!.text());
        for (const nested of importedStylesheetUrls(cssText, url)) await collect(nested);
        const fontCss = importedFontFaceCss(cssText, url);
        if (fontCss) fontFaceRules.push(fontCss);
      } catch {
        throw new Error(`Font stylesheet unavailable: ${redactPptxAssetUrl(url)}`);
      }
    };
    for (const url of importedUrls) await collect(url);
    document.querySelectorAll('style').forEach((style) => {
      const css = importedFontFaceCss(style.textContent || '', document.baseURI);
      if (css) fontFaceRules.push(css);
    });
    if (fontFaceRules.length === 0) return [];

    const combinedCss = fontFaceRules.join("\n");
    const style = document.createElement("style");
    style.setAttribute("data-od-pptx-imported-font-faces", "true");
    style.textContent = combinedCss;
    document.head.appendChild(style);

    const fontsByFamily = new Map<string, Set<string>>();
    for (const rule of combinedCss.match(/@font-face\s*\{[\s\S]*?\}/giu) || []) {
      const family = rule
        .match(/font-family\s*:\s*([^;]+)/iu)?.[1]
        ?.trim()
        .replace(/^['"]|['"]$/g, "");
      if (!family) continue;
      const urls = fontsByFamily.get(family) || new Set<string>();
      for (const match of rule.matchAll(/url\(\s*["']?([^"')]+)["']?\s*\)/giu)) {
        if (match[1]) urls.add(match[1]);
      }
      if (urls.size > 0) fontsByFamily.set(family, urls);
    }
    const fonts = Array.from(fontsByFamily, ([name, urls]) => ({ name, urls: Array.from(urls) }));
    for (const font of fonts) {
      for (const url of font.urls) {
        try {
          const response = await fetch(url, { signal: AbortSignal.timeout(15_000) });
          if (!response.ok || (await response.arrayBuffer()).byteLength === 0) throw new Error('Empty font');
        } catch {
          throw new Error(`Font unavailable: ${font.name.slice(0, 100)} (${redactPptxAssetUrl(url)})`);
        }
      }
    }
    return fonts;
  }

  function isTransparentColor(input: string): boolean {
    const value = input.trim().toLowerCase();
    return value === "" || value === "transparent" || value === "rgba(0, 0, 0, 0)";
  }

  function firstCssColor(input: string): string | null {
    const rgb = input.match(/rgba?\([^)]*\)/i);
    if (rgb) return rgb[0];
    const hex = input.match(/#[0-9a-f]{3,8}\b/i);
    return hex ? hex[0] : null;
  }

  function effectiveBackgroundStyle(slide: HTMLElement): {
    color: string;
    image: string;
    position: string;
    size: string;
    repeat: string;
    origin: string;
    clip: string;
  } | null {
    const candidates: Element[] = [];
    for (let el: Element | null = slide; el; el = el.parentElement) candidates.push(el);
    if (document.body && !candidates.includes(document.body)) candidates.push(document.body);
    if (document.documentElement && !candidates.includes(document.documentElement)) {
      candidates.push(document.documentElement);
    }

    for (const el of candidates) {
      const style = getComputedStyle(el);
      const bgColor = style.backgroundColor;
      const bgImage = style.backgroundImage;
      const hasImage = bgImage && bgImage !== "none";
      const hasColor = bgColor && !isTransparentColor(bgColor);
      const fallbackColor = hasColor ? bgColor : firstCssColor(bgImage);
      if (!hasImage && !hasColor) continue;
      if (!fallbackColor) continue;
      return {
        color: fallbackColor,
        image: bgImage,
        position: style.backgroundPosition,
        size: style.backgroundSize,
        repeat: style.backgroundRepeat,
        origin: style.backgroundOrigin,
        clip: style.backgroundClip,
      };
    }
    return null;
  }

  function ensureExplicitSlideBackgrounds(slides: HTMLElement[]): void {
    for (const slide of slides) {
      slide.querySelectorAll(":scope > [data-od-pptx-bg]").forEach((el) => el.remove());
      // preserveLayeredGradientBackgrounds owns supported layered backgrounds
      // authored directly on a slide. Adding the usual fallback shim as well
      // would export the same semi-transparent texture twice.
      if (hasRasterizableLayeredGradientBackground(getComputedStyle(slide).backgroundImage || "")) {
        continue;
      }
      const background = effectiveBackgroundStyle(slide);
      if (!background) continue;

      const bg = document.createElement("div");
      bg.setAttribute("data-od-pptx-bg", "true");
      bg.setAttribute("aria-hidden", "true");
      bg.style.setProperty("position", "absolute", "important");
      bg.style.setProperty("inset", "0", "important");
      bg.style.setProperty("z-index", slideBackgroundSortSlot, "important");
      bg.style.setProperty("pointer-events", "none", "important");
      bg.style.setProperty("background-color", background.color, "important");
      bg.style.setProperty("background-image", background.image, "important");
      bg.style.setProperty("background-position", background.position, "important");
      bg.style.setProperty("background-size", background.size, "important");
      bg.style.setProperty("background-repeat", background.repeat, "important");
      bg.style.setProperty("background-origin", background.origin, "important");
      bg.style.setProperty("background-clip", background.clip, "important");

      const style = getComputedStyle(slide);
      if (style.position === "static") slide.style.setProperty("position", "relative", "important");
      if (style.overflow === "visible") slide.style.setProperty("overflow", "hidden", "important");
      slide.style.setProperty("background-color", background.color, "important");
      Array.from(slide.children).forEach((child) => {
        if (child.getAttribute("data-od-pptx-bg") === "true") return;
        const childStyle = getComputedStyle(child as Element);
        const element = child as HTMLElement;
        if (childStyle.position === "static") {
          element.style.setProperty("position", "relative", "important");
        }
        if (childStyle.zIndex === "auto") {
          element.style.setProperty("z-index", "1", "important");
        }
      });
      slide.prepend(bg);
    }
  }

  function splitCssBackgroundLayers(input: string): string[] {
    const layers: string[] = [];
    let current = "";
    let depth = 0;
    let quote = "";
    let escaped = false;
    for (const char of input) {
      if (escaped) {
        current += char;
        escaped = false;
        continue;
      }
      if (char === "\\") {
        current += char;
        escaped = true;
        continue;
      }
      if (quote) {
        current += char;
        if (char === quote) quote = "";
        continue;
      }
      if (char === '"' || char === "'") {
        current += char;
        quote = char;
        continue;
      }
      if (char === "(") depth += 1;
      else if (char === ")") depth = Math.max(0, depth - 1);
      if (char === "," && depth === 0) {
        if (current.trim()) layers.push(current.trim());
        current = "";
      } else {
        current += char;
      }
    }
    if (current.trim()) layers.push(current.trim());
    return layers;
  }

  function hasRasterizableLayeredGradientBackground(input: string): boolean {
    const layers = splitCssBackgroundLayers(input);
    if (layers.length < 2) return false;
    // Keep this allowlist aligned with html2canvas 1.4.1's
    // SUPPORTED_IMAGE_FUNCTIONS. In particular, repeating and conic gradients
    // are discarded by its clone parser and must remain on the authored node.
    const supportedGradient =
      /^(?:(?:-(?:moz|ms|o|webkit)-)?(?:linear|radial)-gradient|-webkit-gradient)\(/i;
    return layers.every((layer) => supportedGradient.test(layer));
  }

  function hasTextBackgroundClip(input: string): boolean {
    return splitCssBackgroundLayers(input).some((layer) => layer.toLowerCase() === "text");
  }

  function hasNonNormalBlendMode(input: string): boolean {
    const mode = (input || "normal").trim().toLowerCase();
    return mode !== "" && mode !== "normal";
  }

  function hasBackdropFilter(style: CSSStyleDeclaration): boolean {
    const value = (
      style.backdropFilter ||
      style.getPropertyValue?.("backdrop-filter") ||
      style.getPropertyValue?.("-webkit-backdrop-filter") ||
      "none"
    ).trim().toLowerCase();
    return value !== "" && value !== "none";
  }

  function hasCssMask(style: CSSStyleDeclaration): boolean {
    const maskImages = [
      style.maskImage || style.getPropertyValue("mask-image"),
      style.webkitMaskImage || style.getPropertyValue("-webkit-mask-image"),
    ];
    return maskImages.some((image) => image && image.trim().toLowerCase() !== "none");
  }

  function setCaptureBoxStyles(background: HTMLElement, style: CSSStyleDeclaration): void {
    background.style.setProperty("box-sizing", "border-box", "important");
    background.style.setProperty(
      "padding",
      `${style.paddingTop || "0px"} ${style.paddingRight || "0px"} ${style.paddingBottom || "0px"} ${style.paddingLeft || "0px"}`,
      "important",
    );
    background.style.setProperty(
      "border-width",
      `${style.borderTopWidth || "0px"} ${style.borderRightWidth || "0px"} ${style.borderBottomWidth || "0px"} ${style.borderLeftWidth || "0px"}`,
      "important",
    );
    background.style.setProperty("border-style", "solid", "important");
    background.style.setProperty("border-color", "transparent", "important");
    background.style.setProperty("border-radius", style.borderRadius || "0px", "important");
    background.style.setProperty("box-shadow", style.boxShadow || "none", "important");
    background.style.setProperty("background-color", style.backgroundColor, "important");
    background.style.setProperty("background-image", style.backgroundImage, "important");
    background.style.setProperty("background-position", style.backgroundPosition, "important");
    background.style.setProperty("background-size", style.backgroundSize, "important");
    background.style.setProperty("background-repeat", style.backgroundRepeat, "important");
    background.style.setProperty("background-origin", style.backgroundOrigin, "important");
    background.style.setProperty("background-clip", style.backgroundClip, "important");
    background.style.setProperty("background-blend-mode", style.backgroundBlendMode || "normal", "important");
    background.style.setProperty("clip-path", style.clipPath || "none", "important");
    background.style.setProperty("filter", style.filter || "none", "important");
    const backdropFilter =
      style.backdropFilter ||
      style.getPropertyValue?.("backdrop-filter") ||
      style.getPropertyValue?.("-webkit-backdrop-filter") ||
      "none";
    background.style.setProperty("backdrop-filter", backdropFilter, "important");
    background.style.setProperty("-webkit-backdrop-filter", backdropFilter, "important");
    background.style.setProperty("opacity", style.opacity || "1", "important");
    background.style.setProperty("mix-blend-mode", style.mixBlendMode || "normal", "important");
    background.style.setProperty("transform", style.transform || "none", "important");
    background.style.setProperty("transform-origin", style.transformOrigin || "50% 50%", "important");
    background.style.setProperty("transform-box", style.transformBox || "view-box", "important");
    background.style.setProperty("translate", style.translate || "none", "important");
    background.style.setProperty("rotate", style.rotate || "none", "important");
    background.style.setProperty("scale", style.scale || "none", "important");
  }

  function preserveLayeredPseudoGradientBackgrounds(elements: Set<HTMLElement>): void {
    let nativePseudoBackgroundStyle: HTMLStyleElement | null = null;
    const neutralizeNativePseudoBackground = (
      element: HTMLElement,
      pseudo: "::before" | "::after",
    ): void => {
      element.setAttribute(
        pseudo === "::before"
          ? "data-od-pptx-rasterized-before-background"
          : "data-od-pptx-rasterized-after-background",
        "true",
      );
      if (nativePseudoBackgroundStyle) return;
      nativePseudoBackgroundStyle = document.createElement("style");
      nativePseudoBackgroundStyle.textContent = `
        [data-od-pptx-rasterized-before-background="true"]::before,
        [data-od-pptx-rasterized-after-background="true"]::after{
          background-color:transparent!important;
        }
      `;
      document.head.append(nativePseudoBackgroundStyle);
    };
    for (const element of elements) {
      for (const pseudo of ["::before", "::after"] as const) {
        const style = getComputedStyle(element, pseudo);
        const content = (style.content || "").trim().toLowerCase();
        const isGenerated = content !== "" && content !== "none" && content !== "normal" && style.display !== "none";
        const hasMaterializedCapture = Array.from(element.children).some(
          (child) => child.getAttribute("data-od-pptx-materialized-pseudo") === pseudo,
        );
        if (hasMaterializedCapture) {
          // The Chromium helper already owns the computed fallback color. Keep
          // native pseudo text and borders, but prevent dom-to-pptx from
          // emitting that same color as an opaque fill above the captured PNG.
          neutralizeNativePseudoBackground(element, pseudo);
          continue;
        }
        if (
          !isGenerated ||
          (style.position !== "absolute" && style.position !== "fixed") ||
          !hasRasterizableLayeredGradientBackground(style.backgroundImage || "") ||
          // The html2canvas custom-element path has no blend-mode parser and
          // cannot reproduce this background without its authored backdrop.
          hasNonNormalBlendMode(style.mixBlendMode || "") ||
          hasBackdropFilter(style) ||
          hasTextBackgroundClip(style.backgroundClip || "") ||
          hasTextBackgroundClip(style.webkitBackgroundClip || "") ||
          hasCssMask(style)
        ) {
          continue;
        }

        // dom-to-pptx only reads pseudo-element content, color, and border. A
        // background-only custom element enters its existing html2canvas path,
        // preserving the layered image while the native pseudo handling keeps
        // any authored text or border editable.
        const background = document.createElement("od-pptx-layered-background");
        background.setAttribute("data-od-pptx-layered-bg", "true");
        background.setAttribute("data-od-pptx-pseudo", pseudo);
        background.setAttribute("aria-hidden", "true");
        background.style.setProperty("position", style.position, "important");
        background.style.setProperty("top", style.top || "auto", "important");
        background.style.setProperty("right", style.right || "auto", "important");
        background.style.setProperty("bottom", style.bottom || "auto", "important");
        background.style.setProperty("left", style.left || "auto", "important");
        background.style.setProperty("width", style.width || "auto", "important");
        background.style.setProperty("height", style.height || "auto", "important");
        // Keep the raster background immediately below the converter's fixed
        // native pseudo text/border slots. Native ::after always sorts at the
        // host's z=0 Infinity slot, regardless of its authored z-index.
        background.style.setProperty(
          "z-index",
          pseudo === "::before" ? pseudoBeforeBackgroundSortSlot : pseudoAfterBackgroundSortSlot,
          "important",
        );
        background.style.setProperty("pointer-events", "none", "important");
        setCaptureBoxStyles(background, style);

        // The converter keeps pseudo content and borders editable, but it also
        // emits a native solid fill from background-color while ignoring the
        // layered background-image. The raster helper already owns both, so
        // neutralize only that native fallback after copying its computed color.
        neutralizeNativePseudoBackground(element, pseudo);

        if (pseudo === "::before") element.prepend(background);
        else element.append(background);
      }
    }
  }

  function suppressCapturedSlidePaint(slide: HTMLElement, capture: HTMLElement): void {
    // dom-to-pptx needs the slide itself to remain measurable as the export
    // root. Keep only the replacement image visible inside it and neutralize
    // effects that Chromium already baked into that whole-paint capture.
    slide.querySelectorAll<HTMLElement>("*").forEach((descendant) => {
      if (descendant !== capture && !capture.contains(descendant)) {
        descendant.style.setProperty("display", "none", "important");
      }
    });
    slide.style.setProperty("background", "transparent", "important");
    slide.style.setProperty("border", "0", "important");
    slide.style.setProperty("box-shadow", "none", "important");
    slide.style.setProperty("clip-path", "none", "important");
    slide.style.setProperty("color", "transparent", "important");
    slide.style.setProperty("filter", "none", "important");
    slide.style.setProperty("backdrop-filter", "none", "important");
    slide.style.setProperty("-webkit-backdrop-filter", "none", "important");
    slide.style.setProperty("mask-image", "none", "important");
    slide.style.setProperty("-webkit-mask-image", "none", "important");
    slide.style.setProperty("mix-blend-mode", "normal", "important");
    slide.style.setProperty("opacity", "1", "important");
    slide.style.setProperty("outline", "none", "important");
    slide.style.setProperty("text-shadow", "none", "important");
    slide.style.setProperty("-webkit-text-fill-color", "transparent", "important");
    slide.style.setProperty("transform", "none", "important");
    slide.style.setProperty("translate", "none", "important");
    slide.style.setProperty("rotate", "none", "important");
    slide.style.setProperty("scale", "none", "important");
  }

  function preserveLayeredGradientBackgrounds(slides: HTMLElement[]): void {
    if (document.querySelectorAll("[data-od-pptx-suppress-before], [data-od-pptx-suppress-after]").length > 0) {
      const suppressedPseudoStyle = document.createElement("style");
      suppressedPseudoStyle.textContent = `
        [data-od-pptx-suppress-before="true"]::before,
        [data-od-pptx-suppress-after="true"]::after{
          content:none!important;
          display:none!important;
          border:0!important;
          background:none!important;
        }
      `;
      document.head.append(suppressedPseudoStyle);
    }
    const slideElements = new Set(slides);
    const elements = new Set<HTMLElement>();
    for (const slide of slides) {
      elements.add(slide);
      slide.querySelectorAll<HTMLElement>("*").forEach((el) => elements.add(el));
    }

    const capturedCompositingMembers = new Set<HTMLElement>();
    const capturedEntirePaintRoots = new Set<HTMLElement>();
    for (const element of elements) {
      if (element.getAttribute("data-od-pptx-compositing-context") !== "true") continue;
      const captureId = element.getAttribute("data-od-pptx-layer-capture-id") || "";
      const captured = layeredBackgrounds[captureId];
      if (!captured) continue;
      const slide = slides[captured.slideIndex];
      if (!slide) continue;

      const style = getComputedStyle(element);
      // Export the flattened context beside its source. Ordinary members lose
      // only the backgrounds already present in the PNG; members whose own
      // compositor effect required whole-paint capture are suppressed entirely.
      const image = document.createElement("img");
      image.setAttribute("data-od-pptx-layered-bg", "true");
      image.setAttribute("aria-hidden", "true");
      image.src = captured.dataUrl;
      image.style.setProperty("position", "absolute", "important");
      image.style.setProperty("left", `${captured.left}px`, "important");
      image.style.setProperty("top", `${captured.top}px`, "important");
      image.style.setProperty("width", `${captured.width}px`, "important");
      image.style.setProperty("height", `${captured.height}px`, "important");
      image.style.setProperty("display", "block", "important");
      image.style.setProperty("object-fit", "fill", "important");
      image.style.setProperty("pointer-events", "none", "important");
      image.style.setProperty("z-index", style.zIndex || "auto", "important");
      image.getBoundingClientRect = () => {
        const slideRect = slide.getBoundingClientRect();
        const left = slideRect.left + captured.left;
        const top = slideRect.top + captured.top;
        return {
          bottom: top + captured.height,
          height: captured.height,
          left,
          right: left + captured.width,
          top,
          width: captured.width,
          x: left,
          y: top,
          toJSON: () => ({}),
        } as DOMRect;
      };
      if (element === slide) slide.prepend(image);
      else element.parentElement?.insertBefore(image, element);
      document
        .querySelectorAll<HTMLElement>(`[data-od-pptx-compositing-member="${captureId}"]`)
        .forEach((member) => {
          if (
            member.hasAttribute("data-od-pptx-materialized-pseudo") ||
            member.hasAttribute("data-od-pptx-capture-entire-element")
          ) {
            if (member === slide) suppressCapturedSlidePaint(member, image);
            else member.style.setProperty("display", "none", "important");
            capturedEntirePaintRoots.add(member);
          } else {
            member.style.setProperty("background-image", "none", "important");
            member.style.setProperty("background-color", "transparent", "important");
          }
          capturedCompositingMembers.add(member);
        });
    }

    for (const element of elements) {
      if (capturedCompositingMembers.has(element)) continue;
      if (Array.from(capturedEntirePaintRoots).some((root) => root.contains(element))) continue;
      const style = getComputedStyle(element);
      const captureId = element.getAttribute("data-od-pptx-layer-capture-id") || "";
      const captured = layeredBackgrounds[captureId];
      if (
        !hasRasterizableLayeredGradientBackground(style.backgroundImage || "") ||
        (!captured && (
          hasTextBackgroundClip(style.backgroundClip || "") ||
          hasTextBackgroundClip(style.webkitBackgroundClip || "")
        )) ||
        // The custom-element fallback uses html2canvas, which cannot preserve
        // masks. Production exports provide a Chromium capture for these.
        (hasCssMask(style) && !captured)
      ) {
        continue;
      }
      const isStaticNestedElement = style.position === "static" && !slideElements.has(element);

      if (captured) {
        const slide = slides[captured.slideIndex];
        if (!slide) continue;
        const materializedPseudo = element.getAttribute("data-od-pptx-materialized-pseudo");
        const capturesEntireElement = element.getAttribute("data-od-pptx-capture-entire-element") === "true";
        const background = document.createElement("img");
        background.setAttribute("data-od-pptx-layered-bg", "true");
        if (materializedPseudo) background.setAttribute("data-od-pptx-pseudo", materializedPseudo);
        background.setAttribute("aria-hidden", "true");
        background.src = captured.dataUrl;
        background.style.setProperty("position", "absolute", "important");
        background.style.setProperty("left", `${captured.left}px`, "important");
        background.style.setProperty("top", `${captured.top}px`, "important");
        background.style.setProperty("width", `${captured.width}px`, "important");
        background.style.setProperty("height", `${captured.height}px`, "important");
        background.style.setProperty("display", "block", "important");
        background.style.setProperty("object-fit", "fill", "important");
        background.style.setProperty("pointer-events", "none", "important");
        background.style.setProperty(
          "z-index",
          element === slide
            ? slideBackgroundSortSlot
            : materializedPseudo === "::before"
              ? pseudoBeforeBackgroundSortSlot
              : materializedPseudo === "::after"
                ? pseudoAfterBackgroundSortSlot
                : style.zIndex || "auto",
          "important",
        );
        background.getBoundingClientRect = () => {
          const slideRect = slide.getBoundingClientRect();
          const left = slideRect.left + captured.left;
          const top = slideRect.top + captured.top;
          return {
            bottom: top + captured.height,
            height: captured.height,
            left,
            right: left + captured.width,
            top,
            width: captured.width,
            x: left,
            y: top,
            toJSON: () => ({}),
          } as DOMRect;
        };
        element.style.setProperty("background-image", "none", "important");
        element.style.setProperty("background-color", "transparent", "important");
        if (element === slide) slide.prepend(background);
        else element.parentElement?.insertBefore(background, element);
        if (materializedPseudo || capturesEntireElement) {
          // The helper exists only to give Chromium a real capture target. Its
          // raster image now owns that paint; leaving the custom element in the
          // converter walk would emit the same pseudo as a second media layer.
          if (element === slide) suppressCapturedSlidePaint(element, background);
          else element.style.setProperty("display", "none", "important");
          capturedEntirePaintRoots.add(element);
        }
        continue;
      }

      // dom-to-pptx's native gradient parser assumes one linear-gradient and
      // greedily merges layered gradients into one invalid SVG. In test-only
      // callers without the main-process capture seam, retain the existing
      // custom-element fallback for unmasked layers.
      const background = document.createElement("od-pptx-layered-background");
      background.setAttribute("data-od-pptx-layered-bg", "true");
      background.setAttribute("aria-hidden", "true");
      background.style.setProperty("position", "absolute", "important");
      background.style.setProperty("inset", "0", "important");
      background.style.setProperty(
        "z-index",
        slideElements.has(element) ? slideBackgroundSortSlot : "0",
        "important",
      );
      background.style.setProperty("pointer-events", "none", "important");
      setCaptureBoxStyles(background, style);

      if (isStaticNestedElement) {
        // A static panel and its absolutely positioned descendants share the
        // same outer containing block. Anchor only the capture child to the
        // panel's measured border box so the authored panel never becomes a
        // new containing block.
        background.style.setProperty("inset", "auto", "important");
        background.style.setProperty("left", `${element.offsetLeft}px`, "important");
        background.style.setProperty("top", `${element.offsetTop}px`, "important");
        background.style.setProperty("width", `${element.offsetWidth}px`, "important");
        background.style.setProperty("height", `${element.offsetHeight}px`, "important");
      } else {
        // ensureExplicitSlideBackgrounds already establishes this containing-
        // block contract for slides; positioned authored elements already own
        // the absolutely positioned capture child.
        if (style.position === "static") element.style.setProperty("position", "relative", "important");
      }

      element.style.setProperty("background-image", "none", "important");
      element.style.setProperty("background-color", "transparent", "important");
      element.prepend(background);
    }

    preserveLayeredPseudoGradientBackgrounds(elements);
  }

  function stabilizeLargeSingleLineText(slides: HTMLElement[]): void {
    for (const slide of slides) {
      slide.querySelectorAll<HTMLElement>("*").forEach((el) => {
        const rawText = el.innerText || el.textContent || "";
        const text = rawText.replace(/\s+/g, " ").trim();
        if (!text || rawText.includes("\n")) return;

        const style = getComputedStyle(el);
        const fontSizePx = Number.parseFloat(style.fontSize);
        if (!Number.isFinite(fontSizePx) || fontSizePx < 96) return;

        const lineHeightPx = Number.parseFloat(style.lineHeight);
        if (!Number.isFinite(lineHeightPx) || lineHeightPx <= 0 || lineHeightPx > fontSizePx * 1.05) return;

        const rect = el.getBoundingClientRect();
        if (rect.width <= 1 || rect.height <= 1) return;

        const justify =
          style.textAlign === "center" || style.textAlign === "-webkit-center"
            ? "center"
            : style.textAlign === "right" || style.textAlign === "end"
              ? "flex-end"
              : "flex-start";

        el.style.setProperty("display", "flex", "important");
        el.style.setProperty("align-items", "center", "important");
        el.style.setProperty("justify-content", justify, "important");
        el.style.setProperty("width", `${rect.width}px`, "important");
        el.style.setProperty("height", `${rect.height}px`, "important");
        el.style.setProperty("line-height", "normal", "important");
        el.style.setProperty("white-space", "nowrap", "important");
        el.style.setProperty("overflow", "visible", "important");
      });
    }
  }

  // An authored `<br>` is a deliberate line boundary. Prevent PowerPoint/WPS
  // from applying a second soft wrap inside either line when its font metrics
  // differ slightly from Chromium's. dom-to-pptx maps `white-space: nowrap` to
  // `wrap: false` while retaining explicit breakLine runs.
  function stabilizeAuthoredHeadingLines(slides: HTMLElement[]): void {
    for (const slide of slides) {
      slide.querySelectorAll<HTMLElement>("h1, h2, h3").forEach((heading) => {
        if (heading.querySelector("br")) {
          heading.style.setProperty("white-space", "nowrap", "important");
        }
      });
    }
  }

  // Reorder each text run's font-family so CJK runs name their CJK typeface (not
  // the Latin webfont that leads our template stacks) before dom-to-pptx reads it,
  // so PowerPoint/WPS/Keynote all resolve the same real font. See
  // cjkPromotedFontFamily for the why. Keyed on the element that directly owns the
  // text so a container that only holds Latin markup is never rewritten. Decide on
  // the element's COMBINED direct text: bilingual markup often splits one element
  // across text nodes (`Product Launch<br>产品发布`, `Welcome <strong>…</strong> 欢迎`),
  // so a later CJK chunk must still win even when a Latin chunk comes first.
  function promoteCjkTypefaces(slides: HTMLElement[]): void {
    const touched = new Set<HTMLElement>();
    for (const slide of slides) {
      const walker = document.createTreeWalker(slide, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const el = node.parentElement;
        if (!el || touched.has(el)) continue;
        touched.add(el);
        let combined = "";
        for (const child of el.childNodes) {
          if (child.nodeType === Node.TEXT_NODE) combined += child.nodeValue || "";
        }
        if (!combined.trim()) continue;
        const promoted = cjkPromotedFontFamily(getComputedStyle(el).fontFamily, combined);
        if (promoted) el.style.setProperty("font-family", promoted, "important");
      }
    }
  }

  try {
    const w = window as unknown as {
      domToPptx?: { exportToPptx: (target: unknown, options: unknown) => Promise<Blob> };
    };
    if (!w.domToPptx || typeof w.domToPptx.exportToPptx !== "function") {
      return { error: "dom-to-pptx engine did not load" };
    }
    const slides = Array.prototype.slice
      .call(document.querySelectorAll(slideSelector))
      .filter((el) => !(el as HTMLElement).closest(".mini-slide, .overview, .notes-overlay, .thumb"));
    if (slides.length === 0) return { error: "no slides to export" };
    const importedFonts = await exposeImportedFontFaces();
    await document.fonts?.ready;
    if (phase !== "export-prepared") {
      ensureExplicitSlideBackgrounds(slides as HTMLElement[]);
      stabilizeLargeSingleLineText(slides as HTMLElement[]);
      stabilizeAuthoredHeadingLines(slides as HTMLElement[]);
      promoteCjkTypefaces(slides as HTMLElement[]);
      // dom-to-pptx assumes `node.className` is a string, but SVG elements expose
      // an SVGAnimatedString, so its DOM walk throws on decks containing inline SVG.
      // Normalize those to a plain string in this throwaway render window.
      document.querySelectorAll("*").forEach((el) => {
        const cn = (el as { className?: unknown }).className;
        if (cn != null && typeof cn !== "string") {
          try {
            Object.defineProperty(el, "className", {
              value: (cn as { baseVal?: string }).baseVal ?? "",
              configurable: true,
              writable: true,
            });
          } catch {
            // Leave it; dom-to-pptx may still handle this node.
          }
        }
      });
    }
    if (phase === "prepare") return { prepared: true };
    preserveLayeredGradientBackgrounds(slides as HTMLElement[]);
    const originalWarn = console.warn;
    let embeddingFailed = false;
    // 上游吞掉字体嵌入异常；导出副本内将该诊断提升为失败，避免静默缺字。
    console.warn = (...args: unknown[]) => {
      if (String(args[0]).startsWith('Failed to embed font family:')) embeddingFailed = true;
      originalWarn.apply(console, args);
    };
    let blob: Blob;
    try {
      blob = await w.domToPptx.exportToPptx(slides, {
        fileName,
        skipDownload: true,
        autoEmbedFonts: true,
        ...(importedFonts.length > 0 ? { fonts: importedFonts } : {}),
        svgAsVector: true,
      });
    } finally {
      console.warn = originalWarn;
    }
    if (embeddingFailed) return { error: 'Declared font embedding failed; use an embeddable font before exporting' };
    if (!(blob instanceof Blob) || blob.size === 0 || blob.size > CLIENT_PPTX_MAX_BYTES) {
      return { error: "dom-to-pptx returned no blob" };
    }
    return { buffer: await blob.arrayBuffer() };
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) };
  }
}

export function prepareDeckStage(hideSelector: string, stageSelector: string): void {
  document.querySelectorAll(hideSelector).forEach((el) => {
    (el as HTMLElement).style.setProperty("display", "none", "important");
  });
  // The repo's <deck-stage> runtime fits its canvas to the viewport with
  // `transform: scale(...)` by default and documents that export must set the
  // `noscale` attribute so the DOM is captured at the authored slide size. Set
  // it here (no-op for plain `.slide` decks that have no <deck-stage>), or a
  // deck whose authored canvas differs from the 1920x1080 capture viewport would
  // be measured + captured at the preview-scaled size instead of 1:1.
  document.querySelectorAll(stageSelector).forEach((el) => {
    el.setAttribute("noscale", "");
    const style = (el as HTMLElement).style;
    style.setProperty("transform", "none", "important");
    style.setProperty("transform-origin", "top left", "important");
  });
  const s = document.createElement("style");
  s.textContent =
    "*,*::before,*::after{animation-duration:0s!important;animation-delay:0s!important;transition-duration:0s!important;transition-delay:0s!important}";
  (document.head || document.documentElement).appendChild(s);
}

export function measureAuthoredSlideBox(el: HTMLElement): { w: number; h: number } | null {
  const stage = el.closest(DECK_STAGE_SELECTOR) as HTMLElement | null;
  const stageSize = stage ? deckStageAuthoredSize(stage) : null;
  if (stageSize) return stageSize;

  const attrSize = sizePair(el.getAttribute("width"), el.getAttribute("height"));
  if (attrSize) return attrSize;

  const styleSize = sizePair(el.style?.width, el.style?.height);
  if (styleSize) return styleSize;

  const view = el.ownerDocument?.defaultView;
  const computed = view?.getComputedStyle?.(el);
  const computedSize = computed ? sizePair(computed.width, computed.height) : null;
  if (computedSize) return computedSize;

  const offsetSize = sizePair(el.offsetWidth, el.offsetHeight);
  if (offsetSize) return offsetSize;

  return null;
}

export function deckStageAuthoredSize(stage: HTMLElement): { w: number; h: number } | null {
  const byProp = sizePair(
    (stage as unknown as { designWidth?: unknown }).designWidth,
    (stage as unknown as { designHeight?: unknown }).designHeight,
  );
  if (byProp) return byProp;
  const byAttr = sizePair(stage.getAttribute("width"), stage.getAttribute("height"));
  if (byAttr) return byAttr;
  const byStyle = sizePair(stage.style?.width, stage.style?.height);
  if (byStyle) return byStyle;
  const view = stage.ownerDocument?.defaultView;
  const computed = view?.getComputedStyle?.(stage);
  const byComputed = computed ? sizePair(computed.width, computed.height) : null;
  if (byComputed) return byComputed;
  return sizePair(stage.offsetWidth, stage.offsetHeight);
}

export function sizePair(w: unknown, h: unknown): { w: number; h: number } | null {
  const width = positiveCssNumber(w);
  const height = positiveCssNumber(h);
  return width != null && height != null ? { w: width, h: height } : null;
}

export function positiveCssNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) && value > 1 ? value : null;
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  const match = /^(\d+(?:\.\d+)?)(?:px)?$/i.exec(trimmed);
  if (!match) return null;
  const n = Number(match[1]);
  return Number.isFinite(n) && n > 1 ? n : null;
}
