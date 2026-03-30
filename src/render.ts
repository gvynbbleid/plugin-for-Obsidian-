import { App, Component, MarkdownRenderer, TFile, Notice } from 'obsidian';
import { RenderParam, RenderResult } from './type';
import { copyAttributes, fixAnchors, modifyDest } from './utils';

export function getAllStyles(): string[] {
  const cssTexts: string[] = [];

  Array.from(document.styleSheets).forEach((sheet) => {
    const id = (sheet.ownerNode as HTMLElement)?.id;
    if (id?.startsWith('svelte-')) {
      return;
    }

    const href = (sheet.ownerNode as HTMLLinkElement)?.href;
    const division = `/* ----------${id ? `id:${id}` : href ? `href:${href}` : ''}---------- */`;
    cssTexts.push(division);

    try {
      Array.from(sheet?.cssRules ?? []).forEach((rule) => {
        cssTexts.push(rule.cssText);
      });
    } catch (error) {
      console.error(error);
    }
  });

  cssTexts.push(...getPatchStyle());
  return cssTexts;
}

const CSS_PATCH = `
/* ---------- css patch ---------- */

body {
  overflow: auto !important;
}
@media print {
  .print .markdown-preview-view {
    height: auto !important;
  }
  .md-print-anchor, .blockid {
    white-space: pre !important;
    border-left: none !important;
    border-right: none !important;
    border-top: none !important;
    border-bottom: none !important;
    display: inline-block !important;
    position: absolute !important;
    width: 1px !important;
    height: 1px !important;
    right: 0 !important;
    outline: 0 !important;
    background: 0 0 !important;
    text-decoration: initial !important;
    text-shadow: initial !important;
  }
}
@media print {
  table {
    break-inside: auto;
  }
  tr {
    break-inside: avoid;
    break-after: auto;
  }
}

img.__canvas__ {
  width: 100% !important;
  height: 100% !important;
}
`;

function getPatchStyle(): string[] {
  return [CSS_PATCH, ...getPrintStyle()];
}

function getPrintStyle(): string[] {
  const cssTexts: string[] = [];

  Array.from(document.styleSheets).forEach((sheet) => {
    try {
      const cssRules = sheet?.cssRules ?? [];
      Array.from(cssRules).forEach((rule) => {
        if (rule.constructor.name === 'CSSMediaRule') {
          if ((rule as CSSMediaRule).conditionText === 'print') {
            const res = rule.cssText.replace(/@media print\s*\{(.+)\}/gms, '$1');
            cssTexts.push(res);
          }
        }
      });
    } catch (error) {
      console.error(error);
    }
  });

  return cssTexts;
}

function generateDocId(n: number): string {
  return Array.from({ length: n }, () => (16 * Math.random() | 0).toString(16)).join('');
}

function getFrontMatter(app: App, file: TFile): Record<string, any> {
  const cache = app.metadataCache.getFileCache(file);
  return cache?.frontmatter ?? {};
}

export async function renderMarkdown({
  app,
  file,
  config,
  extra
}: RenderParam): Promise<RenderResult> {
  const startTime = new Date().getTime();
  const ws = app.workspace;
  const leaf = ws.getLeaf(true);
  await leaf.openFile(file);
  const view = leaf.view as any;

  const data = view?.data ?? ws.getActiveFileView()?.data ?? ws.activeEditor?.data;

  if (!data) {
    new Notice('data is empty!');
  }

  const frontMatter = getFrontMatter(app, file);
  const cssclasses: string[] = [];

  for (const [key, val] of Object.entries(frontMatter)) {
    if (key.toLowerCase() === 'cssclass' || key.toLowerCase() === 'cssclasses') {
      if (Array.isArray(val)) {
        cssclasses.push(...val);
      } else {
        cssclasses.push(val);
      }
    }
  }

  const comp = new Component();
  comp.load();

  const printEl = document.body.createDiv('print');
  const viewEl = printEl.createDiv({
    cls: `markdown-preview-view markdown-rendered ${cssclasses.join(' ')}`
  });

  await app.vault.cachedRead(file);
  viewEl.toggleClass('rtl', app.vault.getConfig('rightToLeft'));
  viewEl.toggleClass('show-properties', 'hidden' !== app.vault.getConfig('propertiesInDocument'));

  const title = extra?.title ?? frontMatter?.title ?? file.basename;

  viewEl.createEl('h1', { text: title }, (e) => {
    e.addClass('__title__');
    e.style.display = config.showTitle ? 'block' : 'none';
    e.id = extra?.id ?? '';
  });

  const cache = app.metadataCache.getFileCache(file);
  const blocks = new Map(Object.entries(cache?.blocks ?? {}));

  const lines = (data?.split('\n') ?? []).map((line: string, i: number) => {
    for (const blockData of blocks.values()) {
      const { id, position } = blockData as { id: string; position: { start: { line: number }; end: { line: number } } };
      const blockid = `^${id}`;
      if (line.includes(blockid) && i >= position.start.line && i <= position.end.line) {
        blocks.delete(id);
        return line.replace(blockid, `<span id="${blockid}" class="blockid"></span> ${blockid}`);
      }
    }
    return line;
  });

  [...blocks.values()].forEach((blockData) => {
    const { id, position } = blockData as { id: string; position: { start: { line: number }; end: { line: number } } };
    const idx = position.start.line;
    lines[idx] = `<span id="^${id}" class="blockid"></span>\n\n` + lines[idx];
  });

  const fragment: any = {
    children: undefined,
    appendChild(e: any) {
      this.children = e?.children;
      throw new Error('exit');
    }
  };

  const promises: Promise<any>[] = [];

  try {
    await MarkdownRenderer.render(app, lines.join('\n'), fragment, file.path, comp);
  } catch (error) {
    // Expected error from fragment.appendChild
  }

  const el = createFragment();
  Array.from(fragment.children).forEach((item: any) => {
    el.createDiv({}, (t: any) => {
      return t.appendChild(item);
    });
  });

  viewEl.appendChild(el);

  // Post-process manually since postProcess may not be available
  await Promise.all(promises);

  printEl.findAll('a.internal-link').forEach((el: HTMLElement) => {
    const dataset = (el as any).dataset;
    const href = dataset?.href ?? '';
    const parts = href.split('#');
    const title = parts[0];
    const anchor = parts[1];

    if ((!title || title?.length === 0 || title === file.basename) && anchor?.startsWith('^')) {
      return;
    }

    el.removeAttribute('href');
  });

  try {
    await fixWaitRender(data, viewEl);
  } catch (error) {
    console.warn('wait timeout');
  }

  fixCanvasToImage(viewEl);

  const doc = document.implementation.createHTMLDocument('document');
  doc.body.appendChild(printEl.cloneNode(true));
  printEl.detach();
  comp.unload();
  printEl.remove();

  doc.title = title;
  leaf.detach();

  console.log(`md render time:${new Date().getTime() - startTime}ms`);

  return { doc, frontMatter, file };
}

export function fixDoc(doc: Document, title: string): Document {
  const dest = modifyDest(doc);
  fixAnchors(doc, dest, title);
  encodeEmbeds(doc);
  return doc;
}

function encodeEmbeds(doc: Document): void {
  const spans = Array.from(doc.querySelectorAll('span.markdown-embed')).reverse();
  spans.forEach((span) => (span.innerHTML = encodeURIComponent(span.innerHTML)));
}

async function fixWaitRender(data: string, viewEl: HTMLElement): Promise<void> {
  if (data.includes('```dataview') || data.includes('```gEvent') || data.includes('![[')) {
    await sleep(2000);
  }

  try {
    await waitForDomChange(viewEl);
  } catch (error) {
    await sleep(1000);
  }
}

function fixCanvasToImage(el: HTMLElement): void {
  for (const canvas of Array.from(el.querySelectorAll('canvas'))) {
    const data = (canvas as HTMLCanvasElement).toDataURL();
    const img = document.createElement('img');
    img.src = data;
    copyAttributes(img, canvas.attributes);
    img.className = '__canvas__';
    canvas.replaceWith(img);
  }
}

export function createWebview(scale: number = 1.25): any {
  const webview = document.createElement('webview');
  (webview as any).src = `app://obsidian.md/help.html`;
  webview.setAttribute(
    'style',
    `height:calc(${scale} * 100%);
     width: calc(${scale} * 100%);
     transform: scale(${1 / scale}, ${1 / scale});
     transform-origin: top left;
     border: 1px solid #f2f2f2;
    `
  );
  (webview as any).nodeintegration = true;
  return webview;
}

function waitForDomChange(
  target: HTMLElement,
  timeout: number = 2000,
  interval: number = 200
): Promise<boolean> {
  return new Promise((resolve, reject) => {
    let timer: NodeJS.Timeout;
    const observer = new MutationObserver(() => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        observer.disconnect();
        resolve(true);
      }, interval);
    });

    observer.observe(target, {
      childList: true,
      subtree: true,
      attributes: true,
      characterData: true
    });

    setTimeout(() => {
      observer.disconnect();
      reject(new Error(`timeout ${timeout}ms`));
    }, timeout);
  });
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
