import HTMLtoDOCX from 'html-to-docx';
import { RenderResult, DocxConfig } from './type';
import { safeParseFloat } from './utils';
import { PageSize } from './constant';

interface DocxOptions {
  config: DocxConfig;
  frontMatter: Record<string, any>;
  displayMetadata: boolean;
  maxLevel: number;
  headings: any;
}

export async function generateDocx(renderResult: RenderResult, options: DocxOptions): Promise<Buffer> {
  const { doc, frontMatter, file } = renderResult;
  const { config, displayMetadata } = options;

  console.log('[Export DOCX] Starting export for file:', file?.name ?? 'unknown');
  console.log('[Export DOCX] Document title:', doc?.title);
  console.log('[Export DOCX] Document has body:', !!doc?.body);

  const htmlContent = getHtmlContent(doc, config);

  console.log('[Export DOCX] HTML content length:', htmlContent.length);
  console.log('[Export DOCX] HTML preview:', htmlContent.substring(0, 300));

  if (!htmlContent || htmlContent.trim().length === 0) {
    throw new Error('Empty HTML content - nothing to export');
  }

  const { pageSize, orientation } = getPageSize(config);
  const margins = getMargins(config);

  const headerEnabled = config.displayHeader;
  const footerEnabled = config.displayFooter;

  const headerContent = headerEnabled
    ? parseHeaderFooterTemplate(config.headerTemplate, frontMatter)
    : '<p></p>';

  const footerContent = footerEnabled
    ? parseHeaderFooterTemplate(config.footerTemplate, frontMatter)
    : '<p></p>';

  const header = `<div style="font-size:10px;color:#666;text-align:center;">${headerContent}</div>`;
  const footer = `<div style="font-size:10px;color:#666;text-align:center;">${footerContent}</div>`;

  console.log('[Export DOCX] Calling HTMLtoDOCX...');

  try {
    const buffer = await HTMLtoDOCX(htmlContent, header, {
      header: headerEnabled,
      footer: footerEnabled,
      table: { row: { cantSplit: false } },
      pageNumber: footerEnabled,
      pageSize,
      orientation,
      margins,
      font: 'Calibri',
      fontSize: 22,
      title: displayMetadata ? frontMatter.title ?? doc.title : undefined,
      subject: displayMetadata ? frontMatter.subject : undefined,
      creator: displayMetadata
        ? Array.isArray(frontMatter.author)
          ? frontMatter.author.join(', ')
          : frontMatter.author
        : undefined,
      keywords: displayMetadata
        ? Array.isArray(frontMatter.keywords)
          ? frontMatter.keywords.join(', ')
          : frontMatter.keywords
        : undefined
    }, footer);

    console.log('[Export DOCX] Buffer size:', buffer?.length ?? 0);
    console.log('[Export DOCX] Buffer type:', buffer?.constructor?.name);

    if (!buffer || buffer.length === 0) {
      throw new Error('Generated buffer is empty');
    }

    return buffer;
  } catch (error) {
    console.error('[Export DOCX] Error during conversion:', error);
    throw error;
  }
}

function getHtmlContent(doc: globalThis.Document, config: DocxConfig): string {
  console.log('[Export DOCX] getHtmlContent - document body children:', doc?.body?.childNodes?.length);
  console.log('[Export DOCX] getHtmlContent - document title:', doc?.title);

  const viewEl = doc.querySelector('.markdown-preview-view');
  console.log('[Export DOCX] getHtmlContent - found .markdown-preview-view:', !!viewEl);

  if (!viewEl) {
    console.log('[Export DOCX] getHtmlContent - no .markdown-preview-view, returning fallback');
    return '<p>No content found</p>';
  }

  const clone = viewEl.cloneNode(true) as HTMLElement;
  console.log('[Export DOCX] getHtmlContent - cloned HTML length:', clone.innerHTML.length);
  console.log('[Export DOCX] getHtmlContent - clone HTML preview:', clone.innerHTML.substring(0, 300));

  const titleEl = clone.querySelector('h1.__title__');
  if (titleEl && !config.showTitle) {
    titleEl.remove();
  }

  removeUnsupportedElements(clone);
  cleanAttributes(clone);
  fixInternalLinks(clone);
  fixImages(clone);
  fixTables(clone);
  removeEmptyElements(clone);

  const result = clone.innerHTML;
  console.log('[Export DOCX] getHtmlContent - final HTML length:', result.length);
  console.log('[Export DOCX] getHtmlContent - final HTML preview:', result.substring(0, 300));

  if (!result || result.trim().length === 0) {
    console.log('[Export DOCX] getHtmlContent - HTML is empty after cleaning, using fallback');
    return '<p>Content was empty after processing</p>';
  }

  return result;
}

function removeUnsupportedElements(el: HTMLElement): void {
  const unsupportedSelectors = [
    '.md-print-anchor',
    '.blockid',
    'script',
    'style',
    'svg',
    'canvas',
    'video',
    'audio',
    'iframe',
    'object',
    'embed',
    'details',
    'summary',
    'dialog',
    'template',
    'math',
    'mjx-container',
    '[class*="math"]',
    '[class*="katex"]',
    '[class*="mermaid"]',
    '.callout',
    '.callout-content',
    '.callout-icon',
    '.callout-title',
    '.markdown-embed',
    '.file-embed',
    '.internal-embed',
    '.footnotes',
    '.footnote-ref',
    '.footnote-backref',
    'a.tag',
    'span.cm-hashtag',
    '.metadata-container',
    '.metadata-properties',
    '.property',
    '[data-line]',
    '.cm-line',
    '.source',
    '.code-block-flair',
    '.edit-block-button',
    '.collapse-indicator',
    '.heading-collapse-indicator',
    '.collapse-icon',
    '.list-bullet',
    '.list-marker',
    '.task-list-item-checkbox',
    'input[type="checkbox"]',
    '.search-result',
    '.highlight',
    'mark',
    '.obsidian-search-match-highlight',
  ];

  unsupportedSelectors.forEach((selector) => {
    try {
      const elements = el.querySelectorAll(selector);
      elements.forEach((e: Element) => e.remove());
    } catch (e) {
    }
  });
}

function cleanAttributes(el: HTMLElement): void {
  const allElements = el.querySelectorAll('*');
  allElements.forEach((node) => {
    const elem = node as HTMLElement;
    const attrs = Array.from(elem.attributes);
    attrs.forEach((attr) => {
      const name = attr.name.toLowerCase();
      if (
        name !== 'href' &&
        name !== 'src' &&
        name !== 'alt' &&
        name !== 'title' &&
        name !== 'colspan' &&
        name !== 'rowspan' &&
        name !== 'width' &&
        name !== 'height' &&
        name !== 'align' &&
        name !== 'valign' &&
        name !== 'style' &&
        name !== 'class' &&
        name !== 'id' &&
        name !== 'type' &&
        name !== 'start' &&
        name !== 'reversed' &&
        name !== 'lang' &&
        name !== 'dir' &&
        name !== 'abbr' &&
        name !== 'axis' &&
        name !== 'headers' &&
        name !== 'scope' &&
        name !== 'span' &&
        name !== 'cite' &&
        name !== 'datetime' &&
        name !== 'hreflang' &&
        name !== 'media' &&
        name !== 'rel' &&
        name !== 'target'
      ) {
        elem.removeAttribute(attr.name);
      }
    });

    if (elem.style) {
      const cssText = elem.getAttribute('style') || '';
      const allowedProperties = [
        'color',
        'background-color',
        'background',
        'font-size',
        'font-weight',
        'font-style',
        'font-family',
        'text-align',
        'text-decoration',
        'text-transform',
        'text-indent',
        'line-height',
        'margin',
        'margin-top',
        'margin-bottom',
        'margin-left',
        'margin-right',
        'padding',
        'padding-top',
        'padding-bottom',
        'padding-left',
        'padding-right',
        'border',
        'border-top',
        'border-bottom',
        'border-left',
        'border-right',
        'border-color',
        'border-style',
        'border-width',
        'width',
        'height',
        'max-width',
        'min-width',
        'display',
        'position',
        'top',
        'left',
        'right',
        'bottom',
        'float',
        'clear',
        'vertical-align',
        'white-space',
        'word-wrap',
        'word-break',
        'overflow',
        'list-style-type',
        'list-style-position',
        'list-style-image',
        'opacity',
        'visibility',
        'z-index',
        'box-sizing',
        'outline',
        'cursor',
        'pointer-events',
        'user-select',
        'resize',
        'zoom',
        'page-break-before',
        'page-break-after',
        'page-break-inside',
        'break-before',
        'break-after',
        'break-inside',
      ];

      const filteredStyles = cssText
        .split(';')
        .filter((style) => {
          const prop = style.split(':')[0]?.trim().toLowerCase();
          return prop && allowedProperties.includes(prop);
        })
        .join(';');

      if (filteredStyles) {
        elem.setAttribute('style', filteredStyles);
      } else {
        elem.removeAttribute('style');
      }
    }
  });
}

function fixInternalLinks(el: HTMLElement): void {
  const links = el.querySelectorAll('a.internal-link');
  links.forEach((link) => {
    const a = link as HTMLAnchorElement;
    a.removeAttribute('href');
    a.removeAttribute('class');
  });
}

function fixImages(el: HTMLElement): void {
  const images = el.querySelectorAll('img');
  images.forEach((img) => {
    const image = img as HTMLImageElement;
    const src = image.getAttribute('src');
    if (!src || src.trim().length === 0) {
      image.remove();
      return;
    }

    if (src.startsWith('app://') || src.startsWith('file://') || src.startsWith('obsidian://')) {
      const alt = image.getAttribute('alt') || 'Image';
      const p = el.ownerDocument.createElement('p');
      p.textContent = `[Image: ${alt}]`;
      p.setAttribute('style', 'font-style: italic; color: #666;');
      image.replaceWith(p);
      return;
    }

    image.removeAttribute('loading');
    image.removeAttribute('draggable');
    image.removeAttribute('class');

    if (!image.hasAttribute('alt')) {
      image.setAttribute('alt', '');
    }
  });
}

function fixTables(el: HTMLElement): void {
  const tables = el.querySelectorAll('table');
  tables.forEach((table) => {
    const t = table as HTMLTableElement;
    t.removeAttribute('class');

    const ths = t.querySelectorAll('th');
    ths.forEach((th) => {
      th.removeAttribute('class');
      th.removeAttribute('style');
    });

    const tds = t.querySelectorAll('td');
    tds.forEach((td) => {
      td.removeAttribute('class');
      td.removeAttribute('style');
    });

    const trs = t.querySelectorAll('tr');
    trs.forEach((tr) => {
      tr.removeAttribute('class');
      tr.removeAttribute('style');
    });
  });
}

function removeEmptyElements(el: HTMLElement): void {
  let changed = true;
  while (changed) {
    changed = false;
    const allElements = el.querySelectorAll('*');
    allElements.forEach((node) => {
      const elem = node as HTMLElement;
      const tagName = elem.tagName.toLowerCase();
      const voidElements = ['br', 'hr', 'img', 'input', 'col', 'area', 'base', 'embed', 'link', 'meta', 'param', 'source', 'track', 'wbr'];

      if (voidElements.includes(tagName)) {
        return;
      }

      if (elem.childNodes.length === 0 && !voidElements.includes(tagName)) {
        if (tagName === 'p' || tagName === 'div' || tagName === 'span' || tagName === 'li' || tagName === 'td' || tagName === 'th') {
          elem.remove();
          changed = true;
        }
      }
    });
  }
}

function getPageSize(config: DocxConfig): { pageSize: { width: number; height: number }; orientation: 'portrait' | 'landscape' } {
  let width: number;
  let height: number;

  if (config.pageSize === 'Custom' && config.pageWidth && config.pageHeight) {
    width = safeParseFloat(config.pageWidth, 210);
    height = safeParseFloat(config.pageHeight, 297);
  } else {
    const size = PageSize[config.pageSize] ?? PageSize['A4'];
    width = size[0];
    height = size[1];
  }

  if (config.landscape) {
    [width, height] = [height, width];
  }

  return {
    pageSize: { width: width * 1440 / 25.4, height: height * 1440 / 25.4 },
    orientation: config.landscape ? 'landscape' : 'portrait'
  };
}

function getMargins(config: DocxConfig): { top: number; bottom: number; left: number; right: number; header: number; footer: number; gutter: number } {
  if (config.marginType === '0') {
    return { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0, gutter: 0 };
  } else if (config.marginType === '1') {
    const defaultMargin = 1440;
    return { top: defaultMargin, bottom: defaultMargin, left: defaultMargin, right: defaultMargin, header: 720, footer: 720, gutter: 0 };
  } else if (config.marginType === '2') {
    const smallMargin = 720;
    return { top: smallMargin, bottom: smallMargin, left: smallMargin, right: smallMargin, header: 360, footer: 360, gutter: 0 };
  } else {
    return {
      top: safeParseFloat(config.marginTop, 10) * 1440 / 25.4,
      bottom: safeParseFloat(config.marginBottom, 10) * 1440 / 25.4,
      left: safeParseFloat(config.marginLeft, 10) * 1440 / 25.4,
      right: safeParseFloat(config.marginRight, 10) * 1440 / 25.4,
      header: 720,
      footer: 720,
      gutter: 0
    };
  }
}

function parseHeaderFooterTemplate(
  template: string,
  frontMatter: Record<string, any>
): string {
  let processed = template;
  processed = processed.replace(/\{title\}/g, frontMatter.title ?? '');
  processed = processed.replace(/\{author\}/g, Array.isArray(frontMatter.author) ? frontMatter.author.join(', ') : (frontMatter.author ?? ''));
  processed = processed.replace(/\{date\}/g, new Date().toLocaleDateString());
  return processed;
}
