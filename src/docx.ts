import HTMLtoDOCX from 'html-to-docx';
import { RenderResult, DocxConfig, TreeNode } from './type';
import { safeParseFloat, safeParseInt } from './utils';
import { PageSize } from './constant';

interface DocxOptions {
  config: DocxConfig;
  frontMatter: Record<string, any>;
  displayMetadata: boolean;
  maxLevel: number;
  headings: TreeNode;
}

export async function generateDocx(renderResult: RenderResult, options: DocxOptions): Promise<Uint8Array> {
  const { doc, frontMatter } = renderResult;
  const { config, displayMetadata } = options;

  const htmlContent = getHtmlContent(doc, config);
  const { pageSize, orientation } = getPageSize(config);
  const margins = getMargins(config);

  const headerContent = config.displayHeader
    ? parseHeaderFooterTemplate(config.headerTemplate, frontMatter)
    : undefined;

  const footerContent = config.displayFooter
    ? parseHeaderFooterTemplate(config.footerTemplate, frontMatter)
    : undefined;

  const header = headerContent
    ? `<div style="font-size:10px;color:#666;text-align:center;">${headerContent}</div>`
    : undefined;

  const footer = footerContent
    ? `<div style="font-size:10px;color:#666;text-align:center;">${footerContent}</div>`
    : undefined;

  const buffer = await HTMLtoDOCX(htmlContent, header, {
    table: { row: { cantSplit: false } },
    footer,
    pageNumber: true,
    pageSize,
    orientation,
    margins,
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
  });

  return new Uint8Array(buffer);
}

function getHtmlContent(doc: Document, config: DocxConfig): string {
  const viewEl = doc.querySelector('.markdown-preview-view');
  if (!viewEl) return '<p></p>';

  const clone = viewEl.cloneNode(true) as HTMLElement;

  const titleEl = clone.querySelector('h1.__title__');
  if (titleEl && !config.showTitle) {
    titleEl.remove();
  }

  const anchors = clone.querySelectorAll('.md-print-anchor, .blockid');
  anchors.forEach((el) => el.remove());

  const internalLinks = clone.querySelectorAll('a.internal-link');
  internalLinks.forEach((el) => {
    el.removeAttribute('href');
  });

  return clone.innerHTML;
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

function getMargins(config: DocxConfig): { top: number; bottom: number; left: number; right: number } {
  if (config.marginType === '0') {
    return { top: 0, bottom: 0, left: 0, right: 0 };
  } else if (config.marginType === '1') {
    const defaultMargin = 1440;
    return { top: defaultMargin, bottom: defaultMargin, left: defaultMargin, right: defaultMargin };
  } else if (config.marginType === '2') {
    const smallMargin = 720;
    return { top: smallMargin, bottom: smallMargin, left: smallMargin, right: smallMargin };
  } else {
    return {
      top: safeParseFloat(config.marginTop, 10) * 1440 / 25.4,
      bottom: safeParseFloat(config.marginBottom, 10) * 1440 / 25.4,
      left: safeParseFloat(config.marginLeft, 10) * 1440 / 25.4,
      right: safeParseFloat(config.marginRight, 10) * 1440 / 25.4
    };
  }
}

function parseHeaderFooterTemplate(
  template: string,
  frontMatter: Record<string, any>
): string {
  let processed = template;
  processed = processed.replace(/\{title\}/g, frontMatter.title ?? '');
  processed = processed.replace(/\{author\}/g, frontMatter.author ?? '');
  processed = processed.replace(/\{date\}/g, new Date().toLocaleDateString());
  return processed;
}
