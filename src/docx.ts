import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ImageRun,
  ExternalHyperlink,
  PageBreak,
  UnderlineType,
  convertMillimetersToTwip,
  Header,
  Footer,
  PageNumber,
  PageOrientation,
} from 'docx';
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

  const markdown = getMarkdownContent(doc);
  const { pageSize, orientation } = getPageSize(config);
  const margins = getMargins(config);

  const children = parseMarkdown(markdown, config);

  const headerContent = config.displayHeader
    ? parseHeaderFooterTemplate(config.headerTemplate, frontMatter)
    : '';

  const footerContent = config.displayFooter
    ? parseHeaderFooterTemplate(config.footerTemplate, frontMatter)
    : '';

  const docxDoc = new Document({
    creator: displayMetadata
      ? Array.isArray(frontMatter.author)
        ? frontMatter.author.join(', ')
        : (frontMatter.author ?? 'Unknown')
      : 'Obsidian Export',
    title: displayMetadata ? (frontMatter.title ?? doc.title) : doc.title,
    description: displayMetadata ? frontMatter.subject : undefined,
    keywords: displayMetadata
      ? Array.isArray(frontMatter.keywords)
        ? frontMatter.keywords.join(', ')
        : frontMatter.keywords
      : undefined,
    styles: {
      paragraphStyles: [
        {
          id: 'Normal',
          name: 'Normal',
          run: {
            font: 'Calibri',
            size: 22,
          },
        },
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Calibri',
            size: 32,
            bold: true,
            color: '2E74B5',
          },
          paragraph: {
            spacing: { before: 240, after: 120 },
            outlineLevel: 0,
          },
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Calibri',
            size: 26,
            bold: true,
            color: '2E74B5',
          },
          paragraph: {
            spacing: { before: 200, after: 100 },
            outlineLevel: 1,
          },
        },
        {
          id: 'Heading3',
          name: 'Heading 3',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Calibri',
            size: 24,
            bold: true,
            color: '1F4E79',
          },
          paragraph: {
            spacing: { before: 160, after: 80 },
            outlineLevel: 2,
          },
        },
        {
          id: 'Heading4',
          name: 'Heading 4',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Calibri',
            size: 22,
            bold: true,
            color: '1F4E79',
          },
          paragraph: {
            spacing: { before: 120, after: 60 },
            outlineLevel: 3,
          },
        },
        {
          id: 'Heading5',
          name: 'Heading 5',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Calibri',
            size: 22,
            bold: true,
            italics: true,
            color: '2E74B5',
          },
          paragraph: {
            spacing: { before: 120, after: 60 },
            outlineLevel: 4,
          },
        },
        {
          id: 'Heading6',
          name: 'Heading 6',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Calibri',
            size: 22,
            italics: true,
            color: '2E74B5',
          },
          paragraph: {
            spacing: { before: 120, after: 60 },
            outlineLevel: 5,
          },
        },
        {
          id: 'CodeBlock',
          name: 'Code Block',
          basedOn: 'Normal',
          next: 'Normal',
          run: {
            font: 'Consolas',
            size: 18,
            color: '333333',
          },
          paragraph: {
            spacing: { before: 60, after: 60 },
            shading: {
              type: 'clear',
              fill: 'F5F5F5',
            },
          },
        },
        {
          id: 'Quote',
          name: 'Quote',
          basedOn: 'Normal',
          next: 'Normal',
          run: {
            font: 'Calibri',
            size: 22,
            italics: true,
            color: '666666',
          },
          paragraph: {
            indent: { left: convertMillimetersToTwip(10) },
            spacing: { before: 60, after: 60 },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: pageSize.width,
              height: pageSize.height,
              orientation,
            },
            margin: {
              top: margins.top,
              bottom: margins.bottom,
              left: margins.left,
              right: margins.right,
              header: margins.header,
              footer: margins.footer,
            },
          },
        },
        headers: config.displayHeader
          ? {
              default: new Header({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: parseInlineMarkdown(headerContent).map((run) => {
                      if (run instanceof TextRun) {
                        return new TextRun({
                          ...run,
                          size: 18,
                          color: '666666',
                        });
                      }
                      return run;
                    }),
                  }),
                ],
              }),
            }
          : undefined,
        footers: config.displayFooter
          ? {
              default: new Footer({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: footerContent,
                        size: 18,
                        color: '666666',
                      }),
                      new TextRun({ text: '  ' }),
                      new TextRun({
                        children: [PageNumber.CURRENT],
                        size: 18,
                        color: '666666',
                      }),
                    ],
                  }),
                ],
              }),
            }
          : undefined,
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(docxDoc);
  return buffer;
}

function getMarkdownContent(doc: globalThis.Document): string {
  const viewEl = doc.querySelector('.markdown-preview-view');
  if (!viewEl) return '';

  const clone = viewEl.cloneNode(true) as HTMLElement;

  const titleEl = clone.querySelector('h1.__title__');
  if (titleEl) {
    titleEl.remove();
  }

  const anchors = clone.querySelectorAll('.md-print-anchor, .blockid');
  anchors.forEach((el) => el.remove());

  return clone.innerHTML;
}

function parseMarkdown(html: string, config: DocxConfig): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [];
  const lines = html.split('\n');
  let i = 0;

  while (i < lines.length) {
    const line = lines[i].trim();

    if (!line) {
      i++;
      continue;
    }

    if (/^<h([1-6])[^>]*>(.*?)<\/h[1-6]>$/i.test(line)) {
      const match = line.match(/^<h([1-6])[^>]*>(.*?)<\/h[1-6]>$/i);
      if (match) {
        const level = parseInt(match[1]);
        const content = match[2];
        children.push(createHeading(content, level));
      }
      i++;
      continue;
    }

    if (line.startsWith('<p>')) {
      const match = line.match(/^<p[^>]*>(.*?)<\/p>$/i);
      if (match) {
        const content = match[1];
        if (content.includes('page-break')) {
          children.push(new Paragraph({
            children: [new PageBreak()],
          }));
        } else {
          children.push(new Paragraph({
            children: parseInlineMarkdown(content),
            spacing: { after: 120 },
          }));
        }
      }
      i++;
      continue;
    }

    if (line.startsWith('<ul>') || line.startsWith('<ol>')) {
      const { items, endIndex } = extractList(lines, i);
      children.push(...createList(items, line.startsWith('<ol>')));
      i = endIndex + 1;
      continue;
    }

    if (line.startsWith('<table>')) {
      const { table, endIndex } = extractTable(lines, i);
      if (table) {
        children.push(table);
      }
      i = endIndex + 1;
      continue;
    }

    if (line.startsWith('<blockquote>')) {
      const { content, endIndex } = extractBlockquote(lines, i);
      children.push(new Paragraph({
        style: 'Quote',
        children: parseInlineMarkdown(content),
        spacing: { after: 120 },
      }));
      i = endIndex + 1;
      continue;
    }

    if (line.startsWith('<pre>') || line.startsWith('<code>')) {
      const { content, endIndex } = extractCodeBlock(lines, i);
      children.push(new Paragraph({
        style: 'CodeBlock',
        children: [new TextRun({
          text: content,
          font: 'Consolas',
          size: 18,
          color: '333333',
        })],
        spacing: { after: 120 },
      }));
      i = endIndex + 1;
      continue;
    }

    if (line.startsWith('<hr') || line === '<hr>') {
      children.push(new Paragraph({
        children: [new TextRun({
          text: '\u2500'.repeat(50),
          color: 'CCCCCC',
        })],
        spacing: { after: 120 },
      }));
      i++;
      continue;
    }

    if (line.startsWith('<img')) {
      const srcMatch = line.match(/src="([^"]+)"/);
      const altMatch = line.match(/alt="([^"]*)"/);
      if (srcMatch) {
        children.push(createImageParagraph(srcMatch[1], altMatch?.[1] || ''));
      }
      i++;
      continue;
    }

    if (line.startsWith('<div')) {
      const content = line.replace(/<div[^>]*>/, '').replace(/<\/div>$/, '');
      if (content.trim()) {
        children.push(new Paragraph({
          children: parseInlineMarkdown(content),
          spacing: { after: 120 },
        }));
      }
      i++;
      continue;
    }

    if (line.startsWith('<br') || line === '<br>' || line === '<br/>') {
      children.push(new Paragraph({
        children: [],
        spacing: { after: 60 },
      }));
      i++;
      continue;
    }

    if (line.startsWith('<') && !line.startsWith('</')) {
      const tagMatch = line.match(/^<(\w+)/);
      if (tagMatch) {
        const tagName = tagMatch[1];
        const content = line.replace(new RegExp(`^<${tagName}[^>]*>|<\/${tagName}>$`, 'gi'), '');
        if (content.trim()) {
          children.push(new Paragraph({
            children: parseInlineMarkdown(content),
            spacing: { after: 120 },
          }));
        }
      }
      i++;
      continue;
    }

    if (line.startsWith('#')) {
      const match = line.match(/^(#{1,6})\s+(.+)$/);
      if (match) {
        const level = match[1].length;
        const content = match[2];
        children.push(createHeading(content, level));
      }
      i++;
      continue;
    }

    if (line.startsWith('- ') || line.startsWith('* ') || /^\d+\.\s/.test(line)) {
      const { items, endIndex } = extractMarkdownList(lines, i);
      const isOrdered = /^\d+\.\s/.test(line);
      children.push(...createList(items, isOrdered));
      i = endIndex + 1;
      continue;
    }

    if (line.startsWith('> ')) {
      const content = line.substring(2);
      children.push(new Paragraph({
        style: 'Quote',
        children: parseInlineMarkdown(content),
        spacing: { after: 120 },
      }));
      i++;
      continue;
    }

    if (line.startsWith('```')) {
      const { content, endIndex } = extractFencedCode(lines, i);
      children.push(new Paragraph({
        style: 'CodeBlock',
        children: content.split('\n').map((line) =>
          new TextRun({
            text: line,
            font: 'Consolas',
            size: 18,
            color: '333333',
          })
        ),
        spacing: { after: 120 },
      }));
      i = endIndex + 1;
      continue;
    }

    if (line.startsWith('---') || line.startsWith('***') || line.startsWith('___')) {
      children.push(new Paragraph({
        children: [new TextRun({
          text: '\u2500'.repeat(50),
          color: 'CCCCCC',
        })],
        spacing: { after: 120 },
      }));
      i++;
      continue;
    }

    children.push(new Paragraph({
      children: parseInlineMarkdown(line),
      spacing: { after: 120 },
    }));
    i++;
  }

  return children;
}

function extractList(lines: string[], startIndex: number): { items: string[]; endIndex: number } {
  const items: string[] = [];
  let i = startIndex;
  let inList = true;

  while (i < lines.length && inList) {
    const line = lines[i].trim();

    if (line.startsWith('<li>')) {
      const match = line.match(/^<li[^>]*>(.*?)<\/li>$/i);
      if (match) {
        items.push(match[1]);
      }
      i++;
    } else if (line.startsWith('</ul>') || line.startsWith('</ol>')) {
      inList = false;
      i++;
    } else {
      i++;
    }
  }

  return { items, endIndex: i - 1 };
}

function createList(items: string[], isOrdered: boolean): Paragraph[] {
  return items.map((item, index) => {
    const prefix = isOrdered ? `${index + 1}. ` : '\u2022 ';
    return new Paragraph({
      children: [
        new TextRun({
          text: prefix,
          bold: isOrdered,
        }),
        ...parseInlineMarkdown(item),
      ],
      indent: {
        left: convertMillimetersToTwip(10),
        hanging: convertMillimetersToTwip(5),
      },
      spacing: { after: 60 },
    });
  });
}

function extractTable(lines: string[], startIndex: number): { table: Table | null; endIndex: number } {
  let i = startIndex;
  let inTable = true;
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let isHeaderRow = false;

  while (i < lines.length && inTable) {
    const line = lines[i].trim();

    if (line.startsWith('<tr>')) {
      currentRow = [];
      i++;
      continue;
    }

    if (line.startsWith('<th')) {
      const match = line.match(/^<th[^>]*>(.*?)<\/th>$/i);
      if (match) {
        currentRow.push(match[1]);
      }
      isHeaderRow = true;
      i++;
      continue;
    }

    if (line.startsWith('<td')) {
      const match = line.match(/^<td[^>]*>(.*?)<\/td>$/i);
      if (match) {
        currentRow.push(match[1]);
      }
      i++;
      continue;
    }

    if (line.startsWith('</tr>')) {
      if (currentRow.length > 0) {
        rows.push([...currentRow]);
      }
      currentRow = [];
      i++;
      continue;
    }

    if (line.startsWith('</table>')) {
      inTable = false;
      i++;
      continue;
    }

    i++;
  }

  if (rows.length === 0) {
    return { table: null, endIndex: i - 1 };
  }

  const tableRows = rows.map((row, rowIndex) => {
    const cells = row.map((cell) => {
      const isHeader = rowIndex === 0 && isHeaderRow;
      return new TableCell({
        children: [
          new Paragraph({
            children: parseInlineMarkdown(cell),
            alignment: AlignmentType.LEFT,
          }),
        ],
        shading: {
          fill: isHeader ? 'D9E2F3' : 'FFFFFF',
        },
      });
    });

    return new TableRow({
      children: cells,
      tableHeader: rowIndex === 0 && isHeaderRow,
    });
  });

  const table = new Table({
    rows: tableRows,
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
  });

  return { table, endIndex: i - 1 };
}

function extractBlockquote(lines: string[], startIndex: number): { content: string; endIndex: number } {
  let i = startIndex;
  const contentParts: string[] = [];

  while (i < lines.length) {
    const line = lines[i].trim();
    if (line.startsWith('<blockquote>')) {
      const match = line.match(/^<blockquote[^>]*>(.*?)<\/blockquote>$/i);
      if (match) {
        contentParts.push(match[1]);
        i++;
        break;
      }
    }
    if (line.startsWith('</blockquote>')) {
      i++;
      break;
    }
    contentParts.push(line);
    i++;
  }

  return { content: contentParts.join(' '), endIndex: i - 1 };
}

function extractCodeBlock(lines: string[], startIndex: number): { content: string; endIndex: number } {
  let i = startIndex;
  const contentParts: string[] = [];

  while (i < lines.length) {
    const line = lines[i];
    if (line.startsWith('<pre>') || line.startsWith('<code>')) {
      const match = line.match(/^<(?:pre|code)[^>]*>(.*?)<\/(?:pre|code)>$/i);
      if (match) {
        contentParts.push(match[1]);
        i++;
        if (match[0].includes('</pre>') || match[0].includes('</code>')) {
          break;
        }
        continue;
      }
    }
    if (line.startsWith('</pre>') || line.startsWith('</code>')) {
      i++;
      break;
    }
    contentParts.push(line);
    i++;
  }

  return { content: contentParts.join('\n'), endIndex: i - 1 };
}

function extractFencedCode(lines: string[], startIndex: number): { content: string; endIndex: number } {
  let i = startIndex;
  const contentParts: string[] = [];

  i++;

  while (i < lines.length) {
    const line = lines[i];
    if (line.startsWith('```')) {
      i++;
      break;
    }
    contentParts.push(line);
    i++;
  }

  return { content: contentParts.join('\n'), endIndex: i - 1 };
}

function extractMarkdownList(lines: string[], startIndex: number): { items: string[]; endIndex: number } {
  const items: string[] = [];
  let i = startIndex;

  while (i < lines.length) {
    const line = lines[i].trim();

    if (/^[-*]\s/.test(line)) {
      items.push(line.replace(/^[-*]\s/, ''));
      i++;
    } else if (/^\d+\.\s/.test(line)) {
      items.push(line.replace(/^\d+\.\s/, ''));
      i++;
    } else {
      break;
    }
  }

  return { items, endIndex: i - 1 };
}

function createHeading(content: string, level: number): Paragraph {
  const headingMap: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6,
  };

  return new Paragraph({
    heading: headingMap[level] || HeadingLevel.HEADING_1,
    children: parseInlineMarkdown(content),
  });
}

function parseInlineMarkdown(text: string): (TextRun | ExternalHyperlink)[] {
  const runs: (TextRun | ExternalHyperlink)[] = [];
  let remaining = text;

  while (remaining.length > 0) {
    const boldMatch = remaining.match(/\*\*(.+?)\*\*/);
    const italicMatch = remaining.match(/\*(.+?)\*/);
    const codeMatch = remaining.match(/`(.+?)`/);
    const linkMatch = remaining.match(/\[(.+?)\]\((.+?)\)/);
    const strikeMatch = remaining.match(/~~(.+?)~~/);
    const highlightMatch = remaining.match(/==(.+?)==/);

    const matches = [boldMatch, italicMatch, codeMatch, linkMatch, strikeMatch, highlightMatch].filter(Boolean);

    if (matches.length === 0) {
      runs.push(new TextRun({ text: remaining }));
      break;
    }

    let earliestMatch: RegExpMatchArray | null = null;
    let earliestIndex = remaining.length;

    for (const match of matches) {
      if (match && match.index !== undefined && match.index < earliestIndex) {
        earliestMatch = match;
        earliestIndex = match.index;
      }
    }

    if (earliestMatch) {
      if (earliestIndex > 0) {
        runs.push(new TextRun({ text: remaining.substring(0, earliestIndex) }));
      }

      const fullMatch = earliestMatch[0];
      const capture = earliestMatch[1];
      const matchIndex = earliestIndex;

      if (fullMatch.startsWith('**') && fullMatch.endsWith('**')) {
        runs.push(new TextRun({ text: capture, bold: true }));
      } else if (fullMatch.startsWith('*') && fullMatch.endsWith('*')) {
        runs.push(new TextRun({ text: capture, italics: true }));
      } else if (fullMatch.startsWith('`') && fullMatch.endsWith('`')) {
        runs.push(new TextRun({ text: capture, font: 'Consolas', size: 18 }));
      } else if (fullMatch.startsWith('[') && fullMatch.includes('](')) {
        const urlMatch = fullMatch.match(/\[(.+?)\]\((.+?)\)/);
        if (urlMatch) {
          runs.push(new ExternalHyperlink({
            children: [new TextRun({ text: urlMatch[1], style: 'Hyperlink' })],
            link: urlMatch[2],
          }));
        }
      } else if (fullMatch.startsWith('~~') && fullMatch.endsWith('~~')) {
        runs.push(new TextRun({ text: capture, strike: true }));
      } else if (fullMatch.startsWith('==') && fullMatch.endsWith('==')) {
        runs.push(new TextRun({ text: capture, highlight: 'yellow' }));
      }

      remaining = remaining.substring(matchIndex + fullMatch.length);
    } else {
      runs.push(new TextRun({ text: remaining }));
      break;
    }
  }

  return runs;
}

function createImageParagraph(src: string, alt: string): Paragraph {
  try {
    if (src.startsWith('data:image')) {
      const match = src.match(/^data:image\/(\w+);base64,(.+)$/);
      if (match) {
        const buffer = Buffer.from(match[2], 'base64');
        const format = match[1].toUpperCase();
        const imageType = format === 'PNG' ? 'png' : (format === 'JPEG' || format === 'JPG') ? 'jpg' : 'png';
        return new Paragraph({
          children: [
            new ImageRun({
              type: imageType,
              data: buffer,
              transformation: {
                width: 400,
                height: 300,
              },
            }),
          ],
          spacing: { after: 120 },
        });
      }
    }

    return new Paragraph({
      children: [
        new TextRun({
          text: `[Image: ${alt || src}]`,
          italics: true,
          color: '666666',
        }),
      ],
      spacing: { after: 120 },
    });
  } catch (e) {
    return new Paragraph({
      children: [
        new TextRun({
          text: `[Image: ${alt || src}]`,
          italics: true,
          color: '666666',
        }),
      ],
      spacing: { after: 120 },
    });
  }
}

function getPageSize(config: DocxConfig): { pageSize: { width: number; height: number }; orientation: (typeof PageOrientation)[keyof typeof PageOrientation] } {
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
    pageSize: {
      width: convertMillimetersToTwip(width),
      height: convertMillimetersToTwip(height),
    },
    orientation: config.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
  };
}

function getMargins(config: DocxConfig): { top: number; bottom: number; left: number; right: number; header: number; footer: number } {
  const mmToTwip = (mm: number) => convertMillimetersToTwip(mm);

  if (config.marginType === '0') {
    return { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 };
  } else if (config.marginType === '1') {
    const defaultMargin = mmToTwip(25.4);
    return { top: defaultMargin, bottom: defaultMargin, left: defaultMargin, right: defaultMargin, header: mmToTwip(12.7), footer: mmToTwip(12.7) };
  } else if (config.marginType === '2') {
    const smallMargin = mmToTwip(12.7);
    return { top: smallMargin, bottom: smallMargin, left: smallMargin, right: smallMargin, header: mmToTwip(6.35), footer: mmToTwip(6.35) };
  } else {
    return {
      top: mmToTwip(safeParseFloat(config.marginTop, 25.4)),
      bottom: mmToTwip(safeParseFloat(config.marginBottom, 25.4)),
      left: mmToTwip(safeParseFloat(config.marginLeft, 25.4)),
      right: mmToTwip(safeParseFloat(config.marginRight, 25.4)),
      header: mmToTwip(12.7),
      footer: mmToTwip(12.7),
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
