import { TFile, TFolder } from 'obsidian';

export class TreeNode {
  key: string;
  title: string;
  level: number;
  parent?: TreeNode;
  children: TreeNode[];

  constructor(key: string, title: string, level: number) {
    this.key = key;
    this.title = title;
    this.level = level;
    this.children = [];
  }
}

export function getHeadingTree(doc: Document = document): TreeNode {
  const headings = doc.querySelectorAll('h1, h2, h3, h4, h5, h6');
  const root = new TreeNode('', 'Root', 0);
  let prev: TreeNode = root;

  headings.forEach((heading) => {
    if ((heading as HTMLElement).style.display === 'none') {
      return;
    }

    const level = parseInt(heading.tagName.slice(1));
    const link = heading.querySelector('a.md-print-anchor');
    const href = (link as HTMLAnchorElement)?.href ?? '';
    const regexMatch = /^af:\/\/(.+)$/.exec(href);

    if (!regexMatch) {
      return;
    }

    const newNode = new TreeNode(regexMatch[1], heading.textContent ?? '', level);

    while (prev.level >= level) {
      prev = prev.parent!;
    }

    prev.children.push(newNode);
    newNode.parent = prev;
    prev = newNode;
  });

  return root;
}

export function modifyDest(doc: Document): Map<string, string> {
  const data = new Map<string, string>();

  doc.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach((heading, i) => {
    const link = document.createElement('a');
    const flag = `${heading.tagName.toLowerCase()}-${i}`;
    link.href = `af://${flag}`;
    link.className = 'md-print-anchor';
    heading.appendChild(link);
    data.set((heading as HTMLElement).dataset.heading ?? '', flag);
  });

  return data;
}

export function convertMapKeysToLowercase(map: Map<string, string>): Map<string, string> {
  return new Map(Array.from(map).map(([key, value]) => [key?.toLowerCase(), value]));
}

export function fixAnchors(doc: Document, dest: Map<string, string>, basename: string): void {
  const lowerDest = convertMapKeysToLowercase(dest);

  doc.querySelectorAll('a.internal-link').forEach((el) => {
    const dataset = (el as HTMLElement).dataset;
    const href = dataset.href ?? '';
    const parts = href.split('#');
    const title = parts[0];
    const anchor = parts[1];

    if (anchor?.startsWith('^')) {
      (el as HTMLAnchorElement).href = href.toLowerCase();
    }

    if (anchor && anchor.length > 0) {
      if (title && title.length > 0 && title !== basename) {
        return;
      }

      const flag = dest.get(anchor) || lowerDest.get(anchor?.toLowerCase());

      if (flag && !anchor.startsWith('^')) {
        (el as HTMLAnchorElement).href = `an://${flag}`;
      }
    }
  });
}

export function px2mm(px: number): number {
  return Math.round(px * 0.26458333333719);
}

export function mm2px(mm: number): number {
  return Math.round(mm * 3.779527559);
}

export function traverseFolder(path: TFolder | TFile): TFile[] {
  if (path instanceof TFile) {
    if (path.extension === 'md') {
      return [path];
    } else {
      return [];
    }
  }

  const arr: TFile[] = [];
  for (const item of path.children) {
    arr.push(...traverseFolder(item as TFolder));
  }
  return arr;
}

export function copyAttributes(node: HTMLElement, attributes: NamedNodeMap): void {
  Array.from(attributes).forEach((attr) => {
    node.setAttribute(attr.name, attr.value);
  });
}

export function isNumber(str: string): boolean {
  return !isNaN(parseFloat(str));
}

export function safeParseInt(str: string, defaultVal: number = 0): number {
  try {
    const num = parseInt(String(str));
    return isNaN(num) ? defaultVal : num;
  } catch (e) {
    return defaultVal;
  }
}

export function safeParseFloat(str: string, defaultVal: number = 0): number {
  try {
    const num = parseFloat(String(str));
    return isNaN(num) ? defaultVal : num;
  } catch (e) {
    return defaultVal;
  }
}
