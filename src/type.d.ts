export interface BetterExportDocxSettings {
  prevConfig?: TConfig;
  showTitle: boolean;
  maxLevel: string;
  displayHeader: boolean;
  displayFooter: boolean;
  headerTemplate: string;
  footerTemplate: string;
  displayMetadata: boolean;
  isTimestamp: boolean;
  enabledCss: boolean;
  concurrency: string;
  debug: boolean;
}

export interface TConfig {
  pageSize: string;
  pageWidth?: string;
  pageHeight?: string;
  marginType: string;
  marginTop: string;
  marginBottom: string;
  marginLeft: string;
  marginRight: string;
  showTitle: boolean;
  open: boolean;
  landscape: boolean;
  displayHeader: boolean;
  displayFooter: boolean;
  cssSnippet?: string;
}

export interface RenderParam {
  app: any;
  file: any;
  config: TConfig;
  extra?: {
    title?: string;
    id?: string;
  };
}

export interface RenderResult {
  doc: Document;
  frontMatter: Record<string, any>;
  file: any;
}

export interface TreeNode {
  key: string;
  title: string;
  level: number;
  parent?: TreeNode;
  children: TreeNode[];
}

export interface DocxConfig {
  pageSize: string;
  pageWidth?: string;
  pageHeight?: string;
  marginType: string;
  marginTop: string;
  marginBottom: string;
  marginLeft: string;
  marginRight: string;
  landscape: boolean;
  displayHeader: boolean;
  displayFooter: boolean;
  headerTemplate: string;
  footerTemplate: string;
  showTitle: boolean;
  open: boolean;
}

declare module 'html-to-docx' {
  interface HTMLOptions {
    table?: { row?: { cantSplit?: boolean } };
    footer?: string;
    pageNumber?: boolean;
    pageSize?: { width: number; height: number };
    orientation?: 'portrait' | 'landscape';
    margins?: { top: number; bottom: number; left: number; right: number };
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string;
  }
  function HTMLtoDOCX(html: string, header?: string, options?: HTMLOptions): Promise<Uint8Array>;
  export default HTMLtoDOCX;
}
