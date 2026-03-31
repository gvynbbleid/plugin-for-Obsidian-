declare module 'html-to-docx' {
  interface HTMLOptions {
    header?: boolean;
    headerType?: 'default' | 'first' | 'even';
    footer?: boolean;
    footerType?: 'default' | 'first' | 'even';
    table?: { row?: { cantSplit?: boolean } };
    pageNumber?: boolean;
    skipFirstHeaderFooter?: boolean;
    pageSize?: { width: number; height: number };
    orientation?: 'portrait' | 'landscape';
    margins?: { top: number; bottom: number; left: number; right: number; header?: number; footer?: number; gutter?: number };
    font?: string;
    fontSize?: number;
    complexScriptFontSize?: number;
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string | string[];
    description?: string;
    lastModifiedBy?: string;
    revision?: number;
    createdAt?: Date;
    modifiedAt?: Date;
    decodeUnicode?: boolean;
    lang?: string;
    lineNumber?: boolean;
    lineNumberOptions?: { start?: number; countBy?: number; restart?: 'continuous' | 'newPage' | 'newSection' };
    numbering?: { defaultOrderedListStyleType?: string };
  }
  function HTMLtoDOCX(html: string, header?: string, options?: HTMLOptions, footer?: string): Promise<Buffer>;
  export default HTMLtoDOCX;
}
