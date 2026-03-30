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
