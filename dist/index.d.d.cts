interface ConvertOptions {
  extractPageStyles?: boolean;
}
interface ConvertResult {
  html: string;
  pageStyles: any | null;
  pageStylesCss: string | null;
}

declare class DocxToHtmlConverter {
  constructor(zip: any, ParserClass?: new (opts?: any) => any);
  static create(arrayBuffer: ArrayBuffer, ParserClass?: new (opts?: any) => any): Promise<DocxToHtmlConverter>;
  convert(options?: ConvertOptions): Promise<ConvertResult>;
}

export { type ConvertOptions, type ConvertResult, DocxToHtmlConverter as default };
