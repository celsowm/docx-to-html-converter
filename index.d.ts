export interface ConvertOptions {
  extractPageStyles?: boolean;
}
export interface ConvertResult {
  html: string;
  pageStyles: any | null;
  pageStylesCss: string | null;
}

export default class DocxToHtmlConverter {
  constructor(zip: any, ParserClass?: new (opts?: any) => any);
  static create(arrayBuffer: ArrayBuffer, ParserClass?: new (opts?: any) => any): Promise<DocxToHtmlConverter>;
  convert(options?: ConvertOptions): Promise<ConvertResult>;
}
