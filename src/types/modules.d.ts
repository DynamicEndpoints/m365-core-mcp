/**
 * Type declarations for modules without @types packages
 */

declare module 'csv-writer' {
  interface ObjectCsvStringifierParams {
    header: Array<{ id: string; title: string }>;
  }
  
  interface CsvStringifier {
    getHeaderString(): string;
    stringifyRecords(records: any[]): string;
  }
  
  interface ObjectCsvWriterParams {
    path: string;
    header: Array<{ id: string; title: string }>;
    append?: boolean;
  }
  
  interface CsvWriter {
    writeRecords(records: any[]): Promise<void>;
  }
  
  export function createObjectCsvStringifier(params: ObjectCsvStringifierParams): CsvStringifier;
  export function createObjectCsvWriter(params: ObjectCsvWriterParams): CsvWriter;
}

declare module 'xlsx' {
  interface WorkBook {
    SheetNames: string[];
    Sheets: { [sheet: string]: WorkSheet };
  }
  
  interface WorkSheet {
    [cell: string]: any;
  }
  
  interface WritingOptions {
    type?: 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';
    bookType?: 'xlsx' | 'xlsm' | 'xlsb' | 'xls' | 'csv' | 'txt' | 'html' | 'ods';
  }
  
  export const utils: {
    book_new(): WorkBook;
    book_append_sheet(workbook: WorkBook, worksheet: WorkSheet, name?: string): void;
    json_to_sheet(data: any[], opts?: any): WorkSheet;
    aoa_to_sheet(data: any[][], opts?: any): WorkSheet;
  };
  
  export function write(workbook: WorkBook, options: WritingOptions): any;
  export function writeFile(workbook: WorkBook, filename: string, options?: WritingOptions): void;
}

declare module 'handlebars' {
  export function compile(template: string): (context: any) => string;
  export function registerHelper(name: string, fn: (...args: any[]) => any): void;
  export function registerPartial(name: string, partial: string): void;
}
