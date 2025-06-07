declare module 'html-pdf-node' {
  export function generatePdf(file: { content: string }, options: any): Promise<Buffer>;
  export default { generatePdf };
}
