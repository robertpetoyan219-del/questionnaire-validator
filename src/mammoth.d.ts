declare module "mammoth" {
  interface Result {
    value: string;
    messages: unknown[];
  }
  export function extractRawText(options: { arrayBuffer: ArrayBuffer }): Promise<Result>;
}
