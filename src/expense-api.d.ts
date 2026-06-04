export interface ExpenseApi {
  pickFile: (filters: { name: string; extensions: string[] }[]) => Promise<string | null>;
  pickSave: (defaultName: string) => Promise<string | null>;
  pickFolder: () => Promise<string | null>;
  loadImagesFromFolder: (folderPath: string) => Promise<string[]>;
  pickImages: () => Promise<string[]>;
  parseStatement: (filePath: string) => Promise<import("./types").TransactionRow[]>;
  loadDefaults: () => Promise<import("./types").Defaults>;
  saveDefaults: (defaults: import("./types").Defaults) => Promise<void>;
  generate: (payload: import("./types").GeneratePayload) => Promise<import("./types").GenerateResult>;
}
