export interface Defaults {
  account: string;
  client: string;
  userName: string;
  detail: string;
}

export interface TransactionRow {
  cardColumns: (string | number)[];
  domesticClaimAmount: number;
  personalUseAmount: number;
  userName?: string;
  detail?: string;
  merchant?: string;
  useDate?: string;
}

export interface GeneratePayload {
  statementPath: string;
  templatePath: string;
  outputPath: string;
  transactions: TransactionRow[];
  defaults: Defaults;
  imagePaths: string[];
}

export interface GenerateResult {
  success: boolean;
  outputPath?: string;
  error?: string;
}
