import * as XLSX from "xlsx";
import fs from "fs";

const CARD_HEADERS = [
  "이용일자",
  "카드번호",
  "이용자명",
  "사원번호",
  "부서명",
  "국내이용금액",
  "국내청구금액",
  "해외현지금액",
  "통화코드",
  "가맹점",
  "가맹점사업자번호",
  "승인번호",
  "할부개월수",
  "청구회차",
];

export interface ParsedTransaction {
  cardColumns: (string | number)[];
  domesticClaimAmount: number;
  personalUseAmount: number;
  merchant: string;
  useDate: string;
}

function sheetToRows(wb: XLSX.WorkBook): unknown[][] {
  const name =
    wb.SheetNames.find((n) => n.includes("청구") || n.includes("내역")) ??
    wb.SheetNames[0];
  const sheet = wb.Sheets[name];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false }) as unknown[][];
}

function normalizeHeader(cell: unknown): string {
  return String(cell ?? "")
    .replace(/^\uFEFF/, "")
    .trim();
}

function findHeaderRow(rows: unknown[][]): number {
  for (let r = 0; r < Math.min(rows.length, 30); r++) {
    const row = rows[r];
    if (!Array.isArray(row)) continue;
    if (normalizeHeader(row[0]) === "이용일자") return r;
    if (row.some((c) => normalizeHeader(c) === "이용일자")) return r;
  }
  return -1;
}

function excelSerialToDateString(serial: number): string {
  const dc = (XLSX.SSF as { parse_date_code?: (n: number) => { y: number; m: number; d: number } }).parse_date_code;
  if (dc) {
    const d = dc(serial);
    if (d) {
      return `${d.y}.${String(d.m).padStart(2, "0")}.${String(d.d).padStart(2, "0")}`;
    }
  }
  const utc = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return `${utc.getUTCFullYear()}.${String(utc.getUTCMonth() + 1).padStart(2, "0")}.${String(utc.getUTCDate()).padStart(2, "0")}`;
}

function formatCell(v: unknown): string | number {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    if (v > 40000 && v < 60000) return excelSerialToDateString(v);
    return v;
  }
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const d = String(v.getDate()).padStart(2, "0");
    return `${y}.${m}.${d}`;
  }
  return String(v).trim();
}

function parseAmount(v: unknown): number {
  if (typeof v === "number") return v;
  const s = String(v ?? "").replace(/,/g, "").trim();
  return parseFloat(s) || 0;
}

export function parseCardStatement(filePath: string): ParsedTransaction[] {
  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
  const rows = sheetToRows(wb);
  const headerRow = findHeaderRow(rows);
  if (headerRow < 0) {
    throw new Error(
      "명세서에서 '이용일자' 헤더를 찾지 못했습니다. 카드사 명세서(청구내역상세) 파일인지 확인하세요."
    );
  }

  const result: ParsedTransaction[] = [];
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!Array.isArray(row)) continue;
    const useDate = formatCell(row[0]);
    if (useDate === "" || useDate === 0) continue;

    const cardColumns: (string | number)[] = [];
    for (let c = 0; c < 14; c++) {
      cardColumns.push(formatCell(row[c]));
    }

    const domesticClaimAmount = parseAmount(row[6] ?? row[5] ?? 0);
    if (!domesticClaimAmount && !String(row[9] ?? "").trim()) continue;

    result.push({
      cardColumns,
      domesticClaimAmount,
      personalUseAmount: 0,
      merchant: String(row[9] ?? ""),
      useDate: String(useDate),
    });
  }

  if (result.length === 0) {
    throw new Error(
      "명세서에 거래 행이 없습니다. 엑셀에서 '보호된 보기'를 끄고 저장한 뒤 다시 시도하거나, 카드 명세서 원본(.xls)을 선택하세요."
    );
  }

  return result;
}

export { CARD_HEADERS };