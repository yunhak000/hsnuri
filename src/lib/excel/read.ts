import ExcelJS from "exceljs";
import {
  ORIGINAL_FALLBACK_KEY_COL,
  ORIGINAL_ITEM_COL,
  ORIGINAL_KEY_COL,
} from "@/lib/constants/excel";
import { normalizeHeader } from "@/lib/utils/normalize";

export type TExcelRow = Record<string, unknown>;

export type TReadExcelResult = {
  headers: string[];
  rows: TExcelRow[];
};

const readHeaders = (worksheet: ExcelJS.Worksheet) => {
  const headers: string[] = [];

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    headers[colNumber - 1] = String(cell.value ?? "").trim();
  });

  return headers;
};

const hasAnyHeader = (headers: string[], candidates: string[]) => {
  const set = new Set(headers.map((h) => normalizeHeader(h)));
  return candidates.some((c) => set.has(normalizeHeader(c)));
};

const pickWorksheet = (workbook: ExcelJS.Workbook) => {
  const worksheets = workbook.worksheets;
  if (worksheets.length === 0) return null;

  for (const ws of worksheets) {
    const headers = readHeaders(ws);
    const hasItem = hasAnyHeader(headers, [ORIGINAL_ITEM_COL, "상품명", "품목명"]);
    const hasOrderKey = hasAnyHeader(headers, [
      ORIGINAL_KEY_COL,
      ORIGINAL_FALLBACK_KEY_COL,
      "상품주문번호",
      "★쇼핑몰 주문번호★",
    ]);

    if (hasItem && hasOrderKey) return ws;
  }

  return worksheets[0] ?? null;
};

export const readExcelFile = async (file: File): Promise<TReadExcelResult> => {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = pickWorksheet(workbook);
  if (!worksheet) {
    throw new Error("엑셀 시트를 찾을 수 없습니다.");
  }

  const headers = readHeaders(worksheet);

  const rows: TExcelRow[] = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const rowData: TExcelRow = {};

    headers.forEach((header, index) => {
      const cellValue = row.getCell(index + 1).value;
      rowData[header] = cellValue ?? "";
    });

    rows.push(rowData);
  });

  return { headers, rows };
};
