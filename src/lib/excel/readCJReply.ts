import ExcelJS from "exceljs";
import { CJ_REPLY_KEY_COL, CJ_REPLY_TRACKING_COL } from "@/lib/constants/excel";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

export type TReplyRecord = {
  key: string; // 고객주문번호
  tracking: string; // 운송장번호
  sourceFileName: string;
};

export type TReadReplyResult = {
  records: TReplyRecord[];
  skipped: Array<{ reason: string; sourceFileName: string }>;
};

const buildHeaderIndexMap = (ws: ExcelJS.Worksheet) => {
  const headerMap = new Map<string, number>(); // normalized header -> colIndex(1-based)

  ws.getRow(1).eachCell((cell, colNumber) => {
    const h = String(cell.value ?? "").trim();
    if (!h) return;
    headerMap.set(normalizeHeader(h), colNumber);
  });

  return headerMap;
};

export const readCjReplyFile = async (
  file: File,
): Promise<TReadReplyResult> => {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const ws = wb.worksheets[0];
  if (!ws) {
    throw new Error(`회신 파일 시트를 찾을 수 없습니다: ${file.name}`);
  }

  const headerMap = buildHeaderIndexMap(ws);
  const keyCol = headerMap.get(normalizeHeader(CJ_REPLY_KEY_COL));
  const trackingCol = headerMap.get(normalizeHeader(CJ_REPLY_TRACKING_COL));

  if (!keyCol || !trackingCol) {
    return {
      records: [],
      skipped: [
        {
          reason: `필수 컬럼을 찾지 못함: ${CJ_REPLY_KEY_COL} / ${CJ_REPLY_TRACKING_COL}`,
          sourceFileName: file.name,
        },
      ],
    };
  }

  const records: TReplyRecord[] = [];
  const skipped: Array<{ reason: string; sourceFileName: string }> = [];

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const key = toText(row.getCell(keyCol).value).trim();
    const tracking = toText(row.getCell(trackingCol).value).trim();

    if (!key) {
      skipped.push({
        reason: "고객주문번호 비어있음",
        sourceFileName: file.name,
      });
      return;
    }
    if (!tracking) {
      skipped.push({
        reason: `운송장번호 비어있음 (고객주문번호=${key})`,
        sourceFileName: file.name,
      });
      return;
    }

    records.push({ key, tracking, sourceFileName: file.name });
  });

  return { records, skipped };
};

export type TReadReplyFilesResult = {
  map: Map<string, string[]>; // 고객주문번호 -> 운송장번호들
  orderFileMap: Map<string, Set<string>>; // 고객주문번호 -> 등장한 파일명 set
  skipped: Array<{ reason: string; sourceFileName: string }>;
};

export const readCjReplyFiles = async (
  files: File[],
): Promise<TReadReplyFilesResult> => {
  const map = new Map<string, string[]>();
  const orderFileMap = new Map<string, Set<string>>();
  const skipped: Array<{ reason: string; sourceFileName: string }> = [];

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const ws = wb.worksheets[0];
    if (!ws) {
      skipped.push({
        reason: "시트를 찾을 수 없음",
        sourceFileName: file.name,
      });
      continue;
    }

    const headerMap = buildHeaderIndexMap(ws);
    const keyCol = headerMap.get(normalizeHeader(CJ_REPLY_KEY_COL));
    const trackingCol = headerMap.get(normalizeHeader(CJ_REPLY_TRACKING_COL));

    if (!keyCol || !trackingCol) {
      skipped.push({
        reason: `필수 컬럼을 찾지 못함: ${CJ_REPLY_KEY_COL} / ${CJ_REPLY_TRACKING_COL}`,
        sourceFileName: file.name,
      });
      continue;
    }

    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const orderNo = toText(row.getCell(keyCol).value).trim();
      const tracking = toText(row.getCell(trackingCol).value).trim();

      if (!orderNo) return;
      if (!tracking) return;

      // ✅ 운송장 누적 (같은 파일 내 중복 주문번호는 정상)
      const list = map.get(orderNo) ?? [];
      if (!list.includes(tracking)) list.push(tracking);
      map.set(orderNo, list);

      // ✅ 파일별 주문번호 추적 (서로 다른 파일 중복 검사용)
      const fileSet = orderFileMap.get(orderNo) ?? new Set<string>();
      fileSet.add(file.name);
      orderFileMap.set(orderNo, fileSet);
    });
  }

  return { map, orderFileMap, skipped };
};
