import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import {
  ORIGINAL_KEY_COL,
  ORIGINAL_FALLBACK_KEY_COL,
  ORIGINAL_TRACKING_COL, // "운송장번호"
} from "@/lib/constants/excel";
import { makeDatedFileName } from "@/lib/utils/filename";
import { normalizeHeader, toText } from "@/lib/utils/normalize";
import { patchOriginalWorkbook, type TCellPatch, type TSheetPatch } from "@/lib/excel/patchOriginalWorkbook";

type TRow = Record<string, unknown>;

export type TApplyResult = {
  updatedHeaders: string[];
  updatedRows: TRow[];
  unmatched: Array<{ customerOrderNo: string; tracking: string }>;
  duplicates: Array<{ key: string; count: number }>;
};

const ALT_TRACKING_COL = "송장번호";

const getHeader = (originalHeaders: string[], wanted: string) => {
  const headerMap = new Map<string, string>();
  for (const h of originalHeaders) headerMap.set(normalizeHeader(h), h);
  return headerMap.get(normalizeHeader(wanted)) ?? wanted;
};

const hasHeader = (originalHeaders: string[], wanted: string) =>
  originalHeaders.some((h) => normalizeHeader(h) === normalizeHeader(wanted));

// ✅ 원본 키: 상품주문번호 우선, 없으면 ★쇼핑몰 주문번호★
const getOriginalOrderKey = (
  row: TRow,
  keyHeader: string,
  fallbackKeyHeader: string,
) => {
  const primary = String(row[keyHeader] ?? "").trim();
  if (primary) return primary;

  const fallback = String(row[fallbackKeyHeader] ?? "").trim();
  if (fallback) return fallback;

  return "";
};

// 운송장번호(2) / 송장번호(2) 같은 컬럼명 생성
const trackingNthHeader = (baseTrackingHeader: string, n: number) =>
  `${baseTrackingHeader}(${n})`;

export const applyTracking = (
  originalHeaders: string[],
  originalRows: TRow[],
  replyMap: Map<string, string[]>, // 고객주문번호 -> 운송장번호[]
): TApplyResult => {
  const keyHeader = getHeader(originalHeaders, ORIGINAL_KEY_COL);
  const fallbackKeyHeader = getHeader(
    originalHeaders,
    ORIGINAL_FALLBACK_KEY_COL,
  );

  // ✅ tracking 컬럼 선택 규칙
  // - 원본에 "송장번호"가 있으면: "송장번호" 기반으로 넣고 확장도 "송장번호(2..)"
  // - 아니면: 기존대로 "운송장번호" 기반
  const trackingBaseWanted = hasHeader(originalHeaders, ALT_TRACKING_COL)
    ? ALT_TRACKING_COL
    : ORIGINAL_TRACKING_COL;

  const trackingHeader = getHeader(originalHeaders, trackingBaseWanted);

  // duplicate keys in original (실제 사용 키 기준)
  const counts = new Map<string, number>();
  for (const r of originalRows) {
    const k = getOriginalOrderKey(r, keyHeader, fallbackKeyHeader);
    if (!k) continue;
    counts.set(k, (counts.get(k) ?? 0) + 1);
  }
  const duplicates = Array.from(counts.entries())
    .filter(([, c]) => c > 1)
    .map(([key, count]) => ({ key, count }));

  // apply (immutable clone)
  const updated: TRow[] = originalRows.map((r) => ({ ...r }));
  const unmatched: Array<{ customerOrderNo: string; tracking: string }> = [];

  // index: key -> row indexes
  const index = new Map<string, number[]>();
  updated.forEach((r, i) => {
    const k = getOriginalOrderKey(r, keyHeader, fallbackKeyHeader);
    if (!k) return;
    const arr = index.get(k) ?? [];
    arr.push(i);
    index.set(k, arr);
  });

  // replyMap에서 최대 운송장 개수(= 필요한 추가 컬럼 수)
  let maxTrackingCount = 1;
  for (const list of replyMap.values()) {
    if (list.length > maxTrackingCount) maxTrackingCount = list.length;
  }

  // ✅ 최종 헤더 구성
  const updatedHeaders = [...originalHeaders];

  const hasTrackingBase = updatedHeaders.some(
    (h) => normalizeHeader(h) === normalizeHeader(trackingHeader),
  );
  if (!hasTrackingBase) updatedHeaders.push(trackingHeader);

  for (let n = 2; n <= maxTrackingCount; n += 1) {
    const h = trackingNthHeader(trackingHeader, n);
    if (
      !updatedHeaders.some((x) => normalizeHeader(x) === normalizeHeader(h))
    ) {
      updatedHeaders.push(h);
    }
  }

  // ✅ 매핑 적용
  // - 1) 회신에는 있지만 원본에 없는 주문번호 → unmatched (회신 기준)
  // - 2) 원본에는 있지만 회신에 없는 주문번호 → unmatched (원본 기준)
  const originalKeys = new Set(index.keys());

  for (const [customerOrderNo, trackingList] of replyMap.entries()) {
    const idxs = index.get(customerOrderNo);

    if (!idxs || idxs.length === 0) {
      // 회신에 운송장이 여러 개면 모두 unmatched로 기록 (원본에 없는 주문번호)
      for (const t of trackingList) {
        unmatched.push({ customerOrderNo, tracking: t });
      }
      continue;
    }

    // replyMap에 등장한 키는 "원본에도 회신도 있는 키"로 간주
    originalKeys.delete(customerOrderNo);

    for (const idx of idxs) {
      // 1번째는 base 컬럼
      updated[idx][trackingHeader] = trackingList[0] ?? "";

      // 2번째부터는 확장 컬럼
      for (let n = 2; n <= trackingList.length; n += 1) {
        const h = trackingNthHeader(trackingHeader, n);
        updated[idx][h] = trackingList[n - 1] ?? "";
      }

      // 나머지 확장 컬럼 빈값 보장
      for (let n = trackingList.length + 1; n <= maxTrackingCount; n += 1) {
        const h = trackingNthHeader(trackingHeader, n);
        if (!(h in updated[idx])) updated[idx][h] = "";
      }
    }
  }

  // ✅ 원본에는 있지만 회신에 없는 주문번호를 unmatched에 추가
  for (const key of originalKeys) {
    unmatched.push({ customerOrderNo: key, tracking: "" });
  }

  // ✅ 모든 row에 tracking 컬럼 빈값 보장
  for (const r of updated) {
    if (!(trackingHeader in r)) r[trackingHeader] = "";
    for (let n = 2; n <= maxTrackingCount; n += 1) {
      const h = trackingNthHeader(trackingHeader, n);
      if (!(h in r)) r[h] = "";
    }
  }

  return { updatedHeaders, updatedRows: updated, unmatched, duplicates };
};

const isTrackingCol = (header: string) => {
  const n = normalizeHeader(header);
  return (
    n.startsWith(normalizeHeader("운송장번호")) ||
    n.startsWith(normalizeHeader("송장번호"))
  );
};

const countNonEmptyTrackingCells = (headers: string[], rows: TRow[]) => {
  const trackingHeaders = headers.filter(isTrackingCol);
  let n = 0;
  for (const row of rows) {
    for (const h of trackingHeaders) {
      if (String(toText(row[h] ?? "")).trim() !== "") n += 1;
    }
  }
  return n;
};

const stripXlsx = (name: string) => name.replace(/\.xlsx$/i, "").trim();

const sanitizeDownloadBaseName = (name: string) => {
  const cleaned = String(name ?? "")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  const safe = cleaned.length > 200 ? cleaned.slice(0, 200).trim() : cleaned;
  return safe || "한섬누리";
};

const makeFinalOriginalFileName = (
  originalFileName: string,
  headers: string[],
  rows: TRow[],
) => {
  const base = sanitizeDownloadBaseName(stripXlsx(originalFileName));
  const count = countNonEmptyTrackingCells(headers, rows);
  return `${base}_운송장번호_반영완료_${count}건.xlsx`;
};

const fillTrackingInAllWorksheets = (
  workbook: ExcelJS.Workbook,
  headers: string[],
  rows: TRow[],
  replyTrackingByOrder: Record<string, string[]> = {},
) => {
  const primaryHeader = getHeader(headers, ORIGINAL_KEY_COL);
  const fallbackHeader = getHeader(headers, ORIGINAL_FALLBACK_KEY_COL);
  const trackingHeaders = headers.filter(isTrackingCol);
  const trackingByKey = new Map<string, Map<string, string>>();
  for (const row of rows) {
    const key = getOriginalOrderKey(row, primaryHeader, fallbackHeader);
    if (key) trackingByKey.set(key, new Map(trackingHeaders.map((header) => [
      normalizeHeader(header).replace(/^(운송장번호|송장번호)/, ""), toText(row[header]).trim(),
    ])));
  }
  // 집계용 시트에는 없고 다른 시트에만 있는 주문도 회신 번호로 반영한다.
  for (const [key, trackingList] of Object.entries(replyTrackingByOrder)) {
    trackingByKey.set(key, new Map(trackingList.map((value, index) => [
      index === 0 ? "" : `(${index + 1})`, value.trim(),
    ])));
  }

  const patches: TSheetPatch[] = [];
  for (const worksheet of workbook.worksheets) {
    const cells = new Map<string, TCellPatch>();
    const columns = new Map<string, number>();
    worksheet.getRow(1).eachCell((cell, column) => {
      columns.set(normalizeHeader(toText(cell.value)), column);
    });
    const primaryCol = columns.get(normalizeHeader(ORIGINAL_KEY_COL));
    const fallbackCol = columns.get(normalizeHeader(ORIGINAL_FALLBACK_KEY_COL));
    if (!primaryCol && !fallbackCol) continue;

    const trackingBase = columns.has(normalizeHeader(ALT_TRACKING_COL))
      ? ALT_TRACKING_COL
      : columns.has(normalizeHeader(ORIGINAL_TRACKING_COL))
        ? ORIGINAL_TRACKING_COL
        : hasHeader(headers, ALT_TRACKING_COL) ? ALT_TRACKING_COL : ORIGINAL_TRACKING_COL;

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const key = (primaryCol ? toText(row.getCell(primaryCol).value).trim() : "")
        || (fallbackCol ? toText(row.getCell(fallbackCol).value).trim() : "");
      const source = trackingByKey.get(key);
      if (!source) return;

      for (const [suffix, value] of source) {
        if (!value) continue;
        const targetHeader = `${trackingBase}${suffix}`;
        let column = columns.get(targetHeader);
        if (!column) {
          column = worksheet.columnCount + 1;
          const headerCell = worksheet.getRow(1).getCell(column);
          headerCell.value = targetHeader;
          headerCell.font = { bold: true };
          cells.set(headerCell.address, { address: headerCell.address, value: targetHeader, highlight: false });
          worksheet.getColumn(column).width = 22;
          columns.set(targetHeader, column);
        }
        const cell = row.getCell(column);
        cell.value = value;
        cells.set(cell.address, { address: cell.address, value, highlight: true, text: true });
        cell.style = { ...cell.style };
        cell.numFmt = "@";
      }
    });

    for (const [header, column] of columns) {
      if (!isTrackingCol(header)) continue;
      for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
        const cell = worksheet.getRow(rowNumber).getCell(column);
        const hasTracking = rowNumber > 1 && toText(cell.value).trim() !== "";
        cells.set(cell.address, { ...cells.get(cell.address), address: cell.address, highlight: hasTracking });
        cell.style = {
          ...cell.style,
          fill: hasTracking
            ? { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF59D" } }
            : { type: "pattern", pattern: "none" },
        };
      }
    }
    patches.push({ name: worksheet.name, cells: Array.from(cells.values()) });
  }
  return patches;
};

export const downloadOriginalWithTracking = async (
  headers: string[],
  rows: TRow[],
  originalFileName: string,
  originalWorkbookBuffer?: ArrayBuffer,
  replyTrackingByOrder?: Record<string, string[]>,
) => {
  const wb = new ExcelJS.Workbook();
  if (originalWorkbookBuffer) {
    await wb.xlsx.load(originalWorkbookBuffer);
    const patches = fillTrackingInAllWorksheets(wb, headers, rows, replyTrackingByOrder);
    const buffer = await patchOriginalWorkbook(originalWorkbookBuffer, patches);
    saveAs(new Blob([buffer]), makeFinalOriginalFileName(originalFileName, headers, rows));
    return;
  }
  const ws = wb.addWorksheet("Sheet1");

  // ✅ 컬럼 정의
  ws.columns = headers.map((h) => ({
    header: h,
    key: h,
    width: h.length > 10 ? 22 : 16,
  }));

  // ✅ 데이터 입력
  for (const r of rows) {
    const rowData: Record<string, unknown> = {};
    for (const h of headers) rowData[h] = toText(r[h]);
    ws.addRow(rowData);
  }

  // ✅ 헤더 bold
  ws.getRow(1).font = { bold: true };

  // =========================================================
  // ✅ 1) "전체 표"에 검정 테두리 적용
  // =========================================================
  const lastRow = ws.rowCount; // 실제 row 수 (헤더 포함)
  const lastCol = headers.length;

  // ✅ 수정
  const blackBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };

  for (let r = 1; r <= lastRow; r += 1) {
    for (let c = 1; c <= lastCol; c += 1) {
      const cell = ws.getRow(r).getCell(c);
      cell.border = blackBorder;
      // 보기 좋게 세로 가운데 정렬(선택)
      cell.alignment = { vertical: "middle" };
    }
  }

  // =========================================================
  // ✅ 2) 운송장 값이 있는 데이터 셀만 노란 배경 강조
  // =========================================================
  const isTrackingHeader = (header: string) => {
    const normalized = normalizeHeader(header);
    return (
      normalized.startsWith(normalizeHeader("운송장번호")) ||
      normalized.startsWith(normalizeHeader("송장번호"))
    );
  };

  const yellowFill: ExcelJS.Fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFF59D" }, // 연한 노랑
  };

  headers.forEach((header, colIdx) => {
    if (!isTrackingHeader(header)) return;

    const colNumber = colIdx + 1;
    for (let r = 2; r <= lastRow; r += 1) {
      const cell = ws.getRow(r).getCell(colNumber);
      if (toText(cell.value).trim()) cell.fill = yellowFill;
      // 테두리도 다시 한번 명시(혹시 fill 때문에 덮일까봐)
      cell.border = blackBorder;
    }
  });

  // ✅ 필터 기능 제거(요청사항 유지)

  const buf = await wb.xlsx.writeBuffer();
  const outName = makeFinalOriginalFileName(originalFileName, headers, rows);
  saveAs(new Blob([buf]), outName);
};

export const downloadUnmatchedExcel = async (
  unmatched: Array<{ customerOrderNo: string; tracking: string }>,
) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("미매칭");

  ws.columns = [
    { header: "고객주문번호", key: "customerOrderNo", width: 24 },
    { header: "운송장번호", key: "tracking", width: 20 },
  ];

  unmatched.forEach((u) => ws.addRow(u));

  ws.getRow(1).font = { bold: true };

  // ✅ 필터 기능 제거(요청사항)
  // ws.autoFilter = { ... }  // <- 없음

  const buf = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buf]), makeDatedFileName("한섬누리_미매칭_주문목록.xlsx"));
};
