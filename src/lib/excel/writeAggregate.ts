import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { ORIGINAL_ITEM_COL, ORIGINAL_BOX_COL } from "@/lib/constants/excel";
import { extractKg } from "@/lib/utils/sort";
import { makeDatedFileName } from "@/lib/utils/filename";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

type TRow = Record<string, unknown>;

export type TAggregateRow = {
  itemName: string;
  totalBox: number;
  kg: number | null;
  fruitKey: string;
};

const FRUIT_KEYWORDS = ["천혜향", "한라봉", "레드향", "감귤", "황금향", "카라향", "청견", "세토카", "데코폰"];

const extractFruitKey = (itemName: string) => {
  for (const k of FRUIT_KEYWORDS) {
    if (itemName.includes(k)) return k;
  }
  return itemName.split(/\s+/)[0] ?? itemName;
};

const normalizeItemNameForAggregate = (itemName: string) => {
  return itemName.replace(/(?<![\d.])5(\s*kg\b)/gi, "4.5$1").replace(/(?<![\d.])10(\s*kg\b)/gi, "9$1");
};

/**
 * row에서 후보 키를 normalize 기반으로 찾아 값 뽑기
 */
const pick = (row: TRow, candidates: string[]) => {
  const keyMap = new Map<string, string>();
  for (const k of Object.keys(row)) keyMap.set(normalizeHeader(k), k);

  for (const c of candidates) {
    const actual = keyMap.get(normalizeHeader(c)) ?? c;
    const v = toText(row[actual]).trim();
    if (v) return v;
  }
  return "";
};

export const buildAggregateRows = (originalRows: TRow[]): TAggregateRow[] => {
  const map = new Map<string, number>();

  for (const row of originalRows) {
    // 품목명: 상품명/품목명 둘 다 대비
    const itemName = pick(row, [ORIGINAL_ITEM_COL, "상품명", "품목명"]).trim();
    if (!itemName) continue;

    // 수량: 박스수량 or 수량 둘 다 대비 (원본 상수 ORIGINAL_BOX_COL도 유지)
    const qtyText = pick(row, [ORIGINAL_BOX_COL, "박스수량", "수량"]);
    const qty = Number(qtyText || 0);

    map.set(itemName, (map.get(itemName) ?? 0) + (Number.isFinite(qty) ? qty : 0));
  }

  const normalizeKgForAggregate = (kg: number | null): number | null => {
    if (kg === null) return null;
    if (kg === 5) return 4.5;
    if (kg === 10) return 9;
    return kg;
  };

  const result: TAggregateRow[] = Array.from(map.entries()).map(([itemName, totalBox]) => {
    const rawKg = extractKg(itemName);

    return {
      itemName,
      totalBox,
      kg: normalizeKgForAggregate(rawKg),
      fruitKey: extractFruitKey(itemName),
    };
  });

  // 품목(과일키) → kg 오름차순 → 품목명
  result.sort((a, b) => {
    const fk = a.fruitKey.localeCompare(b.fruitKey, "ko");
    if (fk !== 0) return fk;

    const ak = a.kg ?? Number.POSITIVE_INFINITY;
    const bk = b.kg ?? Number.POSITIVE_INFINITY;
    if (ak !== bk) return ak - bk;

    return a.itemName.localeCompare(b.itemName, "ko");
  });

  return result;
};

export const downloadAggregateExcel = async (aggregateRows: TAggregateRow[]) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("품목별 집계");

  worksheet.columns = [
    { header: "품목명", key: "itemName", width: 70 },
    { header: "총 박스수량", key: "totalBox", width: 16 },
  ];

  aggregateRows.forEach((row) => {
    worksheet.addRow({
      itemName: normalizeItemNameForAggregate(row.itemName),
      totalBox: row.totalBox,
    });
  });

  worksheet.getRow(1).font = { bold: true };

  // (필터 제거 정책이면 autoFilter는 없애는 게 맞음)
  // worksheet.autoFilter = ...

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), makeDatedFileName("한섬누리_품목별_집계.xlsx"));
};
