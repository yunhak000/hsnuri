import ExcelJS from "exceljs";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { CJ_UPLOAD_HEADERS, ORIGINAL_ITEM_COL, ORIGINAL_KEY_COL, ORIGINAL_FALLBACK_KEY_COL } from "@/lib/constants/excel";
import { makeDatedFileName } from "@/lib/utils/filename";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

type TRow = Record<string, unknown>;

/**
 * ===============================
 * ✅ 기본 보내는분 fallback 값
 * ===============================
 */
const DEFAULT_SENDER_NAME = "한섬누리";
const DEFAULT_SENDER_TEL = "070-7107-3874";

/**
 * ===============================
 * ✅ 파일명 안전화
 * - 제어문자 제거
 * - 예약문자 치환
 * - 길이 제한
 * - 빈값 fallback
 * ===============================
 */
const sanitizeFileName = (name: string, fallback: string) => {
  const cleaned = String(name ?? "")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim();

  const safe = cleaned.length > 80 ? cleaned.slice(0, 80).trim() : cleaned;
  return safe || fallback;
};

/**
 * ===============================
 * ✅ row에서 후보 컬럼값 찾기
 * - 헤더 공백/줄바꿈 차이 대응
 * - 값이 있으면 반환
 * ===============================
 */
const pick = (row: TRow, candidates: string[]) => {
  if (candidates.length === 0) return "";

  const keyMap = new Map<string, string>();
  for (const k of Object.keys(row)) {
    keyMap.set(normalizeHeader(k), k);
  }

  for (const c of candidates) {
    const actualKey = keyMap.get(normalizeHeader(c)) ?? c;
    const value = toText(row[actualKey]).trim();
    if (value) return value;
  }

  return "";
};

/**
 * ===============================
 * ✅ CJ 고객주문번호 결정
 * ===============================
 */
const getOrderKeyForCj = (row: TRow) => {
  const primary = pick(row, [ORIGINAL_KEY_COL, "상품주문번호"]);
  if (primary) return primary;

  const fallback = pick(row, [ORIGINAL_FALLBACK_KEY_COL, "★쇼핑몰 주문번호★"]);
  if (fallback) return fallback;

  return "";
};

/**
 * ===============================
 * ✅ CJ 컬럼 매핑 테이블
 * ===============================
 */
const CJ_VALUE_MAP: Record<string, string[]> = {
  // ===== 받는분(수취인) =====
  받는분성명: ["받는분성명", "수취인명", "수령인", "수취인"],
  받는분전화번호: ["받는분전화번호", "수취인전화번호1", "수취인전화번호2", "수취인휴대폰", "수취인연락처"],
  받는분우편번호: ["받는분우편번호", "수취인우편번호(2)", "수취인우편번호"],
  받는분주소: ["받는분주소", "수취인주소", "주소", "수령지주소"],
  배송메세지: ["배송메세지", "배송메시지", "배송요청사항", "요청사항"],

  // ===== 품목 / 수량 =====
  품목명: [ORIGINAL_ITEM_COL, "상품명", "품목명"],
  박스수량: ["박스수량", "수량"],

  // ===== 주문번호 =====
  거래처주문번호: ["사방넷 주문번호", "사방넷주문번호", "주문번호"],
  상품주문번호: [ORIGINAL_KEY_COL, "상품주문번호", ORIGINAL_FALLBACK_KEY_COL, "★쇼핑몰 주문번호★"],

  // ===== 보내는분(주문자 매핑) =====
  보내는분성명: ["주문자", "주문자명", "주문자성명", "구매자", "구매자명"],
  보내는분전화번호: [
    // ✅ 사방넷 양식 우선순위: 1 → 2 → (기타 후보들)
    "주문자전화번호1",
    "주문자전화번호2",

    // 기존 후보들
    "주문자전화번호",
    "주문자연락처",
    "주문자휴대폰",
    "구매자전화번호",
    "구매자연락처",
    "구매자휴대폰",
  ],
  보내는분우편번호: ["주문자우편번호", "구매자우편번호"],
  보내는분주소: ["주문자주소", "구매자주소"],
};

/**
 * ===============================
 * ✅ CJ 업로드용 그룹 생성
 * ===============================
 */
export const buildCjGroupedRows = (originalHeaders: string[], originalRows: TRow[]) => {
  const groups = new Map<string, TRow[]>();

  for (const row of originalRows) {
    const itemName = pick(row, [ORIGINAL_ITEM_COL, "상품명", "품목명"]).trim();

    if (!itemName) continue;

    const orderKey = getOrderKeyForCj(row);
    if (!orderKey) continue;

    const arr = groups.get(itemName) ?? [];
    const out: TRow = {};

    for (const cjHeader of CJ_UPLOAD_HEADERS) {
      const normalized = normalizeHeader(cjHeader);

      // 고객주문번호는 강제 주입
      if (normalized === normalizeHeader("고객주문번호")) {
        out[cjHeader] = orderKey;
        continue;
      }

      // 품목명은 옵션 붙이기
      if (normalized === normalizeHeader("품목명")) {
        const base = pick(row, CJ_VALUE_MAP["품목명"]);
        const option = pick(row, ["옵션", "상품옵션"]);
        out[cjHeader] = option ? `${base} / ${option}` : base;
        continue;
      }

      const candidates = CJ_VALUE_MAP[cjHeader] ?? [cjHeader];
      let value = pick(row, candidates);

      // ===== 보내는분 기본값 처리 =====
      if (!value) {
        if (normalized === normalizeHeader("보내는분성명")) {
          value = DEFAULT_SENDER_NAME;
        }
        if (normalized === normalizeHeader("보내는분전화번호")) {
          value = DEFAULT_SENDER_TEL;
        }
      }

      out[cjHeader] = value;
    }

    arr.push(out);
    groups.set(itemName, arr);
  }

  return groups;
};

/**
 * ===============================
 * ✅ 엑셀 워크북 생성
 * ===============================
 */
const makeCjWorkbookBuffer = async (rows: TRow[]) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.columns = CJ_UPLOAD_HEADERS.map((h) => ({
    header: h,
    key: h,
    width: h.length > 8 ? 18 : 14,
  }));

  for (const r of rows) {
    const rowData: Record<string, unknown> = {};
    for (const h of CJ_UPLOAD_HEADERS) {
      rowData[h] = toText(r[h]);
    }
    ws.addRow(rowData);
  }

  ws.getRow(1).font = { bold: true };

  return wb.xlsx.writeBuffer();
};

/**
 * ===============================
 * ✅ ZIP 다운로드
 * ===============================
 */
export const downloadCjUploadsZip = async (groups: Map<string, TRow[]>, onProgress?: (done: number, total: number) => void) => {
  const zip = new JSZip();
  const entries = Array.from(groups.entries());
  const total = entries.length;

  const usedNames = new Map<string, number>();
  let done = 0;

  for (const [itemName, rows] of entries) {
    const buf = await makeCjWorkbookBuffer(rows);

    const baseSafe = sanitizeFileName(itemName, "품목");
    const count = (usedNames.get(baseSafe) ?? 0) + 1;
    usedNames.set(baseSafe, count);

    const finalName = count === 1 ? baseSafe : `${baseSafe} (${count})`;

    zip.file(`${finalName}.xlsx`, buf);

    done += 1;
    onProgress?.(done, total);
  }

  const zipBlob = await zip.generateAsync({ type: "blob" });
  saveAs(zipBlob, makeDatedFileName("한섬누리_CJ제출용_품목별엑셀.zip"));
};
