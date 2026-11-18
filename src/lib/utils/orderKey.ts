import { ORIGINAL_ORDER_KEY_COLS } from "@/lib/constants/excel";
import { normalizeHeader, toText } from "@/lib/utils/normalize";

type TRow = Record<string, unknown>;

// 업로드, 회신 매칭, 원본 다운로드에서 동일한 우선순위를 사용한다.
export const getOriginalOrderKey = (row: TRow): string => {
  const headers = new Map(Object.keys(row).map((header) => [normalizeHeader(header), header]));
  for (const candidate of ORIGINAL_ORDER_KEY_COLS) {
    const header = headers.get(normalizeHeader(candidate));
    const value = toText(header ? row[header] : undefined).trim();
    if (value) return value;
  }
  return "";
};
