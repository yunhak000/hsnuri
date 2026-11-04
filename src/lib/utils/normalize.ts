import { formatExcelDateTime, normalizeDateTimeText } from "./datetime";

export const normalizeHeader = (s: unknown) =>
  String(s ?? "")
    .trim()
    .replace(/\s+/g, "");

export const toText = (v: unknown): string => {
  if (v == null) return "";

  // ✅ ExcelJS가 cell.value를 Date로 주는 경우
  if (v instanceof Date) return formatExcelDateTime(v);
  if (typeof v === "string") return normalizeDateTimeText(v);

  // ✅ ExcelJS RichText / { text: ... } 형태 지원 (기존 유지)
  if (typeof v === "object" && v !== null && "text" in v) {
    const maybe = (v as { text?: unknown }).text;
    return toText(maybe);
  }

  // ✅ ExcelJS Formula 결과/하이퍼링크 등에서 { result: ... } 로 오는 경우(방어)
  if (typeof v === "object" && v !== null && "result" in v) {
    const maybe = (v as { result?: unknown }).result;
    return toText(maybe);
  }

  return String(v);
};
