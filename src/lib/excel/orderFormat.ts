import { normalizeHeader } from "@/lib/utils/normalize";
import { ORIGINAL_ORDER_KEY_COLS } from "@/lib/constants/excel";
import type { TSource } from "@/lib/constants/sources";

export const isSabangnetHeaders = (headers: string[]) => {
  const names = new Set(headers.map(normalizeHeader));
  return ["사방넷 주문번호", "쇼핑몰명", "상품약어", "★쇼핑몰 주문번호★"]
    .every((header) => names.has(normalizeHeader(header)));
};

export const detectOrderSource = (headers: string[]): TSource | null => {
  if (isSabangnetHeaders(headers)) return "sabangnet";
  const names = new Set(headers.map(normalizeHeader));
  const has = (header: string) => names.has(normalizeHeader(header));
  // A partial Sabangnet format or a CJ export must not become an admin order.
  if (has("사방넷 주문번호") || has("상품약어")) return null;
  if (["기타", "어드민품목명", "거래처주문번호"].every(has)) return null;
  return has("품목명") && ORIGINAL_ORDER_KEY_COLS.some(has) ? "admin" : null;
};
