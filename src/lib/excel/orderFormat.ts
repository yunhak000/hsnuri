import { normalizeHeader } from "@/lib/utils/normalize";

export const isSabangnetHeaders = (headers: string[]) => {
  const names = new Set(headers.map(normalizeHeader));
  return ["사방넷 주문번호", "쇼핑몰명", "상품약어", "★쇼핑몰 주문번호★"]
    .every((header) => names.has(normalizeHeader(header)));
};
