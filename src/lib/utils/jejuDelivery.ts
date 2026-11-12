import { normalizeHeader, toText } from "@/lib/utils/normalize";

export const getJejuDeliverySummary = (rows: Record<string, unknown>[]) => {
  let hasJejuDelivery = false;
  let totalBox = 0;

  for (const row of rows) {
    const addressKey = Object.keys(row).find((key) => normalizeHeader(key) === "받는분주소");
    const address = toText(addressKey ? row[addressKey] : undefined);
    if (!/제주도|제주특별자치도/.test(address)) continue;

    hasJejuDelivery = true;
    const quantityKey = Object.keys(row).find((key) => normalizeHeader(key) === "박스수량");
    const quantity = Number(toText(quantityKey ? row[quantityKey] : undefined).trim().replace(/,/g, ""));
    if (Number.isFinite(quantity)) totalBox += quantity;
  }

  return { hasJejuDelivery, totalBox };
};
