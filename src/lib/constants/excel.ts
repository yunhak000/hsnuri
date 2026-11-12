export const CJ_UPLOAD_HEADERS = [
  "기타",
  "보내는분성명",
  "보내는분전화번호",
  "보내는분우편번호",
  "보내는분주소",
  "품목명",
  "박스수량",
  "받는분성명",
  "받는분전화번호",
  "받는분우편번호",
  "받는분주소",
  "배송메세지",
  "거래처주문번호",
  "운송장번호",
  "상품주문번호",
  "어드민품목명",
  "매출처",
] as const;

export type TCjUploadHeader = (typeof CJ_UPLOAD_HEADERS)[number];

export const ORIGINAL_KEY_COL = "상품주문번호";
export const ORIGINAL_ITEM_COL = "품목명";
export const ORIGINAL_BOX_COL = "박스수량";
export const ORIGINAL_TRACKING_COL = "운송장번호";

export const CJ_REPLY_KEY_COL = "고객주문번호";
export const CJ_REPLY_TRACKING_COL = "운송장번호";

export const ORIGINAL_FALLBACK_KEY_COL = "★쇼핑몰 주문번호★";
