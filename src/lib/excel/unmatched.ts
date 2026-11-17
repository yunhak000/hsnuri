export type TUnmatchedRow = {
  customerOrderNo: string;
  tracking: string;
  source?: "원본" | "회신";
  sourceFileName?: string;
  rowNumber?: number;
  reason?: string;
};
