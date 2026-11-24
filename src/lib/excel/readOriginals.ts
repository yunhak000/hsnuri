import type { TJobState } from "@/lib/db";
import { getOriginalOrderKey } from "@/lib/utils/orderKey";
import { normalizeHeader } from "@/lib/utils/normalize";
import { fingerprintFile, isSameFingerprint } from "@/lib/utils/hash";
import { readExcelFile } from "./read";

// Validate all files before appending them in selection order. Originals of the
// same format are independent jobs, and any number of originals is supported.
export const readOriginalFiles = async (files: File[], existing: Record<string, TJobState> = {}) => {
  if (!files.length) throw new Error("원본 엑셀을 선택해주세요.");
  const jobs: Record<string, TJobState> = {};
  let sequence = Object.keys(existing).length + 1;
  for (const file of files) {
    const fingerprint = fingerprintFile(file);
    if ([...Object.values(existing), ...Object.values(jobs)].some((job) =>
      job.originalFingerprint && isSameFingerprint(job.originalFingerprint, fingerprint))) {
      throw new Error(`${file.name}: 이미 선택하거나 업로드한 원본 파일입니다. 중복 파일을 제외하고 다시 선택해주세요.`);
    }
    const { headers, rows, rowNumbers, originalWorkbookBuffer, suplinSheetName } = await readExcelFile(file);
    const names = new Set(headers.map(normalizeHeader));
    if (!["품목명", "상품명", "상품약어"].some((header) => names.has(normalizeHeader(header)))) {
      throw new Error(`${file.name}: 품목명 또는 상품명 열을 찾지 못했습니다. 주문 원본 엑셀을 확인해주세요.`);
    }
    if (!rows.length || rows.every((row) => !getOriginalOrderKey(row))) {
      throw new Error(`${file.name}: 주문번호를 찾지 못했습니다. 원본 엑셀을 확인해주세요.`);
    }
    while (existing[`original-${sequence}`] || jobs[`original-${sequence}`]) sequence += 1;
    jobs[`original-${sequence++}`] = {
      createdAt: new Date().toISOString(), originalFileName: file.name,
      originalFingerprint: fingerprint,
      originalHeaders: headers, originalRowNumbers: rowNumbers,
      originalRows: rows, originalWorkbookBuffer, suplinSheetName, uploadedReplyFiles: [],
    };
  }
  return jobs;
};
