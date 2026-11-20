import type { TJobState } from "@/lib/db";
import { SOURCE_LABELS, type TSource } from "@/lib/constants/sources";
import { getOriginalOrderKey } from "@/lib/utils/orderKey";
import { readExcelFile } from "./read";
import { detectOrderSource } from "./orderFormat";

// Validate the entire selection before replacing any saved originals.
export const readOriginalFiles = async (files: File[]) => {
  if (!files.length || files.length > 2) {
    throw new Error("사방넷·어드민 원본을 종류별 한 개씩, 최대 두 개 선택해주세요.");
  }
  const jobs: Partial<Record<TSource, TJobState>> = {};
  for (const file of files) {
    const { headers, rows, rowNumbers, originalWorkbookBuffer } = await readExcelFile(file);
    const source = detectOrderSource(headers);
    if (!source) {
      throw new Error(`${file.name}: 사방넷·어드민 원본 양식을 구분하지 못했습니다. 원본 파일의 열 제목을 확인해주세요.`);
    }
    if (jobs[source]) {
      throw new Error(`${SOURCE_LABELS[source]} 파일이 두 개 선택되었습니다. 종류별 한 개씩 선택해주세요. 기존 원본은 변경하지 않았습니다.`);
    }
    if (!rows.length || rows.every((row) => !getOriginalOrderKey(row))) {
      throw new Error(`${file.name}: 주문번호를 찾지 못했습니다. 원본 엑셀을 확인해주세요.`);
    }
    jobs[source] = {
      createdAt: new Date().toISOString(), originalFileName: file.name,
      originalHeaders: headers, originalRowNumbers: rowNumbers,
      originalRows: rows, originalWorkbookBuffer, uploadedReplyFiles: [],
    };
  }
  return jobs;
};
