import Dexie, { type Table } from "dexie";

export type TOriginalRow = Record<string, unknown>;

export type TAggregateRowCache = {
  itemName: string;
  totalBox: number;
  kg?: number | null;
};

export type TUploadedReplyFile = {
  name: string;
  size: number;
  lastModified: number;
};

export type TJobState = {
  createdAt: string; // ISO
  originalFileName: string;

  // 원본 전체 row를 저장 (새로고침 대비)
  originalRows: TOriginalRow[];
  originalHeaders: string[];
  originalWorkbookBuffer?: ArrayBuffer;

  // “품목명 단순 집계” 결과를 캐시(선택)
  aggregateRows?: TAggregateRowCache[];

  // 회신 업로드 중복 방지용
  uploadedReplyFiles?: TUploadedReplyFile[];
};

class HansomDB extends Dexie {
  job!: Table<TJobState, string>;

  constructor() {
    super("hansom-nuri-db");
    this.version(1).stores({
      job: "createdAt",
    });
  }
}

export const db = new HansomDB();

export const saveJob = async (job: TJobState) => {
  await db.job.clear();
  await db.job.add(job);
};

export const loadJob = async (): Promise<TJobState | null> => {
  const all = await db.job.toArray();
  return all[0] ?? null;
};

export const clearJob = async () => {
  await db.job.clear();
};
