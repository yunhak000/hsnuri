import Dexie, { type Table } from "dexie";
import type { TUnmatchedRow } from "@/lib/excel/unmatched";

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
  originalFingerprint?: TUploadedReplyFile;

  // 원본 전체 row를 저장 (새로고침 대비)
  originalRows: TOriginalRow[];
  originalHeaders: string[];
  originalRowNumbers?: number[];
  originalWorkbookBuffer?: ArrayBuffer;
  replyTrackingByOrder?: Record<string, string[]>;

  // “품목명 단순 집계” 결과를 캐시(선택)
  aggregateRows?: TAggregateRowCache[];

  // 회신 업로드 중복 방지용
  uploadedReplyFiles?: TUploadedReplyFile[];
};

export type TWorkspace = {
  id: string;
  jobs: Record<string, TJobState>;
  uploadedReplyFiles: TUploadedReplyFile[];
  replyTrackingByOrder: Record<string, string[]>;
  replyIssues?: TUnmatchedRow[];
  result?: {
    unmatched: TUnmatchedRow[];
    duplicates: Array<{ key: string; count: number }>;
    matchedCount: number;
    totalReplyCount: number;
    sources: Record<string, { matchedCount: number; unmatchedCount: number }>;
  };
};

class HansomDB extends Dexie {
  job!: Table<TJobState, string>;

  workspace!: Table<TWorkspace, string>;

  constructor() {
    super("hansom-nuri-db");
    this.version(1).stores({ job: "createdAt" });
    this.version(2).stores({ job: "createdAt", workspace: "id" });
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

export const saveWorkspace = async (workspace: TWorkspace) => {
  await db.workspace.put(workspace);
};

export const loadWorkspace = async (): Promise<TWorkspace> => {
  const saved = await db.workspace.get("current");
  if (saved) return saved;
  const legacy = await loadJob();
  const workspace: TWorkspace = {
    id: "current",
    jobs: legacy ? { "original-1": legacy } : {},
    uploadedReplyFiles: legacy?.uploadedReplyFiles ?? [],
    replyTrackingByOrder: legacy?.replyTrackingByOrder ?? {},
  };
  await saveWorkspace(workspace);
  return workspace;
};

export const clearWorkspace = async () => {
  await db.transaction("rw", db.job, db.workspace, async () => {
    await db.job.clear();
    await db.workspace.clear();
  });
};
