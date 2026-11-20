import type { TSource, TWorkspace } from "@/lib/db";
import { SOURCE_LABELS } from "@/lib/constants/sources";
import { getOriginalOrderKey } from "@/lib/utils/orderKey";
import { applyTracking } from "./applyTrackingToOriginal";
import type { TUnmatchedRow } from "./unmatched";

// Only route a reply when exactly one source owns the order number.
export const matchWorkspace = (
  workspace: TWorkspace,
  replies: Map<string, string[]>,
  invalidReplies: TUnmatchedRow[] = [],
): TWorkspace => {
  const owners = new Map<string, Set<TSource>>();
  const entries = Object.entries(workspace.jobs) as Array<[TSource, NonNullable<TWorkspace["jobs"][TSource]>]>;
  for (const [source, job] of entries) {
    for (const row of job.originalRows) {
      const key = getOriginalOrderKey(row);
      if (!key) continue;
      const sources = owners.get(key) ?? new Set<TSource>();
      sources.add(source);
      owners.set(key, sources);
    }
  }

  const combined = new Map(Object.entries(workspace.replyTrackingByOrder));
  for (const [key, values] of replies) {
    combined.set(key, [...new Set([...(combined.get(key) ?? []), ...values])]);
  }
  const replyIssues = [...(workspace.replyIssues ?? []), ...invalidReplies];
  const unmatched: TUnmatchedRow[] = [...replyIssues];
  const duplicates: Array<{ key: string; count: number }> = [];
  const jobs = { ...workspace.jobs };
  const sources: NonNullable<TWorkspace["result"]>["sources"] = {};
  let matchedCount = 0;

  for (const [source, job] of entries) {
    const routed = new Map([...combined].filter(([key]) => owners.get(key)?.size === 1 && owners.get(key)?.has(source)));
    const applied = applyTracking(job.originalHeaders, job.originalRows, routed, {
      fileName: `${SOURCE_LABELS[source]} / ${job.originalFileName}`,
      rowNumbers: job.originalRowNumbers,
    });
    const sourceUnmatched = applied.unmatched.map((row) => owners.get(row.customerOrderNo)?.size === 2
      ? { ...row, reason: "사방넷·어드민에 같은 주문번호가 있어 자동 반영 보류" }
      : row);
    unmatched.push(...sourceUnmatched);
    duplicates.push(...applied.duplicates.map((item) => ({ ...item, key: `${SOURCE_LABELS[source]} / ${item.key}` })));
    jobs[source] = { ...job, originalHeaders: applied.updatedHeaders, originalRows: applied.updatedRows, replyTrackingByOrder: Object.fromEntries(routed) };
    sources[source] = { matchedCount: routed.size, unmatchedCount: sourceUnmatched.length };
    matchedCount += routed.size;
  }

  for (const [key, values] of combined) {
    if (owners.get(key)?.size === 1) continue;
    for (const tracking of values) unmatched.push({
      customerOrderNo: key, tracking, source: "회신",
      reason: owners.has(key) ? "사방넷·어드민에 같은 주문번호가 있어 자동 반영 보류" : "두 원본에 없는 주문번호",
    });
  }
  return {
    ...workspace, jobs, replyIssues, replyTrackingByOrder: Object.fromEntries(combined),
    result: { unmatched, duplicates, sources, matchedCount, totalReplyCount: combined.size },
  };
};
