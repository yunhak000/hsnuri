"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import Image from "next/image";
import { readOriginalFiles } from "@/lib/excel/readOriginals";
import {
  buildAggregateRows,
  type TAggregateRow,
} from "@/lib/excel/writeAggregate";
import {
  buildCjGroupedRows,
  downloadCjUploadsZip,
  downloadCjUploadsSingle,
} from "@/lib/excel/writeCJUploads";
import { readCjReplyFiles } from "@/lib/excel/readCJReply";
import {
  downloadOriginalWithTracking,
  downloadUnmatchedExcel,
} from "@/lib/excel/applyTrackingToOriginal";
import { clearWorkspace, loadWorkspace, saveWorkspace, type TWorkspace } from "@/lib/db";
import { getOriginalLabel } from "@/lib/utils/originalLabel";
import { matchWorkspace } from "@/lib/excel/matchWorkspace";
import {
  fingerprintFile,
  isSameFingerprint,
  type TFileFingerprint,
} from "@/lib/utils/hash";
import { normalizeHeader } from "@/lib/utils/normalize";
import { getJejuDeliverySummary } from "@/lib/utils/jejuDelivery";

// 재배포

/**
 * ✅ Row 타입 (any 금지)
 * - 엑셀은 셀 타입이 다양해서 unknown 기반으로 두고,
 *   필요한 시점에 toText / String(...) 등으로 안전 변환하는 방식이 가장 안전함.
 */
type TRow = Record<string, unknown>;

type TStep = 1 | 2 | 3 | 4;

const getErrorMessage = (e: unknown) => {
  if (e instanceof Error) return e.message;
  if (typeof e === "string") return e;
  try {
    return JSON.stringify(e);
  } catch {
    return "알 수 없는 오류가 발생했습니다.";
  }
};

const STEP_META: Array<{ n: TStep; label: string; desc: string }> = [
  { n: 1, label: "원본 업로드", desc: "당일 출고 원본 엑셀을 올려주세요" },
  { n: 2, label: "산출물 생성", desc: "품목별 집계를 확인하고 CJ 파일을 받아주세요" },
  { n: 3, label: "회신 업로드", desc: "CJ 운송장 회신 파일을 올려주세요" },
  { n: 4, label: "최종 다운로드", desc: "운송장이 반영된 결과를 받아주세요" },
];

/**
 * ✅ 공통 UI 클래스 — 한섬누리 스타일 (심플 / 밝은 배경 / 로고 초록 포인트)
 */
const ui = {
  page: "min-h-screen bg-[#f2f4f6] text-[#191f28]",
  container: "mx-auto flex min-h-screen max-w-md flex-col px-5 pt-8 pb-28",
  title: "text-[26px] font-bold leading-snug tracking-tight text-[#191f28]",
  subtitle: "mt-1.5 text-[15px] text-[#8b95a1]",
  cardTitle: "text-[22px] font-bold leading-snug tracking-tight text-[#191f28]",
  cardDesc: "mt-1 text-[15px] text-[#8b95a1]",
  hint: "text-[13px] leading-relaxed text-[#8b95a1]",
  divider: "h-px bg-[#f2f4f6]",
  alertError:
    "whitespace-pre-line rounded-2xl bg-[#fff0f0] px-4 py-3.5 text-[14px] font-medium leading-relaxed text-[#f04452]",
  // Step2 다운로드: 리스트형 카드 버튼
  btnRow: "flex flex-col gap-3",
  listBtn:
    "group flex w-full items-center gap-3.5 rounded-2xl bg-white p-4 text-left shadow-[0_1px_2px_rgba(0,0,0,0.04)] transition-all duration-200 hover:bg-[#f9fafb] active:scale-[0.99] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  listIcon:
    "flex h-11 w-11 shrink-0 items-center justify-center rounded-xl bg-primary-soft text-xl",
  listTitle: "block text-[15px] font-bold text-[#191f28]",
  listDesc: "mt-1 block text-[13px] text-[#8b95a1]",
  // 결과(step4) 보조 버튼
  btnPrimary:
    "inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-primary px-5 py-4 text-[16px] font-bold text-white transition-all duration-200 hover:bg-primary-hover active:scale-[0.98] disabled:bg-primary-muted disabled:text-primary-muted-foreground disabled:cursor-not-allowed disabled:active:scale-100",
  btnSecondary:
    "inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-[#f2f4f6] px-5 py-4 text-[16px] font-bold text-[#4e5968] transition-all duration-200 hover:bg-[#e5e8eb] active:scale-[0.98] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  // 하단 고정 CTA: flex-1로 두어 1개일 땐 100%, 2개일 땐 50:50
  ctaPrimary:
    "flex flex-1 items-center justify-center rounded-2xl bg-primary px-5 py-4 text-[17px] font-bold text-white transition-all duration-200 hover:bg-primary-hover active:scale-[0.98] disabled:bg-primary-muted disabled:text-primary-muted-foreground disabled:cursor-not-allowed disabled:active:scale-100",
  ctaSecondary:
    "flex flex-1 items-center justify-center rounded-2xl bg-[#f2f4f6] px-5 py-4 text-[17px] font-bold text-[#4e5968] transition-all duration-200 hover:bg-[#e5e8eb] active:scale-[0.98] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  // 드롭존
  dropzone:
    "flex flex-col items-center justify-center gap-2 rounded-2xl border-2 border-dashed border-[#d1d6db] bg-white px-5 py-9 text-center transition-colors duration-200 hover:border-primary hover:bg-primary-wash",
  statsGrid: "grid grid-cols-2 gap-2.5",
  stat: "rounded-2xl bg-white p-4",
  statKey: "text-[13px] font-medium text-[#8b95a1]",
  statVal: "mt-1 text-[20px] font-bold text-[#191f28]",
  tableWrap: "max-h-56 overflow-auto rounded-2xl bg-white",
  table: "w-full text-[13px]",
  thead: "sticky top-0 bg-[#f9fafb]",
  th: "p-2.5 text-left font-semibold text-[#6b7684]",
  td: "border-t border-[#f2f4f6] p-2.5 text-[#4e5968]",
  successTag:
    "pop-in flex items-center gap-1.5 text-[14px] font-semibold text-primary",
};

export default function HomePage() {
  const [step, setStep] = useState<TStep>(1);
  // ✅ 2단계: CJ 산출물 다운로드 완료 안내
  const [step2Done, setStep2Done] = useState<boolean>(false);
  const [aggregateCopied, setAggregateCopied] = useState(false);
  const [loading, setLoading] = useState<{ on: boolean; text: string }>({
    on: false,
    text: "",
  });
  const [error, setError] = useState<string>("");

  const [workspace, setWorkspace] = useState<TWorkspace>({ id: "current", jobs: {}, uploadedReplyFiles: [], replyTrackingByOrder: {} });
  const [activeSource, setActiveSource] = useState<string>("");
  const [restoring, setRestoring] = useState(true);
  const job = workspace.jobs[activeSource] ?? null;
  const localResult = workspace.result ?? null;
  const sourceEntries = Object.entries(workspace.jobs);
  const originalLabel = (id: string) => getOriginalLabel(sourceEntries.findIndex(([key]) => key === id));
  const hasReplies = workspace.uploadedReplyFiles.length > 0 || Object.keys(workspace.replyTrackingByOrder).length > 0;
  const replyInputRef = useRef<HTMLInputElement | null>(null);

  // derived
  const aggregateRows: TAggregateRow[] = useMemo(() => {
    if (!job?.originalRows) return [];
    // job.originalRows는 TJobState 타입 상 unknown row일 수 있으니 여기서만 캐스팅
    return buildAggregateRows(job.originalRows as TRow[]);
  }, [job?.originalRows]);

  const jejuDelivery = useMemo(
    () => getJejuDeliverySummary(job?.originalRows ?? []),
    [job?.originalRows],
  );

  /**
   * ✅ 총 수량 계산
   * - "박스수량" 우선
   * - 없으면 "수량"
   * - trim 후 숫자 변환
   */
  const totalQuantity = useMemo(() => {
    if (!job?.originalRows?.length) return 0;

    const rows = job.originalRows as TRow[];

    // 어떤 컬럼을 쓸지 결정
    const hasBoxCol = job.originalHeaders.some(
      (h) => normalizeHeader(h) === normalizeHeader("박스수량"),
    );

    const hasQtyCol = job.originalHeaders.some(
      (h) => normalizeHeader(h) === normalizeHeader("수량"),
    );

    const targetCol = hasBoxCol ? "박스수량" : hasQtyCol ? "수량" : null;

    if (!targetCol) return 0;

    let sum = 0;

    for (const row of rows) {
      const raw = String(row[targetCol] ?? "").trim();

      if (!raw) continue;

      const num = Number(raw.replace(/,/g, "")); // 1,000 같은 값 대비

      if (Number.isFinite(num)) {
        sum += num;
      }
    }

    return sum;
  }, [job?.originalRows, job?.originalHeaders]);

  useEffect(() => {
    (async () => {
      try {
        let saved = await loadWorkspace();
        if (saved.result || Object.keys(saved.replyTrackingByOrder).length > 0) {
          saved = matchWorkspace(saved, new Map());
          await saveWorkspace(saved);
        }
        setWorkspace(saved);
        setActiveSource(Object.keys(saved.jobs)[0] ?? "");
      } catch (e) {
        setError(getErrorMessage(e));
      } finally {
        setRestoring(false);
      }
    })();
  }, []);

  const setBusy = (on: boolean, text = "") => setLoading({ on, text });

  // ✅ 각 단계 완료 여부 (= 다음 단계로 넘어갈 수 있는 조건)
  const step1Done = sourceEntries.some(([, item]) => item.originalRows.length > 0);
  const step3Done = !!localResult;

  const goNext = () => setStep((s) => (s < 4 ? ((s + 1) as TStep) : s));
  const goPrev = () => setStep((s) => (s > 1 ? ((s - 1) as TStep) : s));

  // 현재 단계에서 "다음" 버튼 활성화 조건
  const nextEnabled =
    (step === 1 && step1Done) ||
    (step === 2 && step1Done) ||
    (step === 3 && step3Done);

  /**
   * ✅ 전체 리셋
   * - DB 초기화
   * - React state 초기화
   * - file input value 강제 초기화 (파일명 남는 문제 해결)
   */
  const onReset = async () => {
    if (loading.on) return;
    if (!confirm("모든 작업을 초기화하고 처음부터 시작할까요?")) return;

    try {
      setBusy(true, "작업 초기화 중...");
      await clearWorkspace();

      setWorkspace({ id: "current", jobs: {}, uploadedReplyFiles: [], replyTrackingByOrder: {} });
      setStep2Done(false);
      setAggregateCopied(false);
      setStep(1);
      setError("");

      // ✅ input에 남아있는 "선택된 파일명" 제거
      if (replyInputRef.current) replyInputRef.current.value = "";
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "작업 초기화 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onUploadOriginals = async (files: File[]) => {
    if (!files.length || hasReplies || loading.on || restoring) return;
    try {
      setError("");
      setBusy(true, "원본 엑셀 읽는 중...");
      const incoming = await readOriginalFiles(files, workspace.jobs);
      const sources = Object.keys(incoming);
      const next: TWorkspace = {
        ...workspace,
        jobs: { ...workspace.jobs, ...incoming },
        result: undefined,
      };
      await saveWorkspace(next);
      setWorkspace(next);
      setActiveSource(sources[0]);
      setStep2Done(false);
      setAggregateCopied(false);
    } catch (e: unknown) {
      setError(getErrorMessage(e));
    } finally {
      setBusy(false);
    }
  };

  const onCopyAggregate = async () => {
    try {
      setError("");
      setAggregateCopied(false);
      const text = [
        ["상품분류", "총 박스수량"],
        ...aggregateRows.map((row) => [
          row.itemName.replace(/[\t\r\n]+/g, " "),
          String(row.totalBox),
        ]),
        ["합계", String(aggregateRows.reduce((sum, row) => sum + row.totalBox, 0))],
      ].map((row) => row.join("\t")).join("\n");
      await navigator.clipboard.writeText(text);
      setAggregateCopied(true);
    } catch {
      setError("복사하지 못했어요. 아래 집계 표를 직접 선택해서 복사해 주세요.");
    }
  };

  const onDownloadCjZip = async () => {
    if (!job) return;

    try {
      setError("");
      setBusy(true, "CJ 업로드용 품목별 파일 생성(Zip) 중...");

      const groups = buildCjGroupedRows(
        job.originalHeaders,
        job.originalRows as TRow[],
      );

      await downloadCjUploadsZip(groups);

      setStep2Done(true);
    } catch (e: unknown) {
      setError(
        getErrorMessage(e) || "CJ 업로드용 파일 생성 중 오류가 발생했습니다.",
      );
    } finally {
      setBusy(false);
    }
  };

  const onDownloadCjSingle = async () => {
    if (!job) return;

    try {
      setError("");
      setBusy(true, "CJ 업로드용 통합 엑셀 생성 중...");

      const groups = buildCjGroupedRows(
        job.originalHeaders,
        job.originalRows as TRow[],
      );

      await downloadCjUploadsSingle(groups, job.originalFileName);

      setStep2Done(true);
    } catch (e: unknown) {
      setError(
        getErrorMessage(e) || "CJ 통합 엑셀 생성 중 오류가 발생했습니다.",
      );
    } finally {
      setBusy(false);
    }
  };

  const onUploadReplies = async (files: FileList | null) => {
    if (!step1Done || loading.on) return;
    if (!files || files.length === 0) return;

    try {
      setError("");

      /**
       * ✅ 중복 파일(동일 fingerprint) 업로드 제외
       */
      const existing = workspace.uploadedReplyFiles;
      const newFingerprints: TFileFingerprint[] = [];
      const accepted: File[] = [];
      const dupNames: string[] = [];

      Array.from(files).forEach((f) => {
        const fp = fingerprintFile(f);
        const isDup =
          existing.some((x) => isSameFingerprint(x, fp)) ||
          newFingerprints.some((x) => isSameFingerprint(x, fp));

        if (isDup) {
          dupNames.push(f.name);
          return;
        }

        newFingerprints.push(fp);
        accepted.push(f);
      });

      if (accepted.length === 0) {
        setError(`이미 업로드한 회신 파일입니다: ${dupNames.join(", ")}`);
        return;
      }

      if (dupNames.length > 0) {
        setError(
          `일부 회신 파일은 이미 업로드되어 제외했습니다: ${dupNames.join(", ")}`,
        );
      }

      setBusy(true, "CJ 회신 파일 읽는 중...");
      const { map, unmatched: invalidReplies, skipped } = await readCjReplyFiles(accepted);
      if (skipped.length > 0) {
        setError(skipped.map((item) => `${item.sourceFileName}: ${item.reason}`).join("\n"));
      }

      if (accepted.every((file) => skipped.some((item) => item.sourceFileName === file.name)) && map.size === 0) return;
      setBusy(true, "원본별 운송장번호 매핑 중...");
      const next = matchWorkspace(workspace, map, invalidReplies);
      next.uploadedReplyFiles = [...existing, ...newFingerprints.filter((fp) => !skipped.some((item) => item.sourceFileName === fp.name))];
      await saveWorkspace(next);
      setWorkspace(next);

      setStep(4);
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "회신 처리 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onDownloadFinal = async (source: string) => {
    const job = workspace.jobs[source];
    if (!job) return;

    try {
      setError("");
      setBusy(true, "최종 원본 엑셀 생성 중...");

      await downloadOriginalWithTracking(
        job.originalHeaders,
        job.originalRows as TRow[],
        job.originalFileName,
        job.originalWorkbookBuffer,
        job.replyTrackingByOrder,
        job.suplinSheetName,
      );
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "최종 엑셀 생성 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onDownloadUnmatched = async () => {
    if (!localResult) return;

    try {
      setError("");
      setBusy(true, "미매칭 목록 엑셀 생성 중...");

      await downloadUnmatchedExcel(localResult.unmatched);
    } catch (e: unknown) {
      setError(
        getErrorMessage(e) || "미매칭 엑셀 생성 중 오류가 발생했습니다.",
      );
    } finally {
      setBusy(false);
    }
  };

  const currentMeta = STEP_META[step - 1];
  const progressPct = ((step - 1) / (STEP_META.length - 1)) * 100;

  return (
    <div className={ui.page}>
      <main aria-busy={restoring} className={`${ui.container} ${step === 4 ? "md:max-w-6xl" : ""}`}>
        <header className="mb-6 flex items-center gap-3 border-b border-[#e5e8eb] pb-5">
          <div className="relative h-14 w-14 shrink-0 overflow-hidden mix-blend-multiply">
            <Image
              src="/logo.png"
              alt="한섬누리"
              width={300}
              height={105}
              priority
              className="absolute left-1/2 top-0 h-14 w-40 max-w-none -translate-x-1/2"
            />
          </div>
          <div className="min-w-0 flex-1">
            <p className="text-[15px] font-bold tracking-tight text-[#191f28]">
              한섬누리 출고관리
            </p>
            <p className="mt-1 text-[12px] text-[#6b7684]">
              주문부터 운송장까지, 간편하게
            </p>
          </div>
          <button
            type="button"
            onClick={onReset}
            disabled={loading.on || restoring}
            aria-label="모든 작업 초기화하고 처음부터 다시 시작"
            title="처음부터 다시"
            className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl bg-primary text-white shadow-sm transition-colors hover:bg-primary-hover focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-primary disabled:cursor-not-allowed disabled:bg-primary-muted disabled:text-primary-muted-foreground"
          >
            <svg
              width="22"
              height="22"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              <path d="M3 10a9 9 0 1 1 2.6 8.4" />
              <path d="M3 4v6h6" />
            </svg>
          </button>
        </header>
        {/* ===== 상단: 진행 게이지 ===== */}
        <div className="mb-7">
          <div className="mb-2 flex items-center justify-between">
            <button
              type="button"
              onClick={goPrev}
              disabled={step === 1 || loading.on}
              className="-ml-1 flex h-8 w-8 items-center justify-center rounded-full text-xl text-[#4e5968] transition-colors hover:bg-[#e5e8eb] disabled:opacity-0"
              aria-label="이전"
            >
              ‹
            </button>
            <span className="text-[13px] font-semibold text-[#8b95a1]">
              {step} / {STEP_META.length}
            </span>
          </div>
          <div className="h-1.5 w-full overflow-hidden rounded-full bg-[#e5e8eb]">
            <div
              className="h-full rounded-full bg-primary transition-all duration-500 ease-out"
              style={{ width: `${progressPct}%` }}
            />
          </div>
        </div>

        {/* ===== 단계 헤더 (토스식 큰 타이틀) ===== */}
        <div key={`h-${step}`} className="step-anim mb-6">
          <h1 className={ui.cardTitle}>{currentMeta.label}</h1>
          <p className={ui.cardDesc}>{currentMeta.desc}</p>
        </div>

        {/* ===== Error ===== */}
        {error && <div className={`${ui.alertError} pop-in mb-4`}>{error}</div>}

        {/* ===== Loading Overlay ===== */}
        {loading.on && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#191f28]/25 backdrop-blur-[2px]">
            <div className="flex items-center gap-3 rounded-2xl bg-white px-6 py-5 shadow-xl">
              <div className="h-5 w-5 animate-spin rounded-full border-[2.5px] border-[#e5e8eb] border-t-primary" />
              <div className="text-[15px] font-bold text-[#191f28]">
                {loading.text}
              </div>
            </div>
          </div>
        )}

        {/* ===== 단계 본문 (한 번에 한 단계) ===== */}
        <div key={`b-${step}`} className="step-anim flex-1 space-y-5">
          {/* Step 1 */}
          {step === 1 && (
            <>
              <p className={ui.hint}>원본 엑셀을 여러 개 선택하세요. 올린 순서대로 첫 번째·두 번째·세 번째 엑셀로 관리해요. 한 개만 올려도 돼요.</p>
              <label className={`${ui.dropzone} ${hasReplies ? "opacity-60" : ""}`}>
                <span className={ui.listTitle}>원본 엑셀 올리기</span>
                <span className={ui.hint}>.xlsx 파일을 함께 선택하거나 나중에 추가하세요</span>
                <input type="file" accept=".xlsx" multiple className="hidden" disabled={loading.on || restoring || hasReplies}
                  onChange={(e) => { const files = Array.from(e.target.files ?? []); e.target.value = ""; void onUploadOriginals(files); }} />
              </label>
              {sourceEntries.map(([source, sourceJob]) => (
                <div key={source} className="space-y-1 rounded-2xl bg-white p-4" role="status">
                  <p className={ui.listTitle}>{originalLabel(source)}</p>
                  <p className="break-all text-[13px] text-[#6b7684]">{sourceJob.originalFileName}</p>
                  <p className={ui.hint}>{sourceJob.originalRows.length.toLocaleString()}행 · 업로드 완료</p>
                </div>
              ))}
              {hasReplies && <p className={ui.hint}>회신을 반영한 뒤에는 원본을 추가할 수 없어요. 새 작업은 상단 초기화 버튼으로 시작해주세요.</p>}
            </>
          )}

          {/* Step 2 */}
          {step === 2 && (
            <>
              <div className="flex flex-wrap gap-2" aria-label="CJ 제출용 원본 엑셀 선택">
                {sourceEntries.map(([source]) => (
                  <button key={source} type="button" aria-pressed={activeSource === source}
                    className={`flex-1 rounded-xl px-4 py-3 font-bold ${activeSource === source ? "bg-primary text-white" : "bg-white text-[#4e5968]"}`}
                    onClick={() => { setActiveSource(source); setAggregateCopied(false); setStep2Done(false); }}>
                    {originalLabel(source)}
                  </button>
                ))}
              </div>
              <p className={ui.hint}>{originalLabel(activeSource)} · {job?.originalFileName} · 총 수량 {totalQuantity.toLocaleString()}개</p>
              {jejuDelivery.hasJejuDelivery && (
                <div role="alert" className="rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3.5 text-[15px] font-bold leading-relaxed text-amber-900">
                  ★ 주의하세요. 제주도 배송이 &apos;{jejuDelivery.totalBox.toLocaleString("ko-KR") }건&apos; 있습니다.
                </div>
              )}
              <section className="overflow-hidden rounded-2xl bg-white shadow-[0_1px_2px_rgba(0,0,0,0.04)]" aria-labelledby="aggregate-title">
                <div className="flex items-center justify-between gap-3 p-4">
                  <div>
                    <h2 id="aggregate-title" className={ui.listTitle}>품목별 집계</h2>
                    <p className={ui.listDesc}>복사해서 필요한 곳에 붙여넣으세요.</p>
                  </div>
                  <button
                    type="button"
                    className="shrink-0 rounded-xl bg-primary-soft px-3 py-2 text-[13px] font-bold text-primary disabled:opacity-40"
                    disabled={loading.on || aggregateRows.length === 0}
                    onClick={onCopyAggregate}
                  >
                    전체 복사
                  </button>
                </div>
                {aggregateRows.length > 0 ? (
                  <div className="max-h-96 overflow-auto">
                    <table className={`${ui.table} select-text`}>
                      <thead className={ui.thead}>
                        <tr>
                          <th scope="col" className={ui.th}>상품분류</th>
                          <th scope="col" className={`${ui.th} whitespace-nowrap text-right`}>총 박스수량</th>
                        </tr>
                      </thead>
                      <tbody>
                        {aggregateRows.map((row) => (
                          <tr key={row.itemName}>
                            <td className={`${ui.td} break-words`}>{row.itemName}</td>
                            <td className={`${ui.td} text-right tabular-nums`}>{row.totalBox}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr className="bg-[#f9fafb] font-bold">
                          <th scope="row" className={ui.td}>합계</th>
                          <td className={`${ui.td} text-right tabular-nums`}>{aggregateRows.reduce((sum, row) => sum + row.totalBox, 0)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                ) : (
                  <p className="px-4 pb-4 text-[13px] text-[#8b95a1]">집계할 품목이 없습니다.</p>
                )}
                <p role="status" className="px-4 pb-3 text-[13px] text-primary">
                  {aggregateCopied ? "복사했어요. 원하는 곳에 붙여넣으세요." : ""}
                </p>
              </section>

              <div className={ui.btnRow}>
                <button
                  className={ui.listBtn}
                  disabled={loading.on || restoring}
                  onClick={onDownloadCjZip}
                >
                  <span className={ui.listIcon}>🗂️</span>
                  <span className="flex-1">
                    <span className={ui.listTitle}>{originalLabel(activeSource)} CJ 품목별 ZIP</span>
                    <span className={ui.listDesc}>품목별로 파일이 나뉨</span>
                  </span>
                  <span className="text-[#c4cad1]">›</span>
                </button>

                <button
                  className={ui.listBtn}
                  disabled={loading.on || restoring}
                  onClick={onDownloadCjSingle}
                >
                  <span className={ui.listIcon}>📄</span>
                  <span className="flex-1">
                    <span className={ui.listTitle}>{originalLabel(activeSource)} CJ 통합 엑셀</span>
                    <span className={ui.listDesc}>하나의 파일로 합쳐짐</span>
                  </span>
                  <span className="text-[#c4cad1]">›</span>
                </button>
              </div>

              <p className={ui.hint}>
                집계를 확인한 뒤 다음 단계로 넘어갈 수 있어요. 집계는
                “상품약어 → 품목명 → 상품명” 순으로 분류해 합산합니다.
              </p>

              {step2Done && (
                <div className={ui.successTag}>
                  <span className="flex h-5 w-5 items-center justify-center rounded-full bg-primary text-[11px] text-white">
                    ✓
                  </span>
                  산출물을 받았어요. 다음으로 진행하세요.
                </div>
              )}
            </>
          )}

          {/* Step 3 */}
          {step === 3 && (
            <>
              <label className={ui.dropzone}>
                <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-primary-soft text-3xl">
                  📦
                </div>
                <div className="text-[15px] font-bold text-[#191f28]">
                  CJ 회신 엑셀 올리기
                </div>
                <div className="text-[13px] text-[#8b95a1]">
                  모든 원본의 CJ 회신을 한 번에 선택하세요
                </div>
                <input
                  ref={replyInputRef}
                  type="file"
                  accept=".xlsx"
                  multiple
                  disabled={loading.on || restoring}
                  className="hidden"
                  onChange={(e) => { void onUploadReplies(e.target.files); e.target.value = ""; }}
                />
              </label>

              <div className="rounded-2xl bg-white p-4">
                <p className={ui.hint}>
                  • 같은 파일을 다시 올리면 경고 후 제외돼요.
                  <br />• 같은 주문번호의 운송장은 중복을 제외하고 모아 반영해요.
                  <br />• 여러 원본에 겹치는 주문번호는 자동 반영을 보류하고 미매칭 목록에 표시해요.
                </p>
              </div>

              {step3Done && (
                <div className={ui.successTag}>
                  <span className="flex h-5 w-5 items-center justify-center rounded-full bg-primary text-[11px] text-white">
                    ✓
                  </span>
                  회신이 반영됐어요. 다음으로 진행하세요.
                </div>
              )}
            </>
          )}

          {/* Step 4 */}
          {step === 4 && (
            <>
              {localResult ? (
                <div className={ui.statsGrid}>
                  <div className={ui.stat}>
                    <div className={ui.statKey}>원본 총 행 수</div>
                    <div className={ui.statVal}>{sourceEntries.reduce((sum, [, item]) => sum + item.originalRows.length, 0)}</div>
                  </div>
                  <div className={ui.stat}>
                    <div className={ui.statKey}>회신 키 수</div>
                    <div className={ui.statVal}>
                      {localResult.totalReplyCount}
                    </div>
                  </div>
                  <div className={ui.stat}>
                    <div className={ui.statKey}>매칭 주문번호</div>
                    <div className={`${ui.statVal} text-primary`}>
                      {localResult.matchedCount}
                    </div>
                  </div>
                  <div className={ui.stat}>
                    <div className={ui.statKey}>미매칭</div>
                    <div
                      className={`${ui.statVal} ${localResult.unmatched.length > 0 ? "text-[#f04452]" : ""}`}
                    >
                      {localResult.unmatched.length}
                    </div>
                  </div>
                  <div className={`${ui.stat} col-span-2`}>
                    <div className={ui.statKey}>원본 중복키</div>
                    <div className={ui.statVal}>
                      {localResult.duplicates.length}
                    </div>
                  </div>
                </div>
              ) : (
                <div className="rounded-2xl bg-white p-5 text-center text-[14px] text-[#8b95a1]">
                  회신 업로드 후 결과가 표시됩니다.
                </div>
              )}

              <div className={ui.btnRow}>
                {sourceEntries.map(([source, sourceJob]) => (
                  <div key={source} className="space-y-2 rounded-2xl bg-white p-4">
                    <p className={ui.listTitle}>{originalLabel(source)}</p>
                    <p className={ui.hint}>{sourceJob.originalFileName}</p>
                    <p className={ui.hint}>매칭 {localResult?.sources[source]?.matchedCount ?? 0}건 · 미매칭 {localResult?.sources[source]?.unmatchedCount ?? 0}건</p>
                    <button className={ui.btnPrimary} disabled={loading.on || !step3Done}
                      onClick={() => onDownloadFinal(source)}>{originalLabel(source)} 운송장 반영본 다운로드</button>
                  </div>
                ))}

                <button
                  className={ui.btnSecondary}
                  disabled={
                    !localResult ||
                    localResult.unmatched.length === 0 ||
                    loading.on
                  }
                  onClick={onDownloadUnmatched}
                  title={
                    !localResult || localResult.unmatched.length === 0
                      ? "미매칭이 없습니다"
                      : ""
                  }
                >
                  미매칭 목록 다운로드
                </button>
              </div>

              {/* Unmatched preview */}
              {localResult?.unmatched?.length ? (
                <div className="space-y-2">
                  <div className="px-1 text-[14px] font-bold text-[#191f28]">
                    미매칭 리스트{" "}
                    <span className="text-[#8b95a1]">(상위 20개)</span>
                  </div>
                  <div className="max-h-[32rem] space-y-3 overflow-y-auto md:hidden">
                    {localResult.unmatched.slice(0, 20).map((u, index) => (
                      <article key={index} className="rounded-2xl border border-[#e5e8eb] bg-white p-4 text-[14px]">
                        <dl className="space-y-3">
                          <div>
                            <dt className="text-[12px] font-semibold text-[#6b7684]">고객주문번호</dt>
                            <dd className="mt-1 break-all font-semibold text-[#191f28]">{u.customerOrderNo || "(없음)"}</dd>
                          </div>
                          <div>
                            <dt className="text-[12px] font-semibold text-[#6b7684]">운송장번호</dt>
                            <dd className="mt-1 break-all">{u.tracking || "(없음)"}</dd>
                          </div>
                          <div>
                            <dt className="text-[12px] font-semibold text-[#6b7684]">파일 / 행</dt>
                            <dd className="mt-1 break-words [overflow-wrap:anywhere]">{u.source} {u.sourceFileName}{u.rowNumber != null ? ` / ${u.rowNumber}행` : ""}</dd>
                          </div>
                          <div className="rounded-xl bg-[#fff0f0] p-3">
                            <dt className="text-[12px] font-semibold text-[#b3261e]">미매칭 사유</dt>
                            <dd className="mt-1 break-words leading-relaxed text-[#b3261e] [overflow-wrap:anywhere]">{u.reason}</dd>
                          </div>
                        </dl>
                      </article>
                    ))}
                  </div>
                  <div className="hidden max-h-[32rem] overflow-auto rounded-2xl border border-[#e5e8eb] bg-white md:block">
                    <table className="w-full min-w-[800px] table-fixed text-[14px]">
                      <colgroup>
                        <col className="w-[20%]" />
                        <col className="w-[18%]" />
                        <col className="w-[27%]" />
                        <col className="w-[35%]" />
                      </colgroup>
                      <thead className={`${ui.thead} z-10`}>
                        <tr>
                          <th className={`${ui.th} whitespace-nowrap`}>고객주문번호</th>
                          <th className={`${ui.th} whitespace-nowrap`}>운송장번호</th>
                          <th className={`${ui.th} whitespace-nowrap`}>파일 / 행</th>
                          <th className={`${ui.th} whitespace-nowrap`}>미매칭 사유</th>
                        </tr>
                      </thead>
                      <tbody>
                        {localResult.unmatched.slice(0, 20).map((u, index) => (
                          <tr key={index} className="align-top even:bg-[#f9fafb]">
                            <td className={`${ui.td} break-all font-medium`}>{u.customerOrderNo || "(없음)"}</td>
                            <td className={`${ui.td} break-all font-medium`}>{u.tracking || "(없음)"}</td>
                            <td className={`${ui.td} break-words leading-relaxed [overflow-wrap:anywhere]`}>{u.source} {u.sourceFileName}{u.rowNumber != null ? ` / ${u.rowNumber}행` : ""}</td>
                            <td className={`${ui.td} break-words leading-relaxed [overflow-wrap:anywhere]`}><span className="font-medium text-[#b3261e]">{u.reason}</span></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : null}

              {/* Duplicate key preview */}
              {localResult?.duplicates?.length ? (
                <div className="space-y-2">
                  <div className="px-1 text-[14px] font-bold text-[#191f28]">
                    원본 중복키{" "}
                    <span className="text-[#8b95a1]">(상위 20개)</span>
                  </div>
                  <div className={ui.tableWrap}>
                    <table className={ui.table}>
                      <thead className={ui.thead}>
                        <tr>
                          <th className={ui.th}>상품주문번호</th>
                          <th className={ui.th}>중복 개수</th>
                        </tr>
                      </thead>
                      <tbody>
                        {localResult.duplicates.slice(0, 20).map((d) => (
                          <tr key={d.key}>
                            <td className={ui.td}>{d.key}</td>
                            <td className={ui.td}>{d.count}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : null}
            </>
          )}
        </div>

        <footer className="mt-8 space-y-2 border-t border-[#e5e8eb] pt-5 text-center text-[12px] leading-relaxed text-[#6b7684]">
          <p>한섬누리에서 운영하는 출고관리 서비스입니다.</p>
          <p>
            파일은 서버로 전송되지 않으며,
            <br />
            모든 처리는 브라우저에서 이루어집니다.
          </p>
        </footer>
      </main>

      {/* ===== 하단 고정 CTA 바 ===== */}
      <div className="fixed inset-x-0 bottom-0 z-40 bg-gradient-to-t from-[#f2f4f6] via-[#f2f4f6] to-transparent pt-5">
        <div className="mx-auto max-w-md px-5 pb-5">
          <div className="flex gap-2.5">
            {step > 1 && (
              <button
                type="button"
                className={ui.ctaSecondary}
                disabled={loading.on || restoring}
                onClick={goPrev}
              >
                이전
              </button>
            )}

            {step < 4 && (
              <button
                type="button"
                className={ui.ctaPrimary}
                disabled={!nextEnabled || loading.on}
                onClick={goNext}
                title={!nextEnabled ? "이 단계를 먼저 완료해주세요" : ""}
              >
                다음
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
