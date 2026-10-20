"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { readExcelFile } from "@/lib/excel/read";
import {
  buildAggregateRows,
  downloadAggregateExcel,
  type TAggregateRow,
} from "@/lib/excel/writeAggregate";
import {
  buildCjGroupedRows,
  downloadCjUploadsZip,
  downloadCjUploadsSingle,
} from "@/lib/excel/writeCJUploads";
import { readCjReplyFiles } from "@/lib/excel/readCJReply";
import {
  applyTracking,
  downloadOriginalWithTracking,
  downloadUnmatchedExcel,
} from "@/lib/excel/applyTrackingToOriginal";
import { clearJob, loadJob, saveJob, type TJobState } from "@/lib/db";
import {
  fingerprintFile,
  isSameFingerprint,
  type TFileFingerprint,
} from "@/lib/utils/hash";
import { normalizeHeader } from "@/lib/utils/normalize";

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
  { n: 2, label: "산출물 생성", desc: "필요한 산출물을 최소 1개 받아주세요" },
  { n: 3, label: "회신 업로드", desc: "CJ 운송장 회신 파일을 올려주세요" },
  { n: 4, label: "최종 다운로드", desc: "운송장이 반영된 결과를 받아주세요" },
];

/**
 * ✅ 공통 UI 클래스 (가독성 + 재사용)
 * - 심플 + 스무스한 마법사 UI 기준으로 재정리
 */
const ui = {
  page: "min-h-screen bg-gradient-to-b from-slate-50 via-white to-slate-50 text-slate-900",
  container: "mx-auto max-w-2xl px-5 py-10 md:py-14 space-y-8",
  headerWrap: "space-y-1.5 text-center",
  title: "text-2xl md:text-3xl font-bold tracking-tight",
  subtitle: "text-sm text-slate-500",
  card: "rounded-3xl border border-slate-200/80 bg-white/90 shadow-[0_1px_3px_rgba(15,23,42,0.04),0_12px_32px_-12px_rgba(15,23,42,0.12)] backdrop-blur",
  cardBody: "p-6 md:p-8 space-y-5",
  cardTitle: "text-xl font-bold tracking-tight",
  cardDesc: "text-sm text-slate-500",
  hint: "text-sm text-slate-500",
  divider: "h-px bg-slate-100",
  alertError:
    "whitespace-pre-line rounded-2xl border border-red-200 bg-red-50 p-4 text-sm text-red-700",
  btnRow: "flex flex-col gap-2.5",
  btnPrimary:
    "inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-slate-900 px-5 py-3.5 text-sm font-semibold text-white shadow-sm transition-all duration-200 hover:bg-slate-800 hover:shadow-md active:scale-[0.985] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  btnSecondary:
    "inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-slate-100 px-5 py-3.5 text-sm font-semibold text-slate-900 transition-all duration-200 hover:bg-slate-200 active:scale-[0.985] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  btnGhost:
    "inline-flex items-center justify-center gap-1.5 rounded-2xl px-4 py-3 text-sm font-semibold text-slate-600 transition-all duration-200 hover:bg-slate-100 active:scale-[0.985] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  btnNext:
    "inline-flex flex-1 items-center justify-center gap-1.5 rounded-2xl bg-slate-900 px-5 py-3 text-sm font-semibold text-white shadow-sm transition-all duration-200 hover:bg-slate-800 hover:shadow-md active:scale-[0.985] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  btnDanger:
    "inline-flex items-center justify-center gap-1.5 rounded-2xl px-4 py-3 text-sm font-semibold text-red-600 transition-all duration-200 hover:bg-red-50 active:scale-[0.985] disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100",
  fileWrap:
    "rounded-2xl border-2 border-dashed border-slate-200 bg-slate-50/60 p-5 transition-colors duration-200 hover:border-slate-300",
  fileLabel: "text-sm font-semibold text-slate-700",
  fileInput:
    "mt-3 block w-full cursor-pointer rounded-xl border border-slate-200 bg-white text-sm file:mr-3 file:rounded-lg file:border-0 file:bg-slate-900 file:px-4 file:py-2.5 file:text-sm file:font-semibold file:text-white file:transition-colors hover:file:bg-slate-800",
  statsGrid: "grid grid-cols-2 gap-3 text-sm",
  stat: "rounded-2xl border border-slate-100 bg-slate-50/70 p-4",
  statKey: "text-xs font-medium text-slate-500",
  statVal: "mt-1 text-lg font-bold text-slate-900",
  tableWrap: "max-h-56 overflow-auto rounded-2xl border border-slate-200",
  table: "w-full text-sm",
  thead: "sticky top-0 bg-slate-50",
  th: "border-b border-slate-200 p-2.5 text-left font-semibold text-slate-700",
  td: "border-b border-slate-100 p-2.5 text-slate-800",
  badge:
    "inline-flex items-center rounded-full bg-slate-100 px-2.5 py-0.5 text-xs font-semibold text-slate-600",
};

export default function HomePage() {
  const [step, setStep] = useState<TStep>(1);
  // ✅ 2단계: 산출물을 최소 1개 다운로드했는지 (다음 단계 활성화 조건)
  const [step2Done, setStep2Done] = useState<boolean>(false);
  const [loading, setLoading] = useState<{ on: boolean; text: string }>({
    on: false,
    text: "",
  });
  const [error, setError] = useState<string>("");

  // ✅ Step 4 결과 표시용
  const [localResult, setLocalResult] = useState<{
    unmatched: Array<{ customerOrderNo: string; tracking: string }>;
    duplicates: Array<{ key: string; count: number }>;
    matchedCount: number;
    totalReplyCount: number;
  } | null>(null);

  // persisted job (Dexie)
  const [job, setJob] = useState<TJobState | null>(null);

  // ✅ file input reset을 위해 ref 사용 (리셋 시 파일명/선택값 완전 제거)
  const originalInputRef = useRef<HTMLInputElement | null>(null);
  const replyInputRef = useRef<HTMLInputElement | null>(null);

  // derived
  const aggregateRows: TAggregateRow[] = useMemo(() => {
    if (!job?.originalRows) return [];
    // job.originalRows는 TJobState 타입 상 unknown row일 수 있으니 여기서만 캐스팅
    return buildAggregateRows(job.originalRows as TRow[]);
  }, [job?.originalRows]);

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
      const saved = await loadJob();
      if (saved) setJob(saved);
    })();
  }, []);

  const setBusy = (on: boolean, text = "") => setLoading({ on, text });

  // ✅ 각 단계 완료 여부 (= 다음 단계로 넘어갈 수 있는 조건)
  const step1Done = !!job?.originalRows?.length;
  const step3Done = !!localResult;

  // 특정 단계로 이동 가능한지 (진행 바/이전 버튼 용)
  const canGoTo = (target: TStep): boolean => {
    if (target <= step) return true; // 뒤로 가기는 항상 허용
    if (target === 2) return step1Done;
    if (target === 3) return step1Done && step2Done;
    if (target === 4) return step1Done && step2Done && step3Done;
    return true;
  };

  const goNext = () => setStep((s) => (s < 4 ? ((s + 1) as TStep) : s));
  const goPrev = () => setStep((s) => (s > 1 ? ((s - 1) as TStep) : s));

  // 현재 단계에서 "다음" 버튼 활성화 조건
  const nextEnabled =
    (step === 1 && step1Done) ||
    (step === 2 && step2Done) ||
    (step === 3 && step3Done);

  /**
   * ✅ 전체 리셋
   * - DB 초기화
   * - React state 초기화
   * - file input value 강제 초기화 (파일명 남는 문제 해결)
   */
  const onReset = async () => {
    if (loading.on) return;

    await clearJob();

    setJob(null);
    setLocalResult(null);
    setStep2Done(false);
    setStep(1);
    setError("");

    // ✅ input에 남아있는 "선택된 파일명" 제거
    if (originalInputRef.current) originalInputRef.current.value = "";
    if (replyInputRef.current) replyInputRef.current.value = "";
  };

  const onUploadOriginal = async (file: File | null) => {
    if (!file) return;

    try {
      setError("");
      setBusy(true, "원본 엑셀 읽는 중...");

      const { headers, rows } = await readExcelFile(file);

      const next: TJobState = {
        createdAt: new Date().toISOString(),
        originalFileName: file.name,
        originalHeaders: headers,
        originalRows: rows,
        uploadedReplyFiles: [],
      };

      await saveJob(next);
      setJob(next);

      // ✅ 원본 바뀌면 이전 결과는 의미 없으므로 초기화
      setLocalResult(null);
      setStep2Done(false);
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "원본 엑셀 처리 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onDownloadAggregate = async () => {
    try {
      setError("");
      setBusy(true, "품목별 집계 엑셀 생성 중...");

      await downloadAggregateExcel(aggregateRows);

      setStep2Done(true);
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "집계 엑셀 생성 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
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

      await downloadCjUploadsSingle(groups);

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
    if (!job) return;
    if (!files || files.length === 0) return;

    try {
      setError("");

      /**
       * ✅ 중복 파일(동일 fingerprint) 업로드 제외
       */
      const existing = job.uploadedReplyFiles ?? [];
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
      const { map, orderFileMap } = await readCjReplyFiles(accepted);

      /**
       * ✅ 서로 다른 파일에 동일 고객주문번호가 있으면 전체 업로드 차단
       */
      const duplicatedOrders = Array.from(orderFileMap.entries())
        .filter(([, fileSet]) => fileSet.size >= 2)
        .map(([orderNo, fileSet]) => ({
          orderNo,
          files: Array.from(fileSet),
        }));

      if (duplicatedOrders.length > 0) {
        const messageLines = duplicatedOrders
          .slice(0, 5)
          .map((d) => `- ${d.orderNo} : ${d.files.join(", ")}`);

        setError(
          `CJ 회신 파일 오류: 서로 다른 파일에 같은 고객주문번호가 있습니다.\n\n` +
            messageLines.join("\n"),
        );

        // ✅ 업로드를 막되, 파일선택 상태는 유지하되 로딩 상태는 반드시 해제
        setBusy(false);
        return;
      }

      setBusy(true, "운송장번호 매핑 중...");
      const { updatedHeaders, updatedRows, unmatched, duplicates } =
        applyTracking(job.originalHeaders, job.originalRows as TRow[], map);

      const next: TJobState = {
        ...job,
        originalHeaders: updatedHeaders,
        originalRows: updatedRows,
        uploadedReplyFiles: [...existing, ...newFingerprints],
      };

      await saveJob(next);
      setJob(next);

      const unmatchedReplyKeys = new Set(
        unmatched
          .filter((u) => String(u.tracking ?? "").trim() !== "")
          .map((u) => u.customerOrderNo),
      );

      setLocalResult({
        unmatched,
        duplicates,
        matchedCount: map.size - unmatchedReplyKeys.size,
        totalReplyCount: map.size,
      });

      setStep(4);
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "회신 처리 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onDownloadFinal = async () => {
    if (!job) return;

    try {
      setError("");
      setBusy(true, "최종 원본 엑셀 생성 중...");

      await downloadOriginalWithTracking(
        job.originalHeaders,
        job.originalRows as TRow[],
        job.originalFileName,
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

  return (
    <div className={ui.page}>
      <main className={ui.container}>
        {/* Header */}
        <header className={ui.headerWrap}>
          <h1 className={ui.title}>한섬누리 출고 엑셀 도구</h1>
          <p className={ui.subtitle}>
            원본 업로드부터 운송장 반영까지, 한 단계씩 진행하세요
          </p>
        </header>

        {/* Progress bar (버튼으로만 이동, 완료 단계는 되돌아가기 허용) */}
        <nav className="space-y-3">
          <div className="flex items-center">
            {STEP_META.map((s, i) => {
              const isActive = step === s.n;
              const isDone = step > s.n;
              const reachable = canGoTo(s.n);
              return (
                <div
                  key={s.n}
                  className="flex flex-1 items-center last:flex-none"
                >
                  <button
                    type="button"
                    disabled={!reachable || loading.on}
                    onClick={() => reachable && setStep(s.n)}
                    className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-full text-sm font-bold transition-all duration-300 ${
                      isActive
                        ? "bg-slate-900 text-white shadow-md ring-4 ring-slate-900/10"
                        : isDone
                          ? "bg-slate-900 text-white"
                          : "bg-slate-100 text-slate-400"
                    } ${reachable && !isActive ? "cursor-pointer hover:scale-105" : "cursor-default"} disabled:cursor-not-allowed`}
                    title={s.label}
                  >
                    {isDone ? "✓" : s.n}
                  </button>
                  {i < STEP_META.length - 1 && (
                    <div className="mx-1.5 h-0.5 flex-1 overflow-hidden rounded-full bg-slate-100">
                      <div
                        className="h-full rounded-full bg-slate-900 transition-all duration-500"
                        style={{ width: step > s.n ? "100%" : "0%" }}
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
          <div className="flex items-baseline justify-between px-0.5">
            <span className="text-sm font-semibold text-slate-900">
              {step}. {currentMeta.label}
            </span>
            <span className="text-xs text-slate-400">{step} / 4</span>
          </div>
        </nav>

        {/* Error */}
        {error && <div className={ui.alertError}>{error}</div>}

        {/* Loading Overlay */}
        {loading.on && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/30 backdrop-blur-[2px]">
            <div className="flex items-center gap-3 rounded-2xl bg-white p-6 shadow-lg">
              <div className="h-5 w-5 animate-spin rounded-full border-2 border-slate-200 border-t-slate-900" />
              <div className="text-sm font-semibold text-slate-900">
                {loading.text}
              </div>
            </div>
          </div>
        )}

        {/* Step card (한 번에 한 단계만) */}
        <section key={step} className={`${ui.card} step-anim`}>
          <div className={ui.cardBody}>
            <div className="space-y-1">
              <h2 className={ui.cardTitle}>{currentMeta.label}</h2>
              <p className={ui.cardDesc}>{currentMeta.desc}</p>
            </div>

            {/* ===== Step 1 ===== */}
            {step === 1 && (
              <div className="space-y-4">
                <div className={ui.fileWrap}>
                  <div className={ui.fileLabel}>원본 엑셀(.xlsx)</div>
                  <input
                    ref={originalInputRef}
                    type="file"
                    accept=".xlsx"
                    disabled={loading.on}
                    className={ui.fileInput}
                    onChange={(e) =>
                      onUploadOriginal(e.target.files?.[0] ?? null)
                    }
                  />

                  {job?.originalFileName ? (
                    <div className="mt-3 flex items-center gap-1.5 text-sm text-slate-700">
                      <span className="text-emerald-500">✓</span>
                      <span className="font-semibold text-slate-900">
                        {job.originalFileName}
                      </span>
                    </div>
                  ) : (
                    <div className="mt-3 text-sm text-slate-400">
                      업로드하면 다음 단계로 넘어갈 수 있어요.
                    </div>
                  )}
                </div>

                {job?.originalFileName && (
                  <div className={ui.stat}>
                    <div className={ui.statKey}>총 수량</div>
                    <div className={ui.statVal}>
                      {totalQuantity.toLocaleString()}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ===== Step 2 ===== */}
            {step === 2 && (
              <div className="space-y-4">
                <div className={ui.btnRow}>
                  <button
                    className={ui.btnPrimary}
                    disabled={loading.on}
                    onClick={onDownloadAggregate}
                  >
                    📊 품목별 집계 엑셀 다운로드
                  </button>

                  <button
                    className={ui.btnPrimary}
                    disabled={loading.on}
                    onClick={onDownloadCjZip}
                  >
                    🗂️ CJ 업로드용 품목별 ZIP 다운로드
                  </button>

                  <button
                    className={ui.btnPrimary}
                    disabled={loading.on}
                    onClick={onDownloadCjSingle}
                  >
                    📄 CJ 업로드용 통합 엑셀 다운로드
                  </button>
                </div>

                <div className={ui.divider} />

                <div className="flex items-center justify-between text-sm">
                  <span className="text-slate-500">집계 건수</span>
                  <span className="font-semibold text-slate-900">
                    {aggregateRows.length}
                  </span>
                </div>

                <p className={ui.hint}>
                  최소 1개를 다운로드하면 다음 단계로 넘어갈 수 있어요. 집계는
                  “상품약어 → 상품명 → 품목명” 순으로 분류해 합산합니다.
                </p>

                {step2Done && (
                  <div className="flex items-center gap-1.5 text-sm font-medium text-emerald-600">
                    <span>✓</span> 산출물을 받았어요. 다음으로 진행하세요.
                  </div>
                )}
              </div>
            )}

            {/* ===== Step 3 ===== */}
            {step === 3 && (
              <div className="space-y-4">
                <div className={ui.fileWrap}>
                  <div className={ui.fileLabel}>
                    CJ 회신 엑셀(.xlsx) — 여러 개 선택 가능
                  </div>
                  <input
                    ref={replyInputRef}
                    type="file"
                    accept=".xlsx"
                    multiple
                    disabled={loading.on}
                    className={ui.fileInput}
                    onChange={(e) => onUploadReplies(e.target.files)}
                  />
                  <p className="mt-3 text-sm text-slate-500">
                    • 같은 파일을 다시 올리면 경고 후 제외돼요.
                    <br />• 서로 다른 파일에 동일 고객주문번호가 있으면{" "}
                    <b>전체 업로드가 차단</b>됩니다.
                  </p>
                </div>

                {step3Done && (
                  <div className="flex items-center gap-1.5 text-sm font-medium text-emerald-600">
                    <span>✓</span> 회신이 반영됐어요. 다음으로 진행하세요.
                  </div>
                )}
              </div>
            )}

            {/* ===== Step 4 ===== */}
            {step === 4 && (
              <div className="space-y-4">
                {localResult ? (
              <div className={ui.statsGrid}>
                <div className={ui.stat}>
                  <div className={ui.statKey}>총 수량</div>
                  <div className={ui.statVal}>{totalQuantity}</div>
                </div>
                <div className={ui.stat}>
                  <div className={ui.statKey}>회신 키 수</div>
                  <div className={ui.statVal}>
                    {localResult.totalReplyCount}
                  </div>
                </div>
                <div className={ui.stat}>
                  <div className={ui.statKey}>매핑 성공(추정)</div>
                  <div className={ui.statVal}>{localResult.matchedCount}</div>
                </div>
                <div className={ui.stat}>
                  <div className={ui.statKey}>미매칭</div>
                  <div className={ui.statVal}>
                    {localResult.unmatched.length}
                  </div>
                </div>
                <div className={ui.stat}>
                  <div className={ui.statKey}>원본 중복키</div>
                  <div className={ui.statVal}>
                    {localResult.duplicates.length}
                  </div>
                </div>
              </div>
            ) : (
              <div className="rounded-xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-700">
                회신 업로드 후 결과가 표시됩니다.
              </div>
            )}

            <div className={ui.btnRow}>
              <button
                className={ui.btnPrimary}
                disabled={!job || loading.on || !step3Done}
                onClick={onDownloadFinal}
              >
                ⬇️ 최종 원본 다운로드
              </button>

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
                <div className="text-sm font-semibold text-slate-900">
                  미매칭 리스트 (상위 20개)
                </div>
                <div className={ui.tableWrap}>
                  <table className={ui.table}>
                    <thead className={ui.thead}>
                      <tr>
                        <th className={ui.th}>고객주문번호</th>
                        <th className={ui.th}>운송장번호</th>
                      </tr>
                    </thead>
                    <tbody>
                      {localResult.unmatched.slice(0, 20).map((u) => (
                        <tr key={`${u.customerOrderNo}-${u.tracking}`}>
                          <td className={ui.td}>{u.customerOrderNo}</td>
                          <td className={ui.td}>{u.tracking}</td>
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
                <div className="text-sm font-semibold text-slate-900">
                  원본 중복키 (상위 20개)
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
              </div>
            )}

            <div className={ui.divider} />

            {/* ===== Step navigation ===== */}
            <div className="flex items-center gap-2">
              {step > 1 ? (
                <button
                  type="button"
                  className={ui.btnGhost}
                  disabled={loading.on}
                  onClick={goPrev}
                >
                  ← 이전
                </button>
              ) : (
                <span />
              )}

              <div className="flex-1" />

              {step < 4 ? (
                <button
                  type="button"
                  className={ui.btnNext}
                  disabled={!nextEnabled || loading.on}
                  onClick={goNext}
                  title={!nextEnabled ? "이 단계를 먼저 완료해주세요" : ""}
                >
                  다음 →
                </button>
              ) : (
                <button
                  type="button"
                  className={ui.btnDanger}
                  disabled={loading.on}
                  onClick={() => {
                    if (confirm("모든 작업을 초기화하고 처음부터 시작할까요?"))
                      onReset();
                  }}
                >
                  처음부터 다시
                </button>
              )}
            </div>
          </div>
        </section>

        {/* Footer hint */}
        <footer className="text-center text-xs text-slate-400">
          로컬에서만 처리되며(브라우저), 업로드한 파일은 서버로 전송되지
          않습니다.
        </footer>
      </main>
    </div>
  );
}
