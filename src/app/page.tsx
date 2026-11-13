"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import Image from "next/image";
import { readExcelFile } from "@/lib/excel/read";
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
      const saved = await loadJob();
      if (saved) setJob(saved);
    })();
  }, []);

  const setBusy = (on: boolean, text = "") => setLoading({ on, text });

  // ✅ 각 단계 완료 여부 (= 다음 단계로 넘어갈 수 있는 조건)
  const step1Done = !!job?.originalRows?.length;
  const step3Done = !!localResult;

  const goNext = () => setStep((s) => (s < 4 ? ((s + 1) as TStep) : s));
  const goPrev = () => setStep((s) => (s > 1 ? ((s - 1) as TStep) : s));

  // 현재 단계에서 "다음" 버튼 활성화 조건
  const nextEnabled =
    (step === 1 && step1Done) ||
    (step === 2 && (aggregateRows.length > 0 || step2Done)) ||
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
      await clearJob();

      setJob(null);
      setLocalResult(null);
      setStep2Done(false);
      setAggregateCopied(false);
      setStep(1);
      setError("");

      // ✅ input에 남아있는 "선택된 파일명" 제거
      if (originalInputRef.current) originalInputRef.current.value = "";
      if (replyInputRef.current) replyInputRef.current.value = "";
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "작업 초기화 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onUploadOriginal = async (file: File | null) => {
    if (!file) return;

    try {
      setError("");
      setBusy(true, "원본 엑셀 읽는 중...");

      const { headers, rows, originalWorkbookBuffer } = await readExcelFile(file);

      const next: TJobState = {
        createdAt: new Date().toISOString(),
        originalFileName: file.name,
        originalHeaders: headers,
        originalRows: rows,
        originalWorkbookBuffer,
        uploadedReplyFiles: [],
      };

      await saveJob(next);
      setJob(next);

      // ✅ 원본 바뀌면 이전 결과는 의미 없으므로 초기화
      setLocalResult(null);
      setStep2Done(false);
      setAggregateCopied(false);
    } catch (e: unknown) {
      setError(getErrorMessage(e) || "원본 엑셀 처리 중 오류가 발생했습니다.");
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
        job.originalWorkbookBuffer,
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
      <main className={ui.container}>
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
            disabled={loading.on}
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
              <label className={ui.dropzone}>
                <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-primary-soft text-3xl">
                  📄
                </div>
                <div className="text-[15px] font-bold text-[#191f28]">
                  원본 엑셀 올리기
                </div>
                <div className="text-[13px] text-[#8b95a1]">
                  .xlsx 파일을 선택하세요
                </div>
                <input
                  ref={originalInputRef}
                  type="file"
                  accept=".xlsx"
                  disabled={loading.on}
                  className="hidden"
                  onChange={(e) =>
                    onUploadOriginal(e.target.files?.[0] ?? null)
                  }
                />
              </label>

              {job?.originalFileName && (
                <div className="pop-in space-y-3">
                  <div className="flex items-center gap-2.5 rounded-2xl bg-white p-4">
                    <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-primary text-[13px] text-white">
                      ✓
                    </span>
                    <span className="flex-1 truncate text-[15px] font-semibold text-[#191f28]">
                      {job.originalFileName}
                    </span>
                  </div>
                  <div className={ui.stat}>
                    <div className={ui.statKey}>총 수량</div>
                    <div className={ui.statVal}>
                      {totalQuantity.toLocaleString()}개
                    </div>
                  </div>
                </div>
              )}
            </>
          )}

          {/* Step 2 */}
          {step === 2 && (
            <>
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
                  disabled={loading.on}
                  onClick={onDownloadCjZip}
                >
                  <span className={ui.listIcon}>🗂️</span>
                  <span className="flex-1">
                    <span className={ui.listTitle}>CJ 업로드용 품목별 ZIP</span>
                    <span className={ui.listDesc}>품목별로 파일이 나뉨</span>
                  </span>
                  <span className="text-[#c4cad1]">›</span>
                </button>

                <button
                  className={ui.listBtn}
                  disabled={loading.on}
                  onClick={onDownloadCjSingle}
                >
                  <span className={ui.listIcon}>📄</span>
                  <span className="flex-1">
                    <span className={ui.listTitle}>CJ 업로드용 통합 엑셀</span>
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
                  .xlsx 여러 개를 한 번에 선택할 수 있어요
                </div>
                <input
                  ref={replyInputRef}
                  type="file"
                  accept=".xlsx"
                  multiple
                  disabled={loading.on}
                  className="hidden"
                  onChange={(e) => onUploadReplies(e.target.files)}
                />
              </label>

              <div className="rounded-2xl bg-white p-4">
                <p className={ui.hint}>
                  • 같은 파일을 다시 올리면 경고 후 제외돼요.
                  <br />• 서로 다른 파일에 같은 고객주문번호가 있으면{" "}
                  <b className="text-[#4e5968]">전체 업로드가 차단</b>됩니다.
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
                <button
                  className={ui.btnPrimary}
                  disabled={!job || loading.on || !step3Done}
                  onClick={onDownloadFinal}
                >
                  최종 원본 다운로드
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
                  <div className="px-1 text-[14px] font-bold text-[#191f28]">
                    미매칭 리스트{" "}
                    <span className="text-[#8b95a1]">(상위 20개)</span>
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
                disabled={loading.on}
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
