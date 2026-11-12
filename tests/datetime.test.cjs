/* eslint-disable @typescript-eslint/no-require-imports */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const ts = require("typescript");
const ExcelJS = require("exceljs");

// Load the project's TypeScript without adding a test runtime dependency.
let downloads = [];
const cache = new Map();
function load(relativePath) {
  const filename = path.resolve(__dirname, "..", relativePath);
  if (cache.has(filename)) return cache.get(filename).exports;
  const loaded = { exports: {} };
  cache.set(filename, loaded);
  const code = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { target: ts.ScriptTarget.ES2017, module: ts.ModuleKind.CommonJS, esModuleInterop: true },
  }).outputText;
  const localRequire = (id) => {
    if (id === "file-saver") return { saveAs: (blob, name) => downloads.push({ blob, name }) };
    if (id.startsWith("@/")) return load(`src/${id.slice(2)}.ts`);
    if (id.startsWith(".")) return load(path.relative(path.resolve(__dirname, ".."), path.resolve(path.dirname(filename), `${id}.ts`)));
    return require(id);
  };
  new Function("require", "module", "exports", code)(localRequire, loaded, loaded.exports);
  return loaded.exports;
}

const { toText } = load("src/lib/utils/normalize.ts");
const { formatKstDateTime } = load("src/lib/utils/datetime.ts");
const { getJejuDeliverySummary } = load("src/lib/utils/jejuDelivery.ts");

test("Jeju delivery warning sums boxes once per matching address", () => {
  assert.deepEqual(getJejuDeliverySummary([
    { 받는분주소: "제주도 제주시", 박스수량: 2 },
    { 받는분주소: "제주특별자치도 서귀포시", 박스수량: "3" },
    { 받는분주소: "제주도 제주특별자치도 제주시", 박스수량: 4 },
    { "받는분 주소": "제주도 제주시", "박스 수량": " 1,000 " },
    { 받는분주소: "서울특별시", 박스수량: 100 },
    { 보내는분주소: "제주도", 박스수량: 100 },
    { 받는분주소: "제주도", 박스수량: "수량 미입력", 수량: 100 },
    { 받는분주소: "제주도", 박스수량: null },
  ]), { hasJejuDelivery: true, totalBox: 1009 });
  assert.deepEqual(getJejuDeliverySummary([]), { hasJejuDelivery: false, totalBox: 0 });
  assert.deepEqual(getJejuDeliverySummary([{ 받는분주소: "부산광역시", 박스수량: 2 }]), { hasJejuDelivery: false, totalBox: 0 });
});

test("spreadsheet times stay unchanged and explicit timezones convert to Korea", () => {
  const originalTimezone = process.env.TZ;
  try {
    for (const timezone of ["Asia/Seoul", "UTC", "America/Los_Angeles"]) {
      process.env.TZ = timezone;
      assert.equal(toText(new Date("2026-11-04T13:05:09Z")), "2026-11-04 13:05:09");
      assert.equal(toText({ result: new Date("2026-11-04T13:05:09Z") }), "2026-11-04 13:05:09");
      assert.equal(toText("2026-11-04T13:05:09Z"), "2026-11-04 22:05:09");
      assert.equal(toText("2026-11-04T13:05:09+09:00"), "2026-11-04 13:05:09");
      assert.equal(toText("Wed Nov 04 2026 13:05:09 GMT+0900 (Korean Standard Time)"), "2026-11-04 13:05:09");
      assert.equal(toText({ text: "2026-11-04 13:05:09 KST" }), "2026-11-04 13:05:09");
      assert.equal(toText("2026.11.04 13.05.09"), "2026-11-04 13:05:09");
      assert.equal(toText("2026/1/4 3:05"), "2026-01-04 03:05:00");
      assert.equal(toText("2026-11-04"), "2026-11-04");
      assert.equal(toText("2026-02-30 13:05:09"), "2026-02-30 13:05:09");
      assert.equal(toText("20261104130509"), "20261104130509");
      assert.equal(toText("주문 2026.11.04"), "주문 2026.11.04");
      assert.equal(formatKstDateTime(new Date("2026-11-04T15:00:00Z")), "2026-11-05 00:00:00");
      assert.equal(toText(null), "");
      assert.equal(toText(new Date("invalid")), "");
    }
  } finally {
    if (originalTimezone === undefined) delete process.env.TZ;
    else process.env.TZ = originalTimezone;
  }
});

test("actual final and CJ workbooks export normalized timestamps", async () => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet("원본");
  sheet.addRow(["주문일시", "회신일시"]);
  sheet.addRow([new Date("2026-11-04T13:05:09Z"), "Wed Nov 04 2026 13:05:09 GMT+0900 (Korean Standard Time)"]);
  sheet.getCell("A2").numFmt = "yyyy-mm-dd hh:mm:ss";
  const { readExcelFile } = load("src/lib/excel/read.ts");
  const original = await readExcelFile(new File([await source.xlsx.writeBuffer()], "원본.xlsx"));
  const { downloadOriginalWithTracking } = load("src/lib/excel/applyTrackingToOriginal.ts");
  const { downloadCjUploadsSingle, downloadCjUploadsZip } = load("src/lib/excel/writeCJUploads.ts");
  downloads = [];
  await downloadOriginalWithTracking(original.headers, original.rows, "원본.xlsx");
  const final = new ExcelJS.Workbook();
  await final.xlsx.load(await downloads[0].blob.arrayBuffer());
  assert.equal(final.worksheets[0].getCell("A2").value, "2026-11-04 13:05:09");
  assert.equal(final.worksheets[0].getCell("B2").value, "2026-11-04 13:05:09");

  const groups = new Map([["테스트", [{ 품목명: original.rows[0]["주문일시"] }]]]);
  await downloadCjUploadsSingle(groups);
  await downloadCjUploadsZip(groups);
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(await downloads[2].blob.arrayBuffer());
  const buffers = [await downloads[1].blob.arrayBuffer(), await zip.file("테스트.xlsx").async("nodebuffer")];
  for (const buffer of buffers) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.worksheets[0];
    const col = ws.getRow(1).values.indexOf("품목명");
    assert.equal(ws.getRow(2).getCell(col).value, "2026-11-04 13:05:09");
  }
  for (const { name } of downloads.slice(1)) assert.match(name, /^\d{4}-\d{2}-\d{2}_/);
});

test("CJ ZIP and combined workbook use original product names with item name fallback", async () => {
  const { buildCjGroupedRows, downloadCjUploadsSingle, downloadCjUploadsZip } = load("src/lib/excel/writeCJUploads.ts");
  const rows = [
    { 상품주문번호: "1", 상품약어: "분류A", 품목명: "품목A", 원본엑셀상품명: "원본 상품A", 옵션: "옵션A" },
    { 상품주문번호: "2", 상품약어: "분류A", 품목명: "품목B", 원본엑셀상품명: "", 상품명: "다른 상품B", 상품옵션: "옵션B" },
    { 상품주문번호: "3", 상품약어: "분류B", 품목명: "품목C", 원본엑셀상품명: "  \n " },
    { 상품주문번호: "4", 상품약어: "분류B", 품목명: "품목D" },
  ];
  const groups = buildCjGroupedRows(Object.keys(rows[0]), rows);
  downloads = [];
  await downloadCjUploadsSingle(groups);
  await downloadCjUploadsZip(groups);

  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(await downloads[1].blob.arrayBuffer());
  assert.deepEqual(Object.keys(zip.files), ["분류A.xlsx", "분류B.xlsx"]);
  const cases = [
    [await downloads[0].blob.arrayBuffer(), ["원본 상품A", "품목B", "품목C", "품목D"], ["품목A", "품목B", "품목C", "품목D"]],
    [await zip.file("분류A.xlsx").async("nodebuffer"), ["원본 상품A", "품목B"], ["품목A", "품목B"]],
    [await zip.file("분류B.xlsx").async("nodebuffer"), ["품목C", "품목D"], ["품목C", "품목D"]],
  ];
  for (const [buffer, expected, expectedAdminNames] of cases) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.worksheets[0];
    const col = ws.getRow(1).values.indexOf("품목명");
    const actual = Array.from({ length: ws.rowCount - 1 }, (_, index) => ws.getRow(index + 2).getCell(col).value);
    assert.deepEqual(actual, expected);
    assert.equal(ws.getCell("P1").value, "어드민품목명");
    const adminNames = Array.from({ length: ws.rowCount - 1 }, (_, index) => ws.getCell(`P${index + 2}`).value);
    assert.deepEqual(adminNames, expectedAdminNames);
  }
});
