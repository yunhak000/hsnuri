/* eslint-disable @typescript-eslint/no-require-imports */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const ts = require("typescript");
const ExcelJS = require("exceljs");

// Load the project's TypeScript without adding a test runtime dependency.
let downloads = [];
let templateStatus = 200;
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
  const templateFetch = async (url) => {
    assert.equal(url, "/templates/cj-integrated-order.xlsx");
    return new Response(
      templateStatus === 200 ? fs.readFileSync(path.resolve(__dirname, "../public/templates/cj-integrated-order.xlsx")) : null,
      { status: templateStatus },
    );
  };
  new Function("require", "module", "exports", "fetch", code)(localRequire, loaded, loaded.exports, templateFetch);
  return loaded.exports;
}

const { toText } = load("src/lib/utils/normalize.ts");
const { formatKstDateTime } = load("src/lib/utils/datetime.ts");

// Excel omits explicit false flags on export; omitted and false have the same layout.
const comparableStyle = (style) => JSON.parse(JSON.stringify(style, (key, value) =>
  ["wrapText", "shrinkToFit"].includes(key) && value === false ? undefined : value,
));

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

test("integrated orders use the supplied template, retain identifiers and support CJ replies", async () => {
  const { buildCjGroupedRows, downloadCjUploadsSingle, downloadCjUploadsZip } = load("src/lib/excel/writeCJUploads.ts");
  const originalRows = Array.from({ length: 4 }, (_, index) => ({
    상품주문번호: `000${index + 1}`,
    품목명: index % 2 ? "천혜향" : "감귤",
    옵션: "선물용",
    수량: "2",
    주문자: "테스트 발신인",
    주문자전화번호1: "010-0000-0001",
    주문자우편번호: "00123",
    주문자주소: "테스트 발신 주소",
    수취인명: `테스트 수취인 ${index + 1}`,
    수취인전화번호1: "010-0000-0002",
    수취인우편번호: "04567",
    수취인주소: "테스트 수취 주소",
    배송메시지: "문 앞에 놓아주세요",
    기타: "이 문구는 통합 양식에서 제외",
  }));
  const groups = buildCjGroupedRows(Object.keys(originalRows[0]), originalRows);
  downloads = [];
  await downloadCjUploadsSingle(groups);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await downloads[0].blob.arrayBuffer());
  const source = new ExcelJS.Workbook();
  await source.xlsx.readFile(path.resolve(__dirname, "../public/templates/cj-integrated-order.xlsx"));
  const ws = wb.worksheets[0];
  const template = source.worksheets[0];
  assert.equal(wb.worksheets.length, 1);
  assert.equal(ws.name, "개인");
  assert.deepEqual(ws.getRow(1).values, template.getRow(1).values);
  assert.deepEqual(ws.model.merges, template.model.merges);
  assert.deepEqual(ws.views, template.views);
  assert.equal(ws.rowCount, 5);
  for (let col = 1; col <= template.columnCount; col++) {
    assert.equal(ws.getColumn(col).width, template.getColumn(col).width);
    assert.deepEqual(comparableStyle(ws.getRow(1).getCell(col).style), comparableStyle(template.getRow(1).getCell(col).style));
  }
  const ids = new Set();
  for (let rowNumber = 2; rowNumber <= 5; rowNumber++) {
    const row = ws.getRow(rowNumber);
    assert.ok(!row.getCell(1).value);
    assert.equal(row.getCell(2).value, "테스트 발신인");
    assert.equal(row.getCell(3).value, "010-0000-0001");
    assert.equal(row.getCell(4).value, "00123");
    assert.equal(row.getCell(5).value, "테스트 발신 주소");
    assert.match(row.getCell(6).value, /^(감귤|천혜향) \/ 선물용$/);
    assert.equal(row.getCell(7).value, 2);
    assert.match(row.getCell(8).value, /^테스트 수취인 [1-4]$/);
    assert.equal(row.getCell(9).value, "010-0000-0002");
    assert.equal(row.getCell(10).value, "04567");
    assert.equal(row.getCell(11).value, "테스트 수취 주소");
    assert.equal(row.getCell(12).value, "문 앞에 놓아주세요");
    ids.add(row.getCell(13).value);
    assert.ok(!row.getCell(14).value);
    for (let col = 1; col <= 14; col++) {
      assert.deepEqual(comparableStyle(row.getCell(col).style), comparableStyle(template.getRow(2).getCell(col).style));
    }
    row.getCell(14).value = `000111${rowNumber}`;
  }
  assert.deepEqual(ids, new Set(["0001", "0002", "0003", "0004"]));
  assert.ok(!JSON.stringify(ws.getSheetValues()).includes("신종석"));
  const { readCjReplyFiles } = load("src/lib/excel/readCJReply.ts");
  const reply = await readCjReplyFiles([new File([await wb.xlsx.writeBuffer()], "회신.xlsx")]);
  assert.equal(reply.map.size, 4);
  assert.equal(reply.skipped.length, 0);

  await downloadCjUploadsZip(groups);
  const zip = await require("jszip").loadAsync(await downloads[1].blob.arrayBuffer());
  const oldWorkbook = new ExcelJS.Workbook();
  await oldWorkbook.xlsx.load(await zip.file("감귤.xlsx").async("nodebuffer"));
  assert.equal(oldWorkbook.worksheets[0].getCell("A1").value, "기타");
  assert.equal(oldWorkbook.worksheets[0].getCell("A2").value, "검사원: 신종석 2-006-201");
});

test("integrated orders report a template loading error without downloading a file", async () => {
  const { downloadCjUploadsSingle } = load("src/lib/excel/writeCJUploads.ts");
  downloads = [];
  templateStatus = 404;
  try {
    await assert.rejects(downloadCjUploadsSingle(new Map()), /양식을 불러오지 못했습니다/);
    assert.equal(downloads.length, 0);
  } finally {
    templateStatus = 200;
  }
});
