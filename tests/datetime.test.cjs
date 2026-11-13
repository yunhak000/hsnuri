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

test("aggregate groups Tyvek first and sorts each cultivation group by weight", () => {
  const { buildAggregateRows } = load("src/lib/excel/writeAggregate.ts");
  const names = [
    "제주도 서귀포 노지감귤 로얄과(2S~M) 2.5kg",
    "제주도 서귀포 타이벡 감귤 로얄과 (2S~M) 2.5kg",
    "제주도 서귀포 노지감귤 꼬마과(3S) 5kg",
    "제주도 서귀포 노지감귤 로얄과(2S~M) 5kg",
    "제주도 서귀포 노지감귤 로얄과(2S~M) 10kg",
  ];
  const quantities = [4, 5, 1, 7, 2];
  const rows = names.map((품목명, index) => ({ 품목명, 박스수량: quantities[index] }));
  const result = buildAggregateRows(rows);
  assert.deepEqual(result.map((row) => row.itemName), [names[1], names[0], names[2], names[3], names[4]]);
  assert.deepEqual(result.map((row) => row.totalBox), [5, 4, 1, 7, 2]);
  assert.equal(result.reduce((sum, row) => sum + row.totalBox, 0), 19);
  assert.deepEqual(result.map((row) => row.kg), [2.5, 2.5, 5, 5, 10]);

  const extraTyvek = "제주도 서귀포 타이벡 감귤 로얄과 10kg";
  const expanded = buildAggregateRows([{ 품목명: extraTyvek, 박스수량: 3 }, ...rows]);
  assert.deepEqual(expanded.slice(0, 2).map((row) => row.itemName), [names[1], extraTyvek]);
});

test("aggregate preserves source weights for original and Sabangnet orders and Excel export", async () => {
  const { buildAggregateRows, downloadAggregateExcel } = load("src/lib/excel/writeAggregate.ts");
  const weights = [10, 5, 9, 4.5];
  for (const source of ["original", "sabangnet"]) {
    const rows = weights.map((kg) => source === "original"
      ? { 품목명: `노지감귤 ${kg}kg`, 박스수량: 2 }
      : { 상품약어: `노지감귤 ${kg}kg`, 상품명: "판매 상품명", 수량: 2 });
    const result = buildAggregateRows(rows);
    assert.deepEqual(result.map((row) => row.kg), [4.5, 5, 9, 10]);
    const expectedNames = ["노지감귤 4.5kg", "노지감귤 5kg", "노지감귤 9kg", "노지감귤 10kg"];
    assert.deepEqual(result.map((row) => row.itemName), expectedNames);
    assert.equal(result.reduce((sum, row) => sum + row.totalBox, 0), 8);
    downloads = [];
    await downloadAggregateExcel(result);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await downloads[0].blob.arrayBuffer());
    const sheet = workbook.worksheets[0];
    assert.deepEqual(expectedNames.map((_, index) => sheet.getCell(`A${index + 2}`).value), expectedNames);
  }
});

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

test("Jeju delivery uses Sabangnet quantity and preserves box quantity priority", () => {
  assert.deepEqual(getJejuDeliverySummary([
    { 받는분주소: "제주도 제주시", 수량: "2" },
    { 받는분주소: "제주특별자치도 서귀포시", 수량: "3" },
    { 받는분주소: "제주도", 박스수량: 4, 수량: 20 },
    { 받는분주소: "제주도", 박스수량: 0, 수량: 20 },
    { 받는분주소: "제주도", 박스수량: " ", 수량: 1 },
    { 받는분주소: "서울특별시", 수량: 100 },
  ]), { hasJejuDelivery: true, totalBox: 10 });
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

test("Sabangnet ZIP and combined exports map product, admin product and sales channel", async () => {
  const { buildCjGroupedRows, downloadCjUploadsSingle, downloadCjUploadsZip } = load("src/lib/excel/writeCJUploads.ts");
  const rows = [
    { "★쇼핑몰 주문번호★": "S1", "사방넷 주문번호": "A1", 상품약어: "타이벡_5kg", 상품명: "고당도 타이벡 감귤 5KG", 쇼핑몰명: "카카오톡선물하기", 수량: 2, 옵션: "옵션 추가 금지" },
    { "★쇼핑몰 주문번호★": "S2", "사방넷 주문번호": "A2", 상품약어: "황금향_2kg", 상품명: "제주 황금향 2KG", 쇼핑몰명: "스마트스토어", 수량: 1 },
  ];
  const groups = buildCjGroupedRows(Object.keys(rows[0]), rows);
  downloads = [];
  await downloadCjUploadsSingle(groups);
  await downloadCjUploadsZip(groups);
  const zip = await require("jszip").loadAsync(await downloads[1].blob.arrayBuffer());
  const cases = [[await downloads[0].blob.arrayBuffer(), rows]];
  for (const row of rows) cases.push([await zip.file(`${row.상품약어}.xlsx`).async("nodebuffer"), [row]]);
  for (const [buffer, expected] of cases) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.worksheets[0];
    expected.forEach((row, index) => {
      const number = index + 2;
      assert.equal(ws.getCell(`F${number}`).value, row.상품명);
      assert.equal(ws.getCell(`P${number}`).value, row.상품약어);
      assert.equal(ws.getCell(`Q${number}`).value, `사방넷-(${row.쇼핑몰명})`);
      assert.equal(ws.getCell(`G${number}`).value, String(row.수량));
      assert.equal(ws.getCell(`O${number}`).value, row["★쇼핑몰 주문번호★"]);
    });
  }
});

test("final export fills matching orders across all original sheets without duplicating totals", async () => {
  const source = new ExcelJS.Workbook();
  const all = source.addWorksheet("전체 주문");
  const headers = ["상품약어", "상품명", "수량", "★쇼핑몰 주문번호★", "송장번호"];
  all.addRow(headers);
  all.addRow(["타이벡_1.5KG", "타이벡 1.5kg", 2, "S1", ""]);
  all.addRow(["타이벡_4.5KG", "타이벡 4.5kg", 1, "S2", ""]);
  all.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } };
  all.getColumn(1).width = 35;
  all.getCell("F1").value = "송장번호(2)";
  for (const address of ["F1", "F2", "F3"]) {
    all.getCell(address).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF59D" } };
  }
  const small = source.addWorksheet("타이벡 1.5KG");
  small.addRow(headers);
  small.addRow(["타이벡_1.5KG", "타이벡 1.5kg", 2, "S1", ""]);
  small.addRow(["별도 주문", "유지", 1, "UNKNOWN", "기존 송장"]);
  small.addRow(["반복된 주문", "동일 번호도 반영", 2, "S1", ""]);
  const large = source.addWorksheet("타이벡 4.5KG");
  large.addRow(["운송장번호", "★쇼핑몰 주문번호★", "상품명"]);
  large.addRow(["", "S2", "타이벡 4.5kg"]);
  large.addRow(["", "ONLY_ON_DETAIL", "품목 시트에만 있는 주문"]);
  const notes = source.addWorksheet("안내");
  notes.addRow(["메모", "보존할 내용"]);
  notes.getCell("A2").value = { formula: "1+1", result: 2 };
  notes.state = "hidden";
  notes.mergeCells("B2:C2");
  notes.getCell("B2").value = "합쳐진 안내문";

  const { readExcelFile } = load("src/lib/excel/read.ts");
  const original = await readExcelFile(new File([await source.xlsx.writeBuffer()], "사방넷.xlsx"));
  assert.equal(original.rows.length, 2);
  assert.ok(original.originalWorkbookBuffer instanceof ArrayBuffer);
  const { buildAggregateRows } = load("src/lib/excel/writeAggregate.ts");
  assert.equal(buildAggregateRows(original.rows).reduce((sum, row) => sum + row.totalBox, 0), 3);

  const { applyTracking, downloadOriginalWithTracking } = load("src/lib/excel/applyTrackingToOriginal.ts");
  const applied = applyTracking(original.headers, original.rows, new Map([
    ["S2", ["000222"]],
    ["S1", ["000111", "000112"]],
  ]));
  // The buffer also survives the structured clone used when persisting job state.
  const buffer = structuredClone(original.originalWorkbookBuffer);
  downloads = [];
  await downloadOriginalWithTracking(applied.updatedHeaders, applied.updatedRows, "사방넷.xlsx", buffer, {
    S1: ["000111", "000112"], S2: ["000222"], ONLY_ON_DETAIL: ["000333"],
  });
  const final = new ExcelJS.Workbook();
  await final.xlsx.load(await downloads[0].blob.arrayBuffer());
  assert.deepEqual(final.worksheets.map((sheet) => sheet.name), ["전체 주문", "타이벡 1.5KG", "타이벡 4.5KG", "안내"]);
  for (const name of ["전체 주문", "타이벡 1.5KG"]) {
    const sheet = final.getWorksheet(name);
    assert.equal(sheet.getCell("E2").value, "000111");
    assert.equal(sheet.getCell("F1").value, "송장번호(2)");
    assert.equal(sheet.getCell("F2").value, "000112");
    assert.equal(sheet.getCell("F2").fill.fgColor.argb, "FFFFF59D");
    assert.notEqual(sheet.getCell("F1").fill?.fgColor?.argb, "FFFFF59D");
    assert.notEqual(sheet.getCell("F3").fill?.fgColor?.argb, "FFFFF59D");
  }
  assert.equal(final.getWorksheet("전체 주문").getCell("E3").value, "000222");
  assert.equal(final.getWorksheet("타이벡 4.5KG").getCell("A2").value, "000222");
  assert.equal(final.getWorksheet("타이벡 4.5KG").getCell("A3").value, "000333");
  assert.equal(final.getWorksheet("타이벡 1.5KG").getCell("E3").value, "기존 송장");
  assert.equal(final.getWorksheet("타이벡 1.5KG").getCell("E4").value, "000111");
  assert.equal(final.getWorksheet("전체 주문").getColumn(1).width, 35);
  assert.equal(final.getWorksheet("전체 주문").getCell("A1").fill.fgColor.argb, "FF00FF00");
  assert.deepEqual(final.getWorksheet("안내").getCell("A2").value, { formula: "1+1", result: 2 });
  assert.equal(final.getWorksheet("안내").getCell("B1").value, "보존할 내용");
  assert.equal(final.getWorksheet("안내").state, "hidden");
  assert.equal(final.getWorksheet("안내").getCell("C2").master.address, "B2");
  assert.match(downloads[0].name, /_3건\.xlsx$/);
});

test("single-sheet source keeps its name and highlights only populated tracking cells", async () => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet("기존 주문서 이름");
  sheet.addRow(["상품주문번호", "품목명", "운송장번호", "운송장번호(2)"]);
  sheet.addRow(["A", "감귤 5kg", "", ""]);
  sheet.addRow(["B", "감귤 10kg", "", ""]);
  const { readExcelFile } = load("src/lib/excel/read.ts");
  const original = await readExcelFile(new File([await source.xlsx.writeBuffer()], "원본.xlsx"));
  assert.ok(original.originalWorkbookBuffer instanceof ArrayBuffer);
  const { applyTracking, downloadOriginalWithTracking } = load("src/lib/excel/applyTrackingToOriginal.ts");
  const applied = applyTracking(original.headers, original.rows, new Map([["A", ["001", "002"]], ["B", ["003"]]]));
  for (const buffer of [original.originalWorkbookBuffer, undefined]) {
    downloads = [];
    await downloadOriginalWithTracking(applied.updatedHeaders, applied.updatedRows, "원본.xlsx", buffer);
    const final = new ExcelJS.Workbook();
    await final.xlsx.load(await downloads[0].blob.arrayBuffer());
    const ws = final.worksheets[0];
    if (buffer) assert.equal(ws.name, "기존 주문서 이름");
    assert.equal(ws.getCell("D2").value, "002");
    assert.equal(ws.getCell("D2").fill.fgColor.argb, "FFFFF59D");
    assert.notEqual(ws.getCell("D3").fill?.fgColor?.argb, "FFFFF59D");
    assert.notEqual(ws.getCell("D1").fill?.fgColor?.argb, "FFFFF59D");
  }
});

test("tracking export preserves original table, filter, workbook and unrelated sheet parts byte for byte", async () => {
  const source = new ExcelJS.Workbook();
  for (const name of ["전체 주문", "품목별 주문"]) {
    const sheet = source.addWorksheet(name);
    sheet.addTable({
      name: name === "전체 주문" ? "Orders" : "ProductOrders",
      ref: "A1", headerRow: true,
      columns: ["상품명", "★쇼핑몰 주문번호★", "송장번호"].map((name) => ({ name, filterButton: true })),
      rows: [["감귤 5kg", "S1", ""], ["감귤 10kg", "S2", ""]],
    });
  }
  const notes = source.addWorksheet("관련없는 시트");
  notes.addRow(["그대로 보존", { formula: "1+1", result: 2 }]);
  notes.mergeCells("A2:C2");
  const sourceBuffer = await source.xlsx.writeBuffer();
  const { readExcelFile } = load("src/lib/excel/read.ts");
  const original = await readExcelFile(new File([sourceBuffer], "표가있는원본.xlsx"));
  const { applyTracking, downloadOriginalWithTracking } = load("src/lib/excel/applyTrackingToOriginal.ts");
  const applied = applyTracking(original.headers, original.rows, new Map([["S1", ["0001", "0002"]], ["S2", ["0003"]]]));
  downloads = [];
  await downloadOriginalWithTracking(applied.updatedHeaders, applied.updatedRows, "표가있는원본.xlsx", original.originalWorkbookBuffer);
  const JSZip = require("jszip");
  const before = await JSZip.loadAsync(sourceBuffer);
  const after = await JSZip.loadAsync(await downloads[0].blob.arrayBuffer());
  assert.deepEqual(Object.keys(after.files).sort(), Object.keys(before.files).sort());
  const changedParts = new Set(["xl/styles.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]);
  for (const [path, part] of Object.entries(before.files)) {
    if (part.dir || changedParts.has(path)) continue;
    assert.equal(await after.file(path).async("string"), await part.async("string"), path);
  }
  const result = new ExcelJS.Workbook();
  await result.xlsx.load(await downloads[0].blob.arrayBuffer());
  for (const sheet of result.worksheets.slice(0, 2)) {
    assert.equal(sheet.getCell("C2").value, "0001");
    assert.equal(sheet.getCell("D2").value, "0002");
    assert.notEqual(sheet.getCell("D3").fill?.fgColor?.argb, "FFFFF59D");
  }
});

test("CJ ZIP and combined workbook use original product names with item name fallback", async () => {
  const { buildCjGroupedRows, downloadCjUploadsSingle, downloadCjUploadsZip } = load("src/lib/excel/writeCJUploads.ts");
  const rows = [
    { 상품주문번호: "1", 상품약어: "분류A", 품목명: "품목A", 원본엑셀상품명: "원본 상품A", 옵션: "옵션A", 매출처: "거래처A" },
    { 상품주문번호: "2", 상품약어: "분류A", 품목명: "품목B", 원본엑셀상품명: "", 상품명: "다른 상품B", 상품옵션: "옵션B", 매출처: "거래처B" },
    { 상품주문번호: "3", 상품약어: "분류B", 품목명: "품목C", 원본엑셀상품명: "  \n ", 매출처: "거래처C" },
    { 상품주문번호: "4", 상품약어: "분류B", 품목명: "품목D", 매출처: "" },
  ];
  const groups = buildCjGroupedRows(Object.keys(rows[0]), rows);
  downloads = [];
  await downloadCjUploadsSingle(groups);
  await downloadCjUploadsZip(groups);

  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(await downloads[1].blob.arrayBuffer());
  assert.deepEqual(Object.keys(zip.files), ["분류A.xlsx", "분류B.xlsx"]);
  const cases = [
    [await downloads[0].blob.arrayBuffer(), ["원본 상품A", "품목B", "품목C", "품목D"], ["품목A", "품목B", "품목C", "품목D"], ["거래처A", "거래처B", "거래처C", ""]],
    [await zip.file("분류A.xlsx").async("nodebuffer"), ["원본 상품A", "품목B"], ["품목A", "품목B"], ["거래처A", "거래처B"]],
    [await zip.file("분류B.xlsx").async("nodebuffer"), ["품목C", "품목D"], ["품목C", "품목D"], ["거래처C", ""]],
  ];
  for (const [buffer, expected, expectedAdminNames, expectedSalesAccounts] of cases) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.worksheets[0];
    const col = ws.getRow(1).values.indexOf("품목명");
    const actual = Array.from({ length: ws.rowCount - 1 }, (_, index) => ws.getRow(index + 2).getCell(col).value);
    assert.deepEqual(actual, expected);
    assert.equal(ws.getCell("P1").value, "어드민품목명");
    const adminNames = Array.from({ length: ws.rowCount - 1 }, (_, index) => ws.getCell(`P${index + 2}`).value);
    assert.deepEqual(adminNames, expectedAdminNames);
    assert.equal(ws.getCell("Q1").value, "매출처");
    const salesAccounts = Array.from({ length: ws.rowCount - 1 }, (_, index) => ws.getCell(`Q${index + 2}`).value);
    assert.deepEqual(salesAccounts, expectedSalesAccounts);
  }
});
