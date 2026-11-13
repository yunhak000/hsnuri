import JSZip from "jszip";
import { DOMParser, XMLSerializer, type Document, type Element } from "@xmldom/xmldom";

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export type TCellPatch = {
  address: string;
  value?: string;
  highlight?: boolean;
  text?: boolean;
  styleFrom?: string;
};

export type TSheetPatch = { name: string; cells: TCellPatch[] };

const parse = (xml: string) => new DOMParser({
  onError: (level, message) => {
    if (level !== "warning") throw new Error(`엑셀 내부 XML 오류: ${message}`);
  },
}).parseFromString(xml, "application/xml");

const children = (parent: Element, name: string): Element[] =>
  Array.from(parent.childNodes).filter((node): node is Element =>
    node.nodeType === 1 && (node as Element).localName === name);

const columnNumber = (address: string) => Array.from(address.match(/^[A-Z]+/)?.[0] ?? "")
  .reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);

const requirePart = async (zip: JSZip, path: string) => {
  const part = zip.file(path);
  if (!part) throw new Error(`원본 엑셀 구성요소를 찾지 못했습니다: ${path}`);
  return part.async("string");
};

const resolvePart = (target: string, directory = "xl") => {
  const parts: string[] = [];
  for (const part of (target.startsWith("/") ? target.slice(1) : `${directory}/${target}`).split("/")) {
    if (part === "..") parts.pop();
    else if (part && part !== ".") parts.push(part);
  }
  return parts.join("/");
};

const shiftReferenceUp = (reference: string) => reference.replace(
  /(\$?[A-Z]+\$?)(\d+)/g,
  (_, column: string, row: string) => `${column}${Math.max(1, Number(row) - 1)}`,
);

const removeHeaderRow = (document: Document) => {
  const sheetData = document.getElementsByTagNameNS(NS, "sheetData")[0];
  if (!sheetData) throw new Error("사방넷 첫 번째 시트의 데이터를 찾지 못했습니다.");
  for (const row of children(sheetData, "row")) {
    const rowNumber = Number(row.getAttribute("r"));
    if (rowNumber === 1) {
      sheetData.removeChild(row);
      continue;
    }
    row.setAttribute("r", String(rowNumber - 1));
    for (const cell of children(row, "c")) {
      cell.setAttribute("r", shiftReferenceUp(cell.getAttribute("r") ?? ""));
    }
  }
  // 셀은 서식을 그대로 가지고 한 행 위로 이동한다. 표시 범위와 고정 창 위치도 맞춘다.
  for (const name of ["dimension", "mergeCell", "hyperlink", "conditionalFormatting", "dataValidation", "ignoredError", "selection", "pane", "sheetView"]) {
    for (const element of Array.from(document.getElementsByTagNameNS(NS, name))) {
      for (const attribute of ["ref", "sqref", "activeCell", "topLeftCell"]) {
        if (element.hasAttribute(attribute)) element.setAttribute(attribute, shiftReferenceUp(element.getAttribute(attribute) ?? ""));
      }
      if (name === "pane" && element.hasAttribute("ySplit") && element.getAttribute("state")?.startsWith("frozen")) {
        const ySplit = Math.max(0, Number(element.getAttribute("ySplit")) - 1);
        element.setAttribute("ySplit", String(ySplit));
        if (ySplit === 0 && Number(element.getAttribute("xSplit") ?? 0) === 0) {
          element.parentNode?.removeChild(element);
          for (const selection of Array.from(document.getElementsByTagNameNS(NS, "selection"))) selection.removeAttribute("pane");
        }
      }
    }
  }
  // 첫 행의 제목을 기준으로 하던 필터는 제목 없는 결과물에 남기지 않는다.
  for (const filter of Array.from(document.getElementsByTagNameNS(NS, "autoFilter"))) filter.parentNode?.removeChild(filter);
};

// ExcelJS로 전체 파일을 재생성하지 않고, 원본 ZIP에서 셀과 필요한 스타일만 변경한다.
export const patchOriginalWorkbook = async (
  buffer: ArrayBuffer,
  patches: TSheetPatch[],
  options: { removeFirstSheetHeader?: boolean } = {},
) => {
  const zip = await JSZip.loadAsync(buffer);
  const serializer = new XMLSerializer();
  const workbook = parse(await requirePart(zip, "xl/workbook.xml"));
  const relationships = parse(await requirePart(zip, "xl/_rels/workbook.xml.rels"));
  const targets = new Map(Array.from(relationships.getElementsByTagName("Relationship"))
    .map((node) => [node.getAttribute("Id"), resolvePart(node.getAttribute("Target") ?? "")]));
  const sheetParts = new Map(Array.from(workbook.getElementsByTagNameNS(NS, "sheet"))
    .map((sheet) => [sheet.getAttribute("name"), targets.get(sheet.getAttributeNS(REL_NS, "id"))]));
  const styles = parse(await requirePart(zip, "xl/styles.xml"));
  const fills = styles.getElementsByTagNameNS(NS, "fills")[0];
  const cellXfs = styles.getElementsByTagNameNS(NS, "cellXfs")[0];
  if (!fills || !cellXfs) throw new Error("원본 엑셀의 스타일 정보를 읽을 수 없습니다.");
  const originalXfs = children(cellXfs, "xf");
  const fillIds = new Map<boolean, number>();
  const styleIds = new Map<string, number>();

  const styleFor = (sourceId: number, highlight: boolean | undefined, text: boolean) => {
    if (highlight === undefined && !text) return sourceId;
    const key = `${sourceId}:${highlight}:${text}`;
    const cached = styleIds.get(key);
    if (cached !== undefined) return cached;
    let fillId = fillIds.get(highlight ?? false);
    if (fillId === undefined) {
      fillId = children(fills, "fill").length;
      const fill = styles.createElementNS(NS, "fill");
      const pattern = styles.createElementNS(NS, "patternFill");
      pattern.setAttribute("patternType", highlight ? "solid" : "none");
      if (highlight) {
        const color = styles.createElementNS(NS, "fgColor");
        color.setAttribute("rgb", "FFFFF59D");
        pattern.appendChild(color);
      }
      fill.appendChild(pattern);
      fills.appendChild(fill);
      fills.setAttribute("count", String(fillId + 1));
      fillIds.set(highlight ?? false, fillId);
    }
    const source = originalXfs[sourceId] ?? originalXfs[0];
    if (!source) throw new Error("원본 엑셀의 셀 스타일을 읽을 수 없습니다.");
    const xf = source.cloneNode(true) as Element;
    xf.setAttribute("fillId", String(fillId));
    xf.setAttribute("applyFill", "1");
    if (text) {
      xf.setAttribute("numFmtId", "49");
      xf.setAttribute("applyNumberFormat", "1");
    }
    const id = children(cellXfs, "xf").length;
    cellXfs.appendChild(xf);
    cellXfs.setAttribute("count", String(id + 1));
    styleIds.set(key, id);
    return id;
  };

  for (const patch of patches) {
    if (patch.cells.length === 0) continue;
    const path = sheetParts.get(patch.name);
    if (!path) throw new Error(`원본 시트를 찾지 못했습니다: ${patch.name}`);
    const document: Document = parse(await requirePart(zip, path));
    const sheetData = document.getElementsByTagNameNS(NS, "sheetData")[0];
    if (!sheetData) throw new Error(`시트 데이터를 읽을 수 없습니다: ${patch.name}`);
    const rowMap = new Map(children(sheetData, "row").map((row) => [Number(row.getAttribute("r")), row]));
    const cellMap = new Map(Array.from(document.getElementsByTagNameNS(NS, "c"))
      .map((cell) => [cell.getAttribute("r"), cell]));
    const originalStyleIds = new Map(Array.from(cellMap, ([address, cell]) => [address, cell.getAttribute("s")]));
    let maxRow = 1;
    let maxColumn = 1;

    for (const cellPatch of patch.cells) {
      const rowNumber = Number(cellPatch.address.match(/\d+$/)?.[0]);
      const column = columnNumber(cellPatch.address);
      maxRow = Math.max(maxRow, rowNumber);
      maxColumn = Math.max(maxColumn, column);
      let row = rowMap.get(rowNumber);
      if (!row) {
        row = document.createElementNS(NS, "row");
        row.setAttribute("r", String(rowNumber));
        const next = children(sheetData, "row").find((candidate) => Number(candidate.getAttribute("r")) > rowNumber);
        sheetData.insertBefore(row, next ?? null);
        rowMap.set(rowNumber, row);
      }
      let cell = cellMap.get(cellPatch.address);
      if (!cell) {
        cell = document.createElementNS(NS, "c");
        cell.setAttribute("r", cellPatch.address);
        const next = children(row, "c").find((candidate) => columnNumber(candidate.getAttribute("r") ?? "") > column);
        row.insertBefore(cell, next ?? null);
        cellMap.set(cellPatch.address, cell);
      }
      const styleAddress = cellPatch.styleFrom ?? cellPatch.address;
      const styleColumn = columnNumber(styleAddress);
      const styleRow = rowMap.get(Number(styleAddress.match(/\d+$/)?.[0]));
      const columnStyle = Array.from(document.getElementsByTagNameNS(NS, "col"))
        .find((col) => Number(col.getAttribute("min")) <= styleColumn && Number(col.getAttribute("max")) >= styleColumn)
        ?.getAttribute("style");
      const sourceStyle = Number(originalStyleIds.get(styleAddress) || styleRow?.getAttribute("s") || columnStyle || 0);
      cell.setAttribute("s", String(styleFor(sourceStyle, cellPatch.highlight, cellPatch.text ?? false)));
      if (cellPatch.value !== undefined) {
        while (cell.firstChild) cell.removeChild(cell.firstChild);
        cell.setAttribute("t", "inlineStr");
        const inline = document.createElementNS(NS, "is");
        const text = document.createElementNS(NS, "t");
        text.setAttribute("xml:space", "preserve");
        text.appendChild(document.createTextNode(cellPatch.value));
        inline.appendChild(text);
        cell.appendChild(inline);
      }
    }

    const dimension = document.getElementsByTagNameNS(NS, "dimension")[0];
    if (dimension) {
      const range = dimension.getAttribute("ref") ?? "A1";
      const [start, end = start] = range.split(":");
      maxRow = Math.max(maxRow, Number(end.match(/\d+$/)?.[0] ?? 1));
      maxColumn = Math.max(maxColumn, columnNumber(end));
      let letters = "";
      for (let number = maxColumn; number > 0; number = Math.floor((number - 1) / 26)) {
        letters = String.fromCharCode(65 + (number - 1) % 26) + letters;
      }
      dimension.setAttribute("ref", `${start}:${letters}${maxRow}`);
    }
    zip.file(path, serializer.serializeToString(document));
  }
  if (styleIds.size > 0) zip.file("xl/styles.xml", serializer.serializeToString(styles));
  if (options.removeFirstSheetHeader) {
    const firstSheet = workbook.getElementsByTagNameNS(NS, "sheet")[0];
    const path = firstSheet && sheetParts.get(firstSheet.getAttribute("name"));
    if (!path) throw new Error("사방넷 첫 번째 시트를 찾지 못했습니다.");
    const document = parse(await requirePart(zip, path));
    removeHeaderRow(document);
    zip.file(path, serializer.serializeToString(document));

    // 첫 시트에 연결된 표가 있다면 제목 없는 표로 조정하고 다른 시트의 표는 보존한다.
    const directory = path.slice(0, path.lastIndexOf("/"));
    const fileName = path.slice(path.lastIndexOf("/") + 1);
    const relationPart = zip.file(`${directory}/_rels/${fileName}.rels`);
    if (relationPart) {
      const relations = parse(await relationPart.async("string"));
      for (const relation of Array.from(relations.getElementsByTagName("Relationship"))) {
        if (!relation.getAttribute("Type")?.endsWith("/table")) continue;
        const tablePath = resolvePart(relation.getAttribute("Target") ?? "", directory);
        const tableDocument = parse(await requirePart(zip, tablePath));
        const table = tableDocument.documentElement;
        if (!table) continue;
        const originalRef = table.getAttribute("ref") ?? "";
        const removedTableHeader = Number(originalRef.split(":")[0].match(/\d+$/)?.[0]) === 1;
        table.setAttribute("ref", shiftReferenceUp(originalRef));
        if (removedTableHeader) table.setAttribute("headerRowCount", "0");
        for (const name of ["autoFilter", "sortState"]) {
          for (const element of Array.from(table.getElementsByTagNameNS(NS, name))) {
            if (removedTableHeader) element.parentNode?.removeChild(element);
            else if (element.hasAttribute("ref")) element.setAttribute("ref", shiftReferenceUp(element.getAttribute("ref") ?? ""));
          }
        }
        zip.file(tablePath, serializer.serializeToString(tableDocument));
      }
    }
  }
  return zip.generateAsync({ type: "arraybuffer" });
};
