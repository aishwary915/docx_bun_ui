/**
 * XLSX Export Utility for Univer Sheets
 * Converts Univer workbook data to XLSX format with full formatting preservation
 * Supports: styles, fonts, colors, borders, alignment, formulas, merged cells
 */

import JSZip from "jszip";
import { saveAs } from "file-saver";

interface CellStyle {
  bl?: number; // Bold
  it?: number; // Italic
  ul?: any; // Underline
  st?: any; // Strike-through
  fs?: number; // Font size
  ff?: string; // Font family
  cl?: any; // Text color
  bg?: any; // Background color
  va?: number; // Vertical align (superscript/subscript)
}

// Utility functions
const DocxParsingUtils = {
  escapeXml(str: string): string {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  },

  normalizeColor(color: string): string {
    const hex = color.replace("#", "").toUpperCase();
    return hex.length === 6 ? hex : "000000";
  },
};

/**
 * Export Univer workbook snapshot to XLSX file
 */
export async function exportWorkbookToXlsx(
  snapshot: any,
  filename: string
): Promise<void> {
  try {
    console.log("[XLSX Export] Starting export...");

    const zip = new JSZip();

    const sheets = snapshot.sheets || {};
    const sheetOrder = snapshot.sheetOrder || Object.keys(sheets);
    const styles = snapshot.styles || {};

    // Build comprehensive styles mapping
    const cellStyleMap = new Map<string, number>();
    const cellFormats: string[] = [];
    const fonts: string[] = [];
    const fills: string[] = [];
    const borders: string[] = [];
    const numFmts: string[] = [];

    // Default formats (indices 0-5 are reserved in Excel)
    fonts.push('<font><sz val="11"/><name val="Calibri"/></font>');
    fonts.push('<font><sz val="11"/><name val="Calibri"/><b/></font>');
    fonts.push('<font><sz val="11"/><name val="Calibri"/><i/></font>');

    fills.push('<fill><patternFill patternType="none"/></fill>');
    fills.push('<fill><patternFill patternType="gray125"/></fill>');

    borders.push(
      "<border><left/><right/><top/><bottom/><diagonal/></border>"
    );

    cellFormats.push(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    );
    cellFormats.push(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    );

    // Process styles from snapshot
    Object.entries(styles).forEach(([styleId, style]: [string, any]) => {
      let fontId = 0;
      let fillId = 0;
      let borderId = 0;
      let applyFont = 0;
      let applyFill = 0;
      let applyBorder = 0;
      let applyAlignment = 0;

      // Process font styling
      if (
        style.bl ||
        style.it ||
        style.ul ||
        style.st ||
        style.fs ||
        style.ff ||
        style.cl ||
        style.va
      ) {
        let fontXml = "<font>";
        fontXml += `<sz val="${style.fs || 11}"/>`;

        if (style.bl) fontXml += "<b/>";
        if (style.it) fontXml += "<i/>";
        if (style.ul) fontXml += '<u val="single"/>';
        if (style.st) fontXml += "<strike/>";

        if (style.cl) {
          const colorVal = style.cl.rgb || style.cl;
          const hex = DocxParsingUtils.normalizeColor(colorVal);
          fontXml += `<color rgb="FF${hex}"/>`;
        }

        fontXml += `<name val="${style.ff || "Calibri"}"/>`;

        if (style.va === 1) {
          fontXml += '<vertAlign val="superscript"/>';
        } else if (style.va === 2) {
          fontXml += '<vertAlign val="subscript"/>';
        }

        fontXml += "</font>";
        fontId = fonts.length;
        fonts.push(fontXml);
        applyFont = 1;
      }

      // Process background fill
      if (style.bg) {
        const bgColorVal = style.bg.rgb || style.bg;
        const hex = DocxParsingUtils.normalizeColor(bgColorVal);
        const fillXml = `<fill><patternFill patternType="solid"><fgColor rgb="FF${hex}"/><bgColor indexed="64"/></patternFill></fill>`;
        fillId = fills.length;
        fills.push(fillXml);
        applyFill = 1;
      }

      // Process borders
      if (style.bd) {
        let borderXml = "<border>";
        const { left, right, top, bottom } = style.bd;

        const sides: Array<[string, any]> = [
          ["left", left],
          ["right", right],
          ["top", top],
          ["bottom", bottom],
        ];

        sides.forEach(([sideName, sideStyle]) => {
          if (sideStyle) {
            const borderColor = sideStyle.cl?.rgb || sideStyle.cl || "#000000";
            const hex = DocxParsingUtils.normalizeColor(borderColor);
            borderXml += `<${sideName} style="thin"><color rgb="FF${hex}"/></${sideName}>`;
          } else {
            borderXml += `<${sideName}/>`;
          }
        });

        borderXml += "<diagonal/></border>";
        borderId = borders.length;
        borders.push(borderXml);
        applyBorder = 1;
      }

      // Process alignment
      let alignmentXml = "";
      if (style.al || style.vl) {
        alignmentXml = `<alignment horizontal="${style.al || "left"}" vertical="${
          style.vl || "top"
        }" wrapText="0"/>`;
        applyAlignment = 1;
      }

      // Create cellXf entry
      const xfIdx = cellFormats.length;
      let xfXml = `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`;
      if (applyFont) xfXml += ` applyFont="1"`;
      if (applyFill) xfXml += ` applyFill="1"`;
      if (applyBorder) xfXml += ` applyBorder="1"`;
      if (applyAlignment) xfXml += ` applyAlignment="1"`;
      xfXml += ">";
      if (alignmentXml) xfXml += alignmentXml;
      xfXml += "</xf>";

      cellFormats.push(xfXml);
      cellStyleMap.set(styleId, xfIdx);
    });

    // Helper to get column name from index
    const getColName = (colIndex: number): string => {
      let name = "";
      let num = colIndex + 1;
      while (num > 0) {
        const rem = (num - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        num = Math.floor((num - 1) / 26);
      }
      return name;
    };

    // Build Content_Types.xml
    let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;

    sheetOrder.forEach((_: any, idx: number) => {
      contentTypes += `\n  <Override PartName="/xl/worksheets/sheet${
        idx + 1
      }.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    });

    contentTypes += `\n  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

    zip.file("[Content_Types].xml", contentTypes);

    zip.folder("_rels")?.file(
      ".rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    );

    // Build workbook.xml.rels
    let workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    sheetOrder.forEach((_: any, idx: number) => {
      workbookRels += `\n  <Relationship Id="rId${
        idx + 1
      }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${
        idx + 1
      }.xml"/>`;
    });
    workbookRels += `\n  <Relationship Id="rId${
      sheetOrder.length + 1
    }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
    workbookRels += `\n  <Relationship Id="rId${
      sheetOrder.length + 2
    }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

    const xlFolder = zip.folder("xl");
    xlFolder?.folder("_rels")?.file("workbook.xml.rels", workbookRels);

    // Build shared strings
    const sharedStrings: string[] = [];
    const stringMap = new Map<string, number>();

    // Process all sheets
    const worksheetsFolder = xlFolder?.folder("worksheets");
    sheetOrder.forEach((sheetId: string, sheetIdx: number) => {
      const sheet = sheets[sheetId];
      const cellData = sheet?.cellData || {};
      const mergeData = sheet?.mergeData || [];
      const rowData = sheet?.rowData || {};
      const columnData = sheet?.columnData || {};

      let sheetRows = "";
      let colsXml = "";

      // Handle column widths
      if (columnData && Object.keys(columnData).length > 0) {
        colsXml = "<cols>";
        Object.entries(columnData).forEach(
          ([colIdx, colInfo]: [string, any]) => {
            const width = colInfo.w ? colInfo.w / 7 : 8.43;
            const hidden = colInfo.hd ? 1 : 0;
            colsXml += `<col min="${Number(colIdx) + 1}" max="${
              Number(colIdx) + 1
            }" width="${width}" hidden="${hidden}" customWidth="1"/>`;
          }
        );
        colsXml += "</cols>";
      }

      const rows = Object.keys(cellData)
        .map(Number)
        .sort((a, b) => a - b);
      for (const rowIndex of rows) {
        const cols = cellData[rowIndex];
        const rowInfo = rowData[rowIndex] || {};
        const rowHeight = rowInfo.h ? rowInfo.h : 15;
        const hidden = rowInfo.hd ? 1 : 0;

        const colIndices = Object.keys(cols)
          .map(Number)
          .sort((a, b) => a - b);

        let rowCells = "";
        for (const colIndex of colIndices) {
          const cell = cols[colIndex];
          const cellRef = getColName(colIndex) + (rowIndex + 1);

          let styleIdx = 0;
          if (cell.s) {
            styleIdx = cellStyleMap.get(cell.s) || 0;
          }

          // Handle formulas
          if (cell.f) {
            const formula = cell.f
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;");
            rowCells += `<c r="${cellRef}" s="${styleIdx}"><f>${formula}</f>`;
            if (cell.v !== undefined && cell.v !== null) {
              rowCells += `<v>${cell.v}</v>`;
            }
            rowCells += `</c>`;
          }
          // Handle string values
          else if (typeof cell.v === "string") {
            let strIndex = stringMap.get(cell.v);
            if (strIndex === undefined) {
              strIndex = sharedStrings.length;
              sharedStrings.push(cell.v);
              stringMap.set(cell.v, strIndex);
            }
            rowCells += `<c r="${cellRef}" s="${styleIdx}" t="s"><v>${strIndex}</v></c>`;
          }
          // Handle numeric values
          else if (typeof cell.v === "number") {
            rowCells += `<c r="${cellRef}" s="${styleIdx}"><v>${cell.v}</v></c>`;
          }
          // Handle boolean values
          else if (typeof cell.v === "boolean") {
            rowCells += `<c r="${cellRef}" s="${styleIdx}" t="b"><v>${
              cell.v ? 1 : 0
            }</v></c>`;
          }
        }

        if (rowCells) {
          sheetRows += `<row r="${
            rowIndex + 1
          }" ht="${rowHeight}" hidden="${hidden}" customHeight="1">${rowCells}</row>`;
        }
      }

      // Handle merged cells
      let mergeCellsXml = "";
      if (mergeData && mergeData.length > 0) {
        mergeCellsXml = '<mergeCells count="' + mergeData.length + '">';
        mergeData.forEach((merge: any) => {
          const startCol = getColName(merge.startColumn);
          const endCol = getColName(merge.endColumn);
          mergeCellsXml += `<mergeCell ref="${startCol}${
            merge.startRow + 1
          }:${endCol}${merge.endRow + 1}"/>`;
        });
        mergeCellsXml += "</mergeCells>";
      }

      // Create worksheet XML
      worksheetsFolder?.file(
        `sheet${sheetIdx + 1}.xml`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${colsXml}
  <sheetData>${sheetRows}</sheetData>
  ${mergeCellsXml}
</worksheet>`
      );
    });

    // Create shared strings XML
    const sharedStringsXml = sharedStrings
      .map((s) => `<si><t>${DocxParsingUtils.escapeXml(s)}</t></si>`)
      .join("");
    xlFolder?.file(
      "sharedStrings.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">${sharedStringsXml}</sst>`
    );

    // Create workbook XML
    let sheetsXml = "";
    sheetOrder.forEach((sheetId: string, idx: number) => {
      const sheet = sheets[sheetId];
      sheetsXml += `<sheet name="${DocxParsingUtils.escapeXml(
        sheet?.name || `Sheet${idx + 1}`
      )}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`;
    });

    xlFolder?.file(
      "workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetsXml}</sheets>
</workbook>`
    );

    // Create styles XML
    const numFmtsXml =
      numFmts.length > 0
        ? `<numFmts count="${numFmts.length}">${numFmts.join("")}</numFmts>`
        : "";
    xlFolder?.file(
      "styles.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${numFmtsXml}
  <fonts count="${fonts.length}">${fonts.join("")}</fonts>
  <fills count="${fills.length}">${fills.join("")}</fills>
  <borders count="${borders.length}">${borders.join("")}</borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="${cellFormats.length}">${cellFormats.join("")}</cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`
    );

    // Generate and download
    const blob = await zip.generateAsync({ type: "blob" });
    const exportFilename = filename.endsWith(".xlsx")
      ? filename
      : filename + ".xlsx";
    saveAs(blob, exportFilename);

    console.log(
      "[XLSX Export] ✓ Exported successfully as",
      exportFilename
    );
  } catch (error) {
    console.error("[XLSX Export] Failed:", error);
    throw error;
  }
}
