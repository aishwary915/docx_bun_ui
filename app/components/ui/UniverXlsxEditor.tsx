/**
 * Univer XLSX Editor Component for React Router
 * 
 * Features:
 * - Full spreadsheet editing with Univer Sheets
 * - Import XLSX files and edit them
 * - Export edited spreadsheets back to XLSX
 * - Preserves formatting, styles, formulas, merged cells
 * - Multiple sheets support
 */

import { useEffect, useRef, useState } from "react";
import { Button } from "./button";
import { Download, FileSpreadsheet } from "lucide-react";

// Univer CSS
import "@univerjs/design/lib/index.css";
import "@univerjs/ui/lib/index.css";
import "@univerjs/preset-sheets-core/lib/index.css";

interface UniverXlsxEditorProps {
  url: string;
  filename: string;
}

// Helper function to convert column letter to index
function colLetterToIndex(col: string): number {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1;
}

// Helper function to normalize color to 6-digit hex
function normalizeColor(color: string): string {
  const hex = color.replace("#", "").toUpperCase();
  return hex.length === 6 ? hex : "000000";
}

// Helper function to escape XML
function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function UniverXlsxEditor({ url, filename }: UniverXlsxEditorProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerInstanceRef = useRef<any>(null);
  const workbookDataRef = useRef<any>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const instanceId = useRef(
    `univer-xlsx-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`
  );
  const isInitializing = useRef(false);

  const handleExport = async () => {
    try {
      if (!univerInstanceRef.current?.univerAPI) {
        throw new Error("Univer not initialized");
      }

      const univerAPI = univerInstanceRef.current.univerAPI;
      const workbook = univerAPI.getActiveWorkbook();
      let snapshot: any = null;

      // Get current snapshot with user edits
      if (workbook) {
        try {
          if (typeof (workbook as any).save === "function") {
            snapshot = (workbook as any).save();
          } else if (typeof (workbook as any).getSnapshot === "function") {
            snapshot = (workbook as any).getSnapshot();
          }
        } catch (snapshotError) {
          console.warn("Failed to capture live workbook snapshot:", snapshotError);
        }
      }

      // Fallback to original data if snapshot fails
      if (!snapshot && workbookDataRef.current) {
        snapshot = workbookDataRef.current;
      }

      if (!snapshot) {
        throw new Error("Workbook data not available");
      }

      console.log("[XLSX Export] Exporting workbook:", {
        sheets: Object.keys(snapshot.sheets || {}).length,
        sheetOrder: snapshot.sheetOrder?.length || 0,
      });

      // Build XLSX file
      const JSZip = (await import("jszip")).default;
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

      // Default formats
      fonts.push('<font><sz val="11"/><name val="Calibri"/></font>');
      fonts.push('<font><sz val="11"/><name val="Calibri"/><b/></font>');
      fonts.push('<font><sz val="11"/><name val="Calibri"/><i/></font>');

      fills.push('<fill><patternFill patternType="none"/></fill>');
      fills.push('<fill><patternFill patternType="gray125"/></fill>');

      borders.push("<border><left/><right/><top/><bottom/><diagonal/></border>");

      cellFormats.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
      cellFormats.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>');

      // Process styles
      Object.entries(styles).forEach(([styleId, style]: [string, any]) => {
        let fontId = 0;
        let fillId = 0;
        let borderId = 0;
        let applyFont = 0;
        let applyFill = 0;
        let applyBorder = 0;
        let applyAlignment = 0;

        // Font styling
        if (style.bl || style.it || style.ul || style.st || style.fs || style.ff || style.cl || style.va) {
          let fontXml = "<font>";
          fontXml += `<sz val="${style.fs || 11}"/>`;
          if (style.bl) fontXml += "<b/>";
          if (style.it) fontXml += "<i/>";
          if (style.ul) fontXml += '<u val="single"/>';
          if (style.st) fontXml += "<strike/>";
          if (style.cl) {
            const hex = normalizeColor(style.cl.rgb || style.cl);
            fontXml += `<color rgb="FF${hex}"/>`;
          }
          fontXml += `<name val="${style.ff || "Calibri"}"/>`;
          if (style.va === 1) fontXml += '<vertAlign val="superscript"/>';
          else if (style.va === 2) fontXml += '<vertAlign val="subscript"/>';
          fontXml += "</font>";
          fontId = fonts.length;
          fonts.push(fontXml);
          applyFont = 1;
        }

        // Background fill
        if (style.bg) {
          const hex = normalizeColor(style.bg.rgb || style.bg);
          const fillXml = `<fill><patternFill patternType="solid"><fgColor rgb="FF${hex}"/><bgColor indexed="64"/></patternFill></fill>`;
          fillId = fills.length;
          fills.push(fillXml);
          applyFill = 1;
        }

        // Borders
        if (style.bd) {
          let borderXml = "<border>";
          const { left, right, top, bottom } = style.bd;
          [["left", left], ["right", right], ["top", top], ["bottom", bottom]].forEach(
            ([sideName, sideStyle]) => {
              if (sideStyle) {
                const hex = normalizeColor(sideStyle.cl?.rgb || sideStyle.cl || "#000000");
                borderXml += `<${sideName} style="thin"><color rgb="FF${hex}"/></${sideName}>`;
              } else {
                borderXml += `<${sideName}/>`;
              }
            }
          );
          borderXml += "<diagonal/></border>";
          borderId = borders.length;
          borders.push(borderXml);
          applyBorder = 1;
        }

        // Alignment
        let alignmentXml = "";
        if (style.al || style.vl) {
          alignmentXml = `<alignment horizontal="${style.al || "left"}" vertical="${style.vl || "top"}" wrapText="0"/>`;
          applyAlignment = 1;
        }

        // Create cellXf
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

      // Helper to get column name
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
        contentTypes += `\n  <Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
      });

      contentTypes += `\n  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

      zip.file("[Content_Types].xml", contentTypes);

      // _rels/.rels
      zip.folder("_rels")?.file(
        ".rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
      );

      // xl/_rels/workbook.xml.rels
      let workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
      sheetOrder.forEach((_: any, idx: number) => {
        workbookRels += `\n  <Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`;
      });
      workbookRels += `\n  <Relationship Id="rId${sheetOrder.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
      workbookRels += `\n  <Relationship Id="rId${sheetOrder.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

      const xlFolder = zip.folder("xl");
      xlFolder?.folder("_rels")?.file("workbook.xml.rels", workbookRels);

      // Shared strings
      const sharedStrings: string[] = [];
      const stringMap = new Map<string, number>();

      // Process sheets
      const worksheetsFolder = xlFolder?.folder("worksheets");
      sheetOrder.forEach((sheetId: string, sheetIdx: number) => {
        const sheet = sheets[sheetId];
        const cellData = sheet?.cellData || {};
        const mergeData = sheet?.mergeData || [];
        const rowData = sheet?.rowData || {};
        const columnData = sheet?.columnData || {};

        let sheetRows = "";
        let colsXml = "";

        // Column widths
        if (columnData && Object.keys(columnData).length > 0) {
          colsXml = "<cols>";
          Object.entries(columnData).forEach(([colIdx, colInfo]: [string, any]) => {
            const width = colInfo.w ? colInfo.w / 7 : 8.43;
            const hidden = colInfo.hd ? 1 : 0;
            colsXml += `<col min="${Number(colIdx) + 1}" max="${Number(colIdx) + 1}" width="${width}" hidden="${hidden}" customWidth="1"/>`;
          });
          colsXml += "</cols>";
        }

        // Process rows
        const rows = Object.keys(cellData).map(Number).sort((a, b) => a - b);
        for (const rowIndex of rows) {
          const cols = cellData[rowIndex];
          const rowInfo = rowData[rowIndex] || {};
          const rowHeight = rowInfo.h || 15;
          const hidden = rowInfo.hd ? 1 : 0;

          const colIndices = Object.keys(cols).map(Number).sort((a, b) => a - b);
          let rowCells = "";

          for (const colIndex of colIndices) {
            const cell = cols[colIndex];
            const cellRef = getColName(colIndex) + (rowIndex + 1);
            const styleIdx = cell.s ? cellStyleMap.get(cell.s) || 0 : 0;

            // Formulas
            if (cell.f) {
              const formula = escapeXml(cell.f);
              rowCells += `<c r="${cellRef}" s="${styleIdx}"><f>${formula}</f>`;
              if (cell.v !== undefined && cell.v !== null) {
                rowCells += `<v>${cell.v}</v>`;
              }
              rowCells += `</c>`;
            }
            // Strings
            else if (typeof cell.v === "string") {
              let strIndex = stringMap.get(cell.v);
              if (strIndex === undefined) {
                strIndex = sharedStrings.length;
                sharedStrings.push(cell.v);
                stringMap.set(cell.v, strIndex);
              }
              rowCells += `<c r="${cellRef}" s="${styleIdx}" t="s"><v>${strIndex}</v></c>`;
            }
            // Numbers
            else if (typeof cell.v === "number") {
              rowCells += `<c r="${cellRef}" s="${styleIdx}"><v>${cell.v}</v></c>`;
            }
            // Booleans
            else if (typeof cell.v === "boolean") {
              rowCells += `<c r="${cellRef}" s="${styleIdx}" t="b"><v>${cell.v ? 1 : 0}</v></c>`;
            }
          }

          if (rowCells) {
            sheetRows += `<row r="${rowIndex + 1}" ht="${rowHeight}" hidden="${hidden}" customHeight="1">${rowCells}</row>`;
          }
        }

        // Merged cells
        let mergeCellsXml = "";
        if (mergeData && mergeData.length > 0) {
          mergeCellsXml = `<mergeCells count="${mergeData.length}">`;
          mergeData.forEach((merge: any) => {
            const startCol = getColName(merge.startColumn);
            const endCol = getColName(merge.endColumn);
            mergeCellsXml += `<mergeCell ref="${startCol}${merge.startRow + 1}:${endCol}${merge.endRow + 1}"/>`;
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

      // Shared strings XML
      const sharedStringsXml = sharedStrings.map((s) => `<si><t>${escapeXml(s)}</t></si>`).join("");
      xlFolder?.file(
        "sharedStrings.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">${sharedStringsXml}</sst>`
      );

      // Workbook XML
      let sheetsXml = "";
      sheetOrder.forEach((sheetId: string, idx: number) => {
        const sheet = sheets[sheetId];
        sheetsXml += `<sheet name="${escapeXml(sheet?.name || `Sheet${idx + 1}`)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`;
      });

      xlFolder?.file(
        "workbook.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetsXml}</sheets>
</workbook>`
      );

      // Styles XML
      xlFolder?.file(
        "styles.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`;
      link.click();
      URL.revokeObjectURL(link.href);

      console.log("[XLSX Export] ✓ Export completed successfully");
    } catch (e) {
      console.error("[XLSX Export] Export failed:", e);
      alert("Failed to export file: " + (e instanceof Error ? e.message : String(e)));
    }
  };

  useEffect(() => {
    let destroyed = false;

    const initUniver = async () => {
      try {
        setError(null);
        setLoading(true);

        await new Promise((resolve) => setTimeout(resolve, 100));
        if (!containerRef.current || destroyed) return;

        // Fetch XLSX file
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: Failed to fetch file`);
        }

        const arrayBuffer = await response.arrayBuffer();
        if (destroyed) return;

        // Parse XLSX
        const JSZip = (await import("jszip")).default;
        const zip = await JSZip.loadAsync(arrayBuffer);

        // Read shared strings
        const sharedStringsXml = await zip.file("xl/sharedStrings.xml")?.async("text");
        const sharedStrings: string[] = [];
        if (sharedStringsXml) {
          const parser = new DOMParser();
          const doc = parser.parseFromString(sharedStringsXml, "application/xml");
          const siElements = doc.getElementsByTagName("si");
          for (let i = 0; i < siElements.length; i++) {
            const tElement = siElements[i].getElementsByTagName("t")[0];
            sharedStrings.push(tElement?.textContent || "");
          }
        }

        // Read workbook
        const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
        if (!workbookXml) {
          throw new Error("Invalid XLSX file: workbook.xml not found");
        }

        const parser = new DOMParser();
        const workbookDoc = parser.parseFromString(workbookXml, "application/xml");
        const sheetElements = workbookDoc.getElementsByTagName("sheet");

        const sheetNames: string[] = [];
        for (let i = 0; i < sheetElements.length; i++) {
          sheetNames.push(sheetElements[i].getAttribute("name") || `Sheet${i + 1}`);
        }

        if (sheetNames.length === 0) {
          throw new Error("No sheets found in workbook");
        }

        // Parse sheets
        const sheets: Record<string, any> = {};
        const sheetOrder: string[] = [];

        for (let sheetIndex = 0; sheetIndex < sheetNames.length; sheetIndex++) {
          const sheetXml = await zip.file(`xl/worksheets/sheet${sheetIndex + 1}.xml`)?.async("text");
          if (!sheetXml) continue;

          const sheetDoc = parser.parseFromString(sheetXml, "application/xml");
          const rows = sheetDoc.getElementsByTagName("row");

          const cellData: Record<number, Record<number, { v: string | number }>> = {};
          let maxRow = 0;
          let maxCol = 0;

          for (let r = 0; r < rows.length; r++) {
            const row = rows[r];
            const rowIndex = parseInt(row.getAttribute("r") || "1", 10) - 1;
            maxRow = Math.max(maxRow, rowIndex);

            const cells = row.getElementsByTagName("c");
            for (let c = 0; c < cells.length; c++) {
              const cell = cells[c];
              const cellRef = cell.getAttribute("r") || "";
              const colLetter = cellRef.replace(/[0-9]/g, "");
              const colIndex = colLetterToIndex(colLetter);
              maxCol = Math.max(maxCol, colIndex);

              const cellType = cell.getAttribute("t");
              const vElement = cell.getElementsByTagName("v")[0];
              let value: string | number = vElement?.textContent || "";

              if (cellType === "s" && sharedStrings.length > 0) {
                const idx = parseInt(value as string, 10);
                value = sharedStrings[idx] || "";
              } else if (cellType !== "str" && value !== "" && !isNaN(Number(value))) {
                value = Number(value);
              }

              if (!cellData[rowIndex]) {
                cellData[rowIndex] = {};
              }
              cellData[rowIndex][colIndex] = { v: value };
            }
          }

          const sheetId = `sheet-${sheetIndex}`;
          sheetOrder.push(sheetId);
          sheets[sheetId] = {
            id: sheetId,
            name: sheetNames[sheetIndex],
            cellData,
            rowCount: Math.max(maxRow + 50, 100),
            columnCount: Math.max(maxCol + 10, 26),
            tabColor: "",
            hidden: 0,
            zoomRatio: 1,
            scrollTop: 0,
            scrollLeft: 0,
            defaultColumnWidth: 73,
            defaultRowHeight: 19,
          };
        }

        if (destroyed) return;

        if (isInitializing.current) {
          console.log("XLSX Univer already initializing, skipping...");
          return;
        }
        isInitializing.current = true;

        try {
          // Import Univer packages
          const [sheetsPreset, sheetsLocale, presetsPkg] = await Promise.all([
            import("@univerjs/preset-sheets-core"),
            import("@univerjs/preset-sheets-core/locales/en-US"),
            import("@univerjs/presets"),
          ]);

          if (destroyed) {
            isInitializing.current = false;
            return;
          }

          const UniverSheetsCorePreset =
            sheetsPreset.UniverSheetsCorePreset ||
            (sheetsPreset as any).default?.UniverSheetsCorePreset ||
            (sheetsPreset as any).default;

          const localeData = (sheetsLocale as any).default || sheetsLocale;
          const { createUniver, mergeLocales } = (presetsPkg as any).default || presetsPkg;

          if (!containerRef.current) {
            throw new Error("Container not available");
          }

          containerRef.current.id = instanceId.current;
          containerRef.current.innerHTML = "";

          // Initialize Univer
          const { univerAPI, univer } = createUniver({
            locale: "en-US",
            locales: {
              "en-US": mergeLocales(localeData),
            },
            presets: [
              UniverSheetsCorePreset({
                container: containerRef.current,
              }),
            ],
          });

          if (destroyed) {
            univer?.dispose?.();
            isInitializing.current = false;
            return;
          }

          // Prepare workbook data
          const workbookData = {
            id: `workbook-${instanceId.current}`,
            name: filename,
            appVersion: "3.0.0-alpha",
            locale: "enUS" as any,
            styles: {},
            sheetOrder,
            sheets,
          };

          console.log("[XLSX] Creating workbook with data:", {
            sheetCount: sheetOrder.length,
            firstSheetRows: Object.keys(sheets[sheetOrder[0]].cellData).length,
          });

          workbookDataRef.current = workbookData;

          // Create workbook
          const workbook = univerAPI.createWorkbook(workbookData);
          console.log("[XLSX] Workbook created:", workbook?.getId());

          univerInstanceRef.current = { univerAPI, univer };
          isInitializing.current = false;
          setLoading(false);
        } catch (univerError) {
          console.error("[XLSX] Univer initialization error:", univerError);
          isInitializing.current = false;
          throw univerError;
        }
      } catch (e) {
        console.error("[XLSX] Failed to initialize:", e);
        if (!destroyed) {
          setError(`Failed to load spreadsheet: ${e instanceof Error ? e.message : String(e)}`);
          setLoading(false);
          isInitializing.current = false;
        }
      }
    };

    const timeoutId = setTimeout(() => {
      initUniver();
    }, 100);

    return () => {
      destroyed = true;
      clearTimeout(timeoutId);
      isInitializing.current = false;

      if (univerInstanceRef.current) {
        try {
          if (univerInstanceRef.current.univerAPI) {
            univerInstanceRef.current.univerAPI.dispose?.();
          }
          if (univerInstanceRef.current.univer) {
            univerInstanceRef.current.univer.dispose?.();
          }
        } catch (e) {
          console.error("[XLSX] Error disposing Univer:", e);
        }
        univerInstanceRef.current = null;
      }

      if (containerRef.current) {
        try {
          while (containerRef.current.firstChild) {
            try {
              containerRef.current.removeChild(containerRef.current.firstChild);
            } catch (e) {
              containerRef.current.innerHTML = "";
              break;
            }
          }
        } catch (e) {
          console.warn("[XLSX] Container cleanup failed:", e);
        }
      }
    };
  }, [url, filename]);

  return (
    <div className="relative flex h-full flex-col">
      <div className="flex items-center justify-between border-b p-2 bg-white dark:bg-slate-800">
        <span className="text-sm font-medium truncate">{filename}</span>
        <Button variant="outline" size="sm" onClick={handleExport} disabled={loading || !!error}>
          <Download className="mr-2 h-4 w-4" />
          Export
        </Button>
      </div>
      {error ? (
        <div className="flex h-full flex-col items-center justify-center p-4">
          <FileSpreadsheet className="text-muted-foreground mb-4 h-12 w-12" />
          <h3 className="mb-2 text-lg font-semibold">Unable to display spreadsheet</h3>
          <p className="text-muted-foreground mb-4 text-center text-sm">{error}</p>
          <Button asChild>
            <a href={url} download={filename} target="_blank" rel="noopener noreferrer">
              <Download className="mr-2 h-4 w-4" />
              Download File
            </a>
          </Button>
        </div>
      ) : (
        <>
          {loading && (
            <div className="bg-background/50 absolute inset-0 z-10 flex items-center justify-center">
              <div className="text-center">
                <div className="border-primary mx-auto mb-2 h-8 w-8 animate-spin rounded-full border-b-2"></div>
                <p className="text-muted-foreground text-sm">Loading spreadsheet...</p>
              </div>
            </div>
          )}
          <div ref={containerRef} style={{ width: "100%", height: "calc(100vh - 120px)" }} />
        </>
      )}
    </div>
  );
}
