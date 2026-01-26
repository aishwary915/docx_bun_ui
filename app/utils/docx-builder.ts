// DOCX OOXML Builder - Converts Univer IDocumentData to DOCX format
// High-fidelity export with proper tables, images, text formatting, and lists

import type { IDocumentData } from "@univerjs/core";
import JSZip from "jszip";
import { create } from "xmlbuilder2";

// Image counter for relationship IDs
let globalImageCounter = 1;
let globalRIdCounter = 3; // Start at 3 to leave room for styles and numbering

interface ImageInfo {
  drawingId: string;
  rId: string;
  filename: string;
  width: number;
  height: number;
  base64Data: string;
  extension: string;
}

interface TableRange {
  start: number;
  end: number;
  tableId: string;
}

// ✅ OOXML VALIDATION: Validate document structure before export
function validateOOXMLStructure(documentData: IDocumentData): string[] {
  const errors: string[] = [];
  
  if (!documentData.body) {
    errors.push("Document body is undefined");
    return errors;
  }
  
  const dataStream = documentData.body.dataStream || "";
  const paragraphs = documentData.body.paragraphs || [];
  const customRanges = documentData.body.customRanges || [];

  // Check dataStream ends with \r\n (required by Univer)
  if (!dataStream.endsWith("\r\n")) {
    errors.push("dataStream must end with \\r\\n for valid Univer document");
  }

  // Check all paragraphs have valid startIndex
  paragraphs.forEach((p: any, i: number) => {
    if (p.startIndex >= dataStream.length) {
      errors.push(
        `Paragraph ${i} startIndex ${p.startIndex} exceeds dataStream length ${dataStream.length}`
      );
    }
  });

  // Check customRanges point to valid positions
  customRanges.forEach((range: any, i: number) => {
    if (range.startIndex >= dataStream.length) {
      errors.push(
        `CustomRange ${i} (${range.rangeId}) startIndex ${range.startIndex} exceeds dataStream length`
      );
    }

    // Verify range points to correct marker (only for tables and images, NOT hyperlinks)
    // rangeType: 0 = hyperlink (no marker), 1 = drawing (0x1A marker), 2 = table (0x1A marker)
    if (range.rangeType > 0 && range.startIndex < dataStream.length) {
      const char = dataStream.charAt(range.startIndex);
      const charCode = char.charCodeAt(0);
      // Table marker: 0x1A, Drawing marker: 0x1A
      if (charCode !== 0x1a) {
        errors.push(
          `CustomRange ${i} (${range.rangeId}) type ${
            range.rangeType
          } should point to marker 0x1A but found 0x${charCode.toString(
            16
          )} at position ${range.startIndex}`
        );
      }
    }
  });

  // Check textRuns don't exceed dataStream
  const textRuns = documentData.body?.textRuns || [];
  textRuns.forEach((tr: any, i: number) => {
    if (tr.ed > dataStream.length) {
      errors.push(
        `TextRun ${i} end position ${tr.ed} exceeds dataStream length ${dataStream.length}`
      );
    }
  });

  return errors;
}

export async function buildDocxFromUniverData(
  documentData: IDocumentData
): Promise<Blob> {
  // ✅ Validate structure before export
  const validationErrors = validateOOXMLStructure(documentData);
  if (validationErrors.length > 0) {
    console.error("[DOCX Export] ❌ Validation failed:", validationErrors);
    throw new Error(
      `Invalid document structure (${
        validationErrors.length
      } errors):\n${validationErrors.slice(0, 5).join("\n")}${
        validationErrors.length > 5
          ? `\n... and ${validationErrors.length - 5} more errors`
          : ""
      }`
    );
  }

  console.log("[DOCX Export] ✅ Document structure validated successfully");
  // Reset counters
  globalImageCounter = 1;
  globalRIdCounter = 3;

  const zip = new JSZip();

  // Prepare image info for document building
  const imageInfos: ImageInfo[] = [];
  const drawings = documentData.drawings || {};
  const drawingsOrder = documentData.drawingsOrder || Object.keys(drawings);

  // Process drawings in order
  for (const drawingId of drawingsOrder) {
    const drawing: any = drawings[drawingId]; // Use any to handle Univer's complex drawing types
    if (
      drawing &&
      (drawing.drawingType === "image" || drawing.drawingType === 1) &&
      drawing.imageProperties?.base64Cache
    ) {
      const base64Full = drawing.imageProperties.base64Cache;
      const base64Data = base64Full.includes(",")
        ? base64Full.split(",")[1]
        : base64Full;

      if (base64Data) {
        const extension = getImageExtension(
          drawing.imageProperties.source || base64Full
        );
        const filename = `image${globalImageCounter}.${extension}`;

        imageInfos.push({
          drawingId,
          rId: `rId${globalRIdCounter}`,
          filename,
          width:
            drawing.transform?.size?.width ||
            drawing.docTransform?.size?.width ||
            200,
          height:
            drawing.transform?.size?.height ||
            drawing.docTransform?.size?.height ||
            200,
          base64Data,
          extension,
        });

        globalImageCounter++;
        globalRIdCounter++;
      }
    }
  }

  // Check if document has lists
  const hasLists =
    documentData.body?.paragraphs?.some((p: any) => p.bullet) || false;

  // Build XML content
  const documentXml = buildDocumentXml(documentData, imageInfos);
  const stylesXml = buildStylesXml(documentData);
  const numberingXml = hasLists ? buildNumberingXml(documentData) : null;
  const relsXml = buildRelsXml(documentData, imageInfos, hasLists);
  const contentTypesXml = buildContentTypesXml(documentData, imageInfos);

  // Add files to ZIP
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.folder("_rels")?.file(".rels", buildRootRels());
  zip.folder("word")?.file("document.xml", documentXml);
  zip.folder("word")?.file("styles.xml", stylesXml);
  if (numberingXml) {
    zip.folder("word")?.file("numbering.xml", numberingXml);
  }
  zip.folder("word/_rels")?.file("document.xml.rels", relsXml);

  // Add images to media folder
  for (const imgInfo of imageInfos) {
    try {
      const imageBuffer = Uint8Array.from(atob(imgInfo.base64Data), (c) =>
        c.charCodeAt(0)
      );
      zip.folder("word/media")?.file(imgInfo.filename, imageBuffer);
      console.log(`[DOCX Export] Added image: ${imgInfo.filename}`);
    } catch (err) {
      console.error(
        `[DOCX Export] Failed to add image ${imgInfo.filename}:`,
        err
      );
    }
  }

  console.log("[DOCX Export] ZIP structure created, generating blob...");
  return await zip.generateAsync({ type: "blob" });
}

function buildDocumentXml(
  documentData: IDocumentData,
  imageInfos: ImageInfo[]
): string {
  const dataStream = documentData.body?.dataStream || "";
  const textRuns = documentData.body?.textRuns || [];
  const paragraphs = documentData.body?.paragraphs || [];
  const bodyTables = documentData.body?.tables || [];
  const tables = (documentData as any).tables || {};

  const doc = create({ version: "1.0", encoding: "UTF-8", standalone: "yes" })
    .ele("w:document", {
      "xmlns:wpc":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
      "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "xmlns:o": "urn:schemas-microsoft-com:office:office",
      "xmlns:r":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
      "xmlns:v": "urn:schemas-microsoft-com:vml",
      "xmlns:wp14":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
      "xmlns:wp":
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "xmlns:w10": "urn:schemas-microsoft-com:office:word",
      "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
      "xmlns:wpg":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
      "xmlns:wpi":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
      "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
      "xmlns:wps":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
      "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
      "mc:Ignorable": "w14 wp14",
    })
    .ele("w:body");

  // Build index of table ranges
  const tableRanges: TableRange[] = (bodyTables || []).map((t: any) => ({
    start: t.startIndex,
    end: t.endIndex,
    tableId: t.tableId,
  }));

  // Find image marker positions (0x1A character)
  const imageMarkerPositions: number[] = [];
  for (let i = 0; i < dataStream.length; i++) {
    if (dataStream.charCodeAt(i) === 0x1a) {
      imageMarkerPositions.push(i);
    }
  }

  // Map image markers to image info based on order
  const imageAtPosition: Map<number, ImageInfo> = new Map();
  imageMarkerPositions.forEach((pos, idx) => {
    if (idx < imageInfos.length) {
      imageAtPosition.set(pos, imageInfos[idx]);
    }
  });

  console.log(
    `[DOCX Export] Building document: ${paragraphs.length} paragraphs, ${tableRanges.length} tables, ${imageInfos.length} images`
  );
  console.log(
    `[DOCX Export] Table ranges:`,
    tableRanges.map((t) => `${t.tableId}: ${t.start}-${t.end}`)
  );
  console.log(
    `[DOCX Export] Available table IDs in tables object:`,
    Object.keys(tables)
  );
  console.log(
    `[DOCX Export] Image positions:`,
    Array.from(imageAtPosition.entries()).map(
      ([pos, img]) => `${img.drawingId} at ${pos}`
    )
  );

  // Handle empty document case
  if (paragraphs.length === 0 && tableRanges.length === 0) {
    // Create at least one empty paragraph for valid DOCX
    const p = doc.ele("w:p");
    if (dataStream.length > 0) {
      // If there's content but no paragraphs defined, output the raw text
      const cleanText = dataStream.replace(/[\x00-\x1F]/g, "").trim();
      if (cleanText.length > 0) {
        const r = p.ele("w:r");
        r.ele("w:t", { "xml:space": "preserve" }).txt(cleanText);
      }
    } else {
      p.ele("w:r");
    }
  }

  // Track processed content
  const processedParagraphs = new Set<number>();
  let currentIndex = 0;

  // Process content in document order
  while (currentIndex < dataStream.length && paragraphs.length > 0) {
    // Check if we're at a table position
    const tableRange = tableRanges.find(
      (t) => currentIndex >= t.start && currentIndex <= t.end
    );

    if (tableRange) {
      // Build table
      const tableData = tables[tableRange.tableId];
      if (tableData) {
        console.log(
          `[DOCX Export] ✓ Building table ${tableRange.tableId} at index ${tableRange.start}`
        );
        buildTable(doc, tableData, dataStream, textRuns, paragraphs);
      } else {
        console.error(
          `[DOCX Export] ✗ Table data missing for ${tableRange.tableId}! Available tables:`,
          Object.keys(tables)
        );
      }
      currentIndex = tableRange.end + 1;
      continue;
    }

    // Find the paragraph at or after current position
    const paraIndex = paragraphs.findIndex(
      (p: any) =>
        p.startIndex >= currentIndex && !processedParagraphs.has(p.startIndex)
    );
    if (paraIndex === -1) break;

    const para = paragraphs[paraIndex];

    // Skip if paragraph is inside a table
    const insideTable = tableRanges.some(
      (t) => para.startIndex >= t.start && para.startIndex <= t.end
    );

    if (insideTable) {
      processedParagraphs.add(para.startIndex);
      currentIndex = para.startIndex + 1;
      continue;
    }

    // Calculate paragraph end
    const nextPara = paragraphs[paraIndex + 1];
    const paraStart = para.startIndex;
    const paraEnd = nextPara ? nextPara.startIndex : dataStream.length;

    // Build paragraph
    const p = doc.ele("w:p");

    // Add paragraph properties
    const pPr = p.ele("w:pPr");

    if (para.paragraphStyle) {
      addParagraphProperties(pPr, para.paragraphStyle);
    }

    if (para.bullet) {
      addBulletProperties(pPr, para.bullet);
    }

    // Process paragraph content - include runs that overlap with this paragraph
    const paraRuns = textRuns.filter(
      (run: any) => run.st < paraEnd && run.ed > paraStart
    );

    if (paraRuns.length === 0) {
      // Check if there's an image in this paragraph
      const imgPos = imageMarkerPositions.find(
        (pos) => pos >= paraStart && pos < paraEnd
      );
      if (imgPos !== undefined) {
        const imgInfo = imageAtPosition.get(imgPos);
        if (imgInfo) {
          addImageToRun(p, imgInfo);
        }
      } else {
        // Empty paragraph - add empty run
        p.ele("w:r");
      }
    } else {
      // Process runs, checking for images between/within runs
      let lastRunEnd = paraStart;

      paraRuns.forEach((run: any) => {
        // Clip run to paragraph boundaries
        const runStart = Math.max(run.st, paraStart);
        const runEnd = Math.min(run.ed, paraEnd);

        // Check for image between last run and this run
        const imgBetween = imageMarkerPositions.find(
          (pos) => pos >= lastRunEnd && pos < runStart
        );
        if (imgBetween !== undefined) {
          const imgInfo = imageAtPosition.get(imgBetween);
          if (imgInfo) {
            addImageToRun(p, imgInfo);
          }
        }

        // Get run text from the clipped range, filtering out control characters
        const runText = dataStream.substring(runStart, runEnd);

        // Check if this run contains an image marker
        const imgInRun = imageMarkerPositions.find(
          (pos) => pos >= runStart && pos < runEnd
        );

        if (imgInRun !== undefined) {
          const imgInfo = imageAtPosition.get(imgInRun);

          // Text before image
          const beforeImg = runText.substring(0, imgInRun - runStart);
          if (beforeImg && beforeImg.replace(/[\x00-\x1F]/g, "").length > 0) {
            const r1 = p.ele("w:r");
            if (run.ts && Object.keys(run.ts).length > 0) {
              const rPr = r1.ele("w:rPr");
              addRunProperties(rPr, run.ts);
            }
            r1.ele("w:t", { "xml:space": "preserve" }).txt(
              beforeImg.replace(/[\x00-\x1F]/g, "")
            );
          }

          // Add image
          if (imgInfo) {
            addImageToRun(p, imgInfo);
          }

          // Text after image
          const afterImg = runText.substring(imgInRun - runStart + 1);
          if (afterImg && afterImg.replace(/[\x00-\x1F]/g, "").length > 0) {
            const r2 = p.ele("w:r");
            if (run.ts && Object.keys(run.ts).length > 0) {
              const rPr = r2.ele("w:rPr");
              addRunProperties(rPr, run.ts);
            }
            r2.ele("w:t", { "xml:space": "preserve" }).txt(
              afterImg.replace(/[\x00-\x1F]/g, "")
            );
          }
        } else {
          // Regular text run - filter control characters
          const cleanText = runText.replace(/[\x00-\x1F]/g, "");
          if (cleanText.length > 0) {
            const r = p.ele("w:r");

            if (run.ts && Object.keys(run.ts).length > 0) {
              const rPr = r.ele("w:rPr");
              addRunProperties(rPr, run.ts);
            }

            r.ele("w:t", { "xml:space": "preserve" }).txt(cleanText);
          }
        }

        lastRunEnd = runEnd;
      });

      // Check for image after last run
      const imgAfter = imageMarkerPositions.find(
        (pos) => pos >= lastRunEnd && pos < paraEnd
      );
      if (imgAfter !== undefined) {
        const imgInfo = imageAtPosition.get(imgAfter);
        if (imgInfo) {
          addImageToRun(p, imgInfo);
        }
      }
    }

    processedParagraphs.add(para.startIndex);
    currentIndex = paraEnd;
  }

  // Ensure document has at least one paragraph (required by OOXML spec)
  if (processedParagraphs.size === 0 && tableRanges.length === 0) {
    const p = doc.ele("w:p");
    p.ele("w:r");
  }

  // Add section properties
  const sectPr = doc.ele("w:sectPr");

  // Page size in twips (1 point = 20 twips)
  const pageWidth = documentData.documentStyle?.pageSize?.width || 612; // Letter width in points
  const pageHeight = documentData.documentStyle?.pageSize?.height || 792; // Letter height in points

  sectPr.ele("w:pgSz", {
    "w:w": Math.round(pageWidth * 20).toString(),
    "w:h": Math.round(pageHeight * 20).toString(),
  });

  // Margins in twips
  const marginTop = documentData.documentStyle?.marginTop || 72;
  const marginBottom = documentData.documentStyle?.marginBottom || 72;
  const marginLeft = documentData.documentStyle?.marginLeft || 90;
  const marginRight = documentData.documentStyle?.marginRight || 90;

  sectPr.ele("w:pgMar", {
    "w:top": Math.round(marginTop * 20).toString(),
    "w:bottom": Math.round(marginBottom * 20).toString(),
    "w:left": Math.round(marginLeft * 20).toString(),
    "w:right": Math.round(marginRight * 20).toString(),
    "w:header": "720",
    "w:footer": "720",
    "w:gutter": "0",
  });

  return doc.end({ prettyPrint: true });
}

function addImageToRun(parentElement: any, imgInfo: ImageInfo): void {
  // Convert points to EMUs (English Metric Units)
  // 1 point = 12700 EMUs
  const widthEmu = Math.round(imgInfo.width * 12700);
  const heightEmu = Math.round(imgInfo.height * 12700);

  const r = parentElement.ele("w:r");
  const drawing = r.ele("w:drawing");

  // Inline drawing
  const inline = drawing.ele("wp:inline", {
    distT: "0",
    distB: "0",
    distL: "0",
    distR: "0",
  });

  // Extent (size)
  inline.ele("wp:extent", {
    cx: widthEmu.toString(),
    cy: heightEmu.toString(),
  });

  // Effect extent
  inline.ele("wp:effectExtent", {
    l: "0",
    t: "0",
    r: "0",
    b: "0",
  });

  // Document properties
  const docId = imgInfo.rId.replace("rId", "");
  inline.ele("wp:docPr", {
    id: docId,
    name: `Picture ${docId}`,
  });

  // Non-visual properties
  const cNvGraphicFramePr = inline.ele("wp:cNvGraphicFramePr");
  cNvGraphicFramePr.ele("a:graphicFrameLocks", {
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    noChangeAspect: "1",
  });

  // Graphic element
  const graphic = inline.ele("a:graphic", {
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  });

  const graphicData = graphic.ele("a:graphicData", {
    uri: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  });

  // Picture element
  const pic = graphicData.ele("pic:pic", {
    "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
  });

  // Non-visual picture properties
  const nvPicPr = pic.ele("pic:nvPicPr");
  nvPicPr.ele("pic:cNvPr", {
    id: docId,
    name: imgInfo.filename,
  });
  nvPicPr.ele("pic:cNvPicPr");

  // Blip fill
  const blipFill = pic.ele("pic:blipFill");
  blipFill.ele("a:blip", {
    "r:embed": imgInfo.rId,
  });
  const stretch = blipFill.ele("a:stretch");
  stretch.ele("a:fillRect");

  // Shape properties
  const spPr = pic.ele("pic:spPr");
  const xfrm = spPr.ele("a:xfrm");
  xfrm.ele("a:off", { x: "0", y: "0" });
  xfrm.ele("a:ext", { cx: widthEmu.toString(), cy: heightEmu.toString() });

  const prstGeom = spPr.ele("a:prstGeom", { prst: "rect" });
  prstGeom.ele("a:avLst");

  console.log(
    `[DOCX Export] Added inline image: ${imgInfo.filename} (${imgInfo.width}x${imgInfo.height}pt)`
  );
}

function buildTable(
  doc: any,
  tableData: any,
  dataStream: string,
  textRuns: any[],
  _paragraphs: any[]
): void {
  const tbl = doc.ele("w:tbl");

  // Table properties
  const tblPr = tbl.ele("w:tblPr");
  tblPr.ele("w:tblW", { "w:w": "5000", "w:type": "pct" }); // 100% width

  // Table layout
  tblPr.ele("w:tblLayout", { "w:type": "autofit" });

  // Add table borders
  const tblBorders = tblPr.ele("w:tblBorders");

  if (tableData.tableProperties?.borders) {
    const borders = tableData.tableProperties.borders;
    addTableBorder(tblBorders, "w:top", borders.top);
    addTableBorder(tblBorders, "w:bottom", borders.bottom);
    addTableBorder(tblBorders, "w:left", borders.left);
    addTableBorder(tblBorders, "w:right", borders.right);
    addTableBorder(tblBorders, "w:insideH", borders.insideH || borders.top);
    addTableBorder(tblBorders, "w:insideV", borders.insideV || borders.left);
  } else {
    // Default borders
    const defaultBorder = {
      "w:val": "single",
      "w:sz": "4",
      "w:color": "000000",
      "w:space": "0",
    };
    tblBorders.ele("w:top", defaultBorder);
    tblBorders.ele("w:bottom", defaultBorder);
    tblBorders.ele("w:left", defaultBorder);
    tblBorders.ele("w:right", defaultBorder);
    tblBorders.ele("w:insideH", defaultBorder);
    tblBorders.ele("w:insideV", defaultBorder);
  }

  // Table alignment
  if (tableData.tableProperties?.align !== undefined) {
    const alignMap = ["left", "center", "right"];
    tblPr.ele("w:jc", {
      "w:val": alignMap[tableData.tableProperties.align] || "left",
    });
  }

  // Table cell margins
  const tblCellMar = tblPr.ele("w:tblCellMar");
  tblCellMar.ele("w:top", { "w:w": "55", "w:type": "dxa" });
  tblCellMar.ele("w:left", { "w:w": "108", "w:type": "dxa" });
  tblCellMar.ele("w:bottom", { "w:w": "55", "w:type": "dxa" });
  tblCellMar.ele("w:right", { "w:w": "108", "w:type": "dxa" });

  // Table grid (column definitions)
  const tblGrid = tbl.ele("w:tblGrid");
  const numCols =
    tableData.tableColumns?.length ||
    tableData.tableRows?.[0]?.cells?.length ||
    3;

  for (let i = 0; i < numCols; i++) {
    const col = tableData.tableColumns?.[i];
    const colWidth = col?.size ? Math.round(col.size * 20) : 2880; // Default ~2 inches
    tblGrid.ele("w:gridCol", { "w:w": colWidth.toString() });
  }

  // Table rows
  const rows = tableData.tableRows || [];
  rows.forEach((row: any, rowIndex: number) => {
    const tr = tbl.ele("w:tr");

    // Row properties
    const trPr = tr.ele("w:trPr");
    if (row.height || row.h) {
      trPr.ele("w:trHeight", {
        "w:val": Math.round((row.height || row.h || 20) * 20).toString(),
        "w:hRule": "atLeast",
      });
    }

    // Cells
    const cells = row.cells || [];
    cells.forEach((cell: any, _cellIndex: number) => {
      // Skip cells that are continuation of vertical merge
      if (cell.rowSpan === 0) return;

      const tc = tbl.ele("w:tc");

      // Cell properties
      const tcPr = tc.ele("w:tcPr");

      // Cell width
      if (cell.width) {
        tcPr.ele("w:tcW", {
          "w:w": Math.round(cell.width * 20).toString(),
          "w:type": "dxa",
        });
      } else {
        tcPr.ele("w:tcW", { "w:w": "0", "w:type": "auto" });
      }

      // Column span
      if (cell.colSpan && cell.colSpan > 1) {
        tcPr.ele("w:gridSpan", { "w:val": cell.colSpan.toString() });
      }

      // Row span (vertical merge)
      if (cell.rowSpan && cell.rowSpan > 1) {
        tcPr.ele("w:vMerge", { "w:val": "restart" });
      }

      // Vertical alignment
      if (cell.verticalAlign !== undefined) {
        const vAlignMap = ["top", "center", "bottom"];
        tcPr.ele("w:vAlign", {
          "w:val": vAlignMap[cell.verticalAlign] || "top",
        });
      }

      // Background color
      if (cell.background?.rgb) {
        const bgColor = cell.background.rgb.replace("#", "");
        tcPr.ele("w:shd", {
          "w:val": "clear",
          "w:color": "auto",
          "w:fill": bgColor,
        });
      }

      // Cell borders
      if (cell.borders) {
        const tcBorders = tcPr.ele("w:tcBorders");
        if (cell.borders.top) {
          addCellBorder(tcBorders, "w:top", cell.borders.top);
        }
        if (cell.borders.bottom) {
          addCellBorder(tcBorders, "w:bottom", cell.borders.bottom);
        }
        if (cell.borders.left) {
          addCellBorder(tcBorders, "w:left", cell.borders.left);
        }
        if (cell.borders.right) {
          addCellBorder(tcBorders, "w:right", cell.borders.right);
        }
      }

      // Cell content - get text with styling
      const cellContent = getCellContent(cell, dataStream, textRuns);

      if (cellContent.runs.length > 0) {
        const p = tc.ele("w:p");

        // Add paragraph properties if first row (header styling)
        if (rowIndex === 0) {
          const pPr = p.ele("w:pPr");
          pPr.ele("w:jc", { "w:val": "center" });
        }

        cellContent.runs.forEach((runInfo: any) => {
          const r = p.ele("w:r");

          // Apply text styling
          if (runInfo.ts && Object.keys(runInfo.ts).length > 0) {
            const rPr = r.ele("w:rPr");
            addRunProperties(rPr, runInfo.ts);

            // Add bold for header row
            if (rowIndex === 0 && !runInfo.ts.bl) {
              rPr.ele("w:b");
            }
          } else if (rowIndex === 0) {
            // Header row styling
            const rPr = r.ele("w:rPr");
            rPr.ele("w:b");
          }

          r.ele("w:t", { "xml:space": "preserve" }).txt(runInfo.text);
        });
      } else {
        // Empty cell - needs paragraph
        const p = tc.ele("w:p");
        p.ele("w:r");
      }
    });
  });
}

function addTableBorder(parent: any, name: string, border: any): void {
  if (border) {
    parent.ele(name, {
      "w:val": "single",
      "w:sz": Math.round((border.w || 1) * 8).toString(),
      "w:color": (border.cl?.rgb || "#000000").replace("#", ""),
      "w:space": "0",
    });
  } else {
    parent.ele(name, {
      "w:val": "single",
      "w:sz": "4",
      "w:color": "000000",
      "w:space": "0",
    });
  }
}

function addCellBorder(parent: any, name: string, border: any): void {
  parent.ele(name, {
    "w:val": "single",
    "w:sz": Math.round((border.w || 1) * 8).toString(),
    "w:color": (border.cl?.rgb || "#000000").replace("#", ""),
    "w:space": "0",
  });
}

function getCellContent(
  cell: any,
  dataStream: string,
  textRuns: any[]
): { runs: any[] } {
  const runs: any[] = [];

  if (cell.startIndex !== undefined && cell.endIndex !== undefined) {
    // Find text runs that overlap with this cell
    const cellRuns = textRuns.filter(
      (run: any) => run.st < cell.endIndex && run.ed > cell.startIndex
    );

    if (cellRuns.length > 0) {
      cellRuns.forEach((run: any) => {
        const start = Math.max(run.st, cell.startIndex);
        const end = Math.min(run.ed, cell.endIndex);
        const text = dataStream
          .substring(start, end)
          .replace(/[\x00-\x1F]/g, "")
          .trim();

        if (text.length > 0) {
          runs.push({ text, ts: run.ts || {} });
        }
      });
    } else {
      // No text runs, get raw text
      const text = dataStream
        .substring(cell.startIndex, cell.endIndex)
        .replace(/[\x00-\x1F]/g, "")
        .trim();

      if (text.length > 0) {
        runs.push({ text, ts: {} });
      }
    }
  }

  return { runs };
}

function addParagraphProperties(pPr: any, style: any): void {
  // Alignment - Univer uses: undefined=left, 2=center, 3=right, 4=justify
  if (style.horizontalAlign !== undefined) {
    // Map Univer alignment values to OOXML alignment strings
    const alignMap: Record<number, string> = {
      2: "center",   // Univer 2 = center
      3: "right",    // Univer 3 = right
      4: "both",     // Univer 4 = justify
    };
    const alignment = alignMap[style.horizontalAlign];
    if (alignment) {
      pPr.ele("w:jc", { "w:val": alignment });
    }
    // If horizontalAlign is undefined or 0 or 1, omit w:jc (defaults to left)
  }

  // Spacing - combine into single element
  const spacingAttrs: any = {};

  if (style.lineSpacing) {
    spacingAttrs["w:line"] = Math.round(style.lineSpacing * 240).toString();
    spacingAttrs["w:lineRule"] = "auto";
  }

  if (style.spaceAbove?.v) {
    spacingAttrs["w:before"] = Math.round(style.spaceAbove.v * 20).toString();
  }

  if (style.spaceBelow?.v) {
    spacingAttrs["w:after"] = Math.round(style.spaceBelow.v * 20).toString();
  }

  if (Object.keys(spacingAttrs).length > 0) {
    pPr.ele("w:spacing", spacingAttrs);
  }

  // Indentation
  const indAttrs: any = {};

  if (style.indentStart?.v) {
    indAttrs["w:left"] = Math.round(style.indentStart.v * 20).toString();
  }
  if (style.indentEnd?.v) {
    indAttrs["w:right"] = Math.round(style.indentEnd.v * 20).toString();
  }
  if (style.hanging?.v) {
    indAttrs["w:hanging"] = Math.round(style.hanging.v * 20).toString();
  }
  if (style.indentFirstLine?.v) {
    indAttrs["w:firstLine"] = Math.round(
      style.indentFirstLine.v * 20
    ).toString();
  }

  if (Object.keys(indAttrs).length > 0) {
    pPr.ele("w:ind", indAttrs);
  }

  // Paragraph borders
  if (style.border) {
    const pBdr = pPr.ele("w:pBdr");
    if (style.border.bottom) {
      pBdr.ele("w:bottom", {
        "w:val": "single",
        "w:sz": Math.round((style.border.bottom.w || 1) * 8).toString(),
        "w:color": (style.border.bottom.cl?.rgb || "#000000").replace("#", ""),
        "w:space": "1",
      });
    }
    if (style.border.top) {
      pBdr.ele("w:top", {
        "w:val": "single",
        "w:sz": Math.round((style.border.top.w || 1) * 8).toString(),
        "w:color": (style.border.top.cl?.rgb || "#000000").replace("#", ""),
        "w:space": "1",
      });
    }
  }

  // Paragraph background/shading
  if (style.background?.rgb) {
    pPr.ele("w:shd", {
      "w:val": "clear",
      "w:color": "auto",
      "w:fill": style.background.rgb.replace("#", ""),
    });
  }
}

function addRunProperties(rPr: any, ts: any): void {
  // Bold
  if (ts.bl === 1) {
    rPr.ele("w:b");
    rPr.ele("w:bCs");
  }

  // Italic
  if (ts.it === 1) {
    rPr.ele("w:i");
    rPr.ele("w:iCs");
  }

  // Underline
  if (ts.ul) {
    rPr.ele("w:u", { "w:val": "single" });
  }

  // Strike
  if (ts.st === 1 || ts.st?.s === 1) {
    rPr.ele("w:strike");
  }

  // Font size (convert to half-points)
  if (ts.fs) {
    rPr.ele("w:sz", { "w:val": Math.round(ts.fs * 2).toString() });
    rPr.ele("w:szCs", { "w:val": Math.round(ts.fs * 2).toString() });
  }

  // Font family
  if (ts.ff) {
    rPr.ele("w:rFonts", {
      "w:ascii": ts.ff,
      "w:hAnsi": ts.ff,
      "w:cs": ts.ff,
      "w:eastAsia": ts.ff,
    });
  }

  // Text color - handle both formats
  if (ts.cl) {
    let color = null;
    if (ts.cl.rgb) {
      color = ts.cl.rgb.replace("#", "");
    } else if (typeof ts.cl === "string") {
      color = ts.cl.replace("#", "");
    }
    if (color && color !== "000000") {
      rPr.ele("w:color", { "w:val": color });
    }
  }

  // Background/highlight - handle both formats
  if (ts.bg) {
    let bgColor = null;
    if (ts.bg.rgb) {
      bgColor = ts.bg.rgb;
    } else if (typeof ts.bg === "string") {
      bgColor = ts.bg;
    }

    if (bgColor) {
      const highlightColor = mapRgbToHighlight(bgColor);
      if (highlightColor) {
        rPr.ele("w:highlight", { "w:val": highlightColor });
      } else {
        rPr.ele("w:shd", {
          "w:val": "clear",
          "w:color": "auto",
          "w:fill": bgColor.replace("#", ""),
        });
      }
    }
  }

  // Superscript/Subscript
  if (ts.va === 1) {
    rPr.ele("w:vertAlign", { "w:val": "superscript" });
  } else if (ts.va === 2) {
    rPr.ele("w:vertAlign", { "w:val": "subscript" });
  }
}

function addBulletProperties(pPr: any, bullet: any): void {
  const numPr = pPr.ele("w:numPr");
  numPr.ele("w:ilvl", { "w:val": (bullet.nestingLevel || 0).toString() });
  numPr.ele("w:numId", { "w:val": bullet.listId || "1" });
}

function buildStylesXml(_documentData: IDocumentData): string {
  const styles = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("w:styles", {
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:r":
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "w14",
  });

  // Document defaults
  const docDefaults = styles.ele("w:docDefaults");

  // Run properties defaults
  const rPrDefault = docDefaults.ele("w:rPrDefault").ele("w:rPr");
  rPrDefault.ele("w:rFonts", {
    "w:ascii": "Calibri",
    "w:hAnsi": "Calibri",
    "w:eastAsia": "Calibri",
    "w:cs": "Times New Roman",
  });
  rPrDefault.ele("w:sz", { "w:val": "22" }); // 11pt
  rPrDefault.ele("w:szCs", { "w:val": "22" });
  rPrDefault.ele("w:lang", {
    "w:val": "en-US",
    "w:eastAsia": "en-US",
    "w:bidi": "ar-SA",
  });

  // Paragraph properties defaults
  const pPrDefault = docDefaults.ele("w:pPrDefault").ele("w:pPr");
  pPrDefault.ele("w:spacing", {
    "w:after": "200",
    "w:line": "276",
    "w:lineRule": "auto",
  });

  // Latent styles
  const latentStyles = styles.ele("w:latentStyles", {
    "w:defLockedState": "0",
    "w:defUIPriority": "99",
    "w:defSemiHidden": "0",
    "w:defUnhideWhenUsed": "0",
    "w:defQFormat": "0",
    "w:count": "376",
  });

  latentStyles.ele("w:lsdException", {
    "w:name": "Normal",
    "w:uiPriority": "0",
    "w:qFormat": "1",
  });
  latentStyles.ele("w:lsdException", {
    "w:name": "heading 1",
    "w:uiPriority": "9",
    "w:qFormat": "1",
  });
  latentStyles.ele("w:lsdException", {
    "w:name": "heading 2",
    "w:uiPriority": "9",
    "w:qFormat": "1",
  });

  // Normal style
  const normalStyle = styles.ele("w:style", {
    "w:type": "paragraph",
    "w:styleId": "Normal",
    "w:default": "1",
  });
  normalStyle.ele("w:name", { "w:val": "Normal" });
  normalStyle.ele("w:qFormat");

  // Default paragraph font style
  const defaultFontStyle = styles.ele("w:style", {
    "w:type": "character",
    "w:styleId": "DefaultParagraphFont",
    "w:default": "1",
  });
  defaultFontStyle.ele("w:name", { "w:val": "Default Paragraph Font" });
  defaultFontStyle.ele("w:uiPriority", { "w:val": "1" });
  defaultFontStyle.ele("w:semiHidden");
  defaultFontStyle.ele("w:unhideWhenUsed");

  // Table normal style
  const tableNormalStyle = styles.ele("w:style", {
    "w:type": "table",
    "w:styleId": "TableNormal",
    "w:default": "1",
  });
  tableNormalStyle.ele("w:name", { "w:val": "Normal Table" });
  tableNormalStyle.ele("w:uiPriority", { "w:val": "99" });
  tableNormalStyle.ele("w:semiHidden");
  tableNormalStyle.ele("w:unhideWhenUsed");
  const tblPr = tableNormalStyle.ele("w:tblPr");
  tblPr.ele("w:tblInd", { "w:w": "0", "w:type": "dxa" });
  const tblCellMar = tblPr.ele("w:tblCellMar");
  tblCellMar.ele("w:top", { "w:w": "0", "w:type": "dxa" });
  tblCellMar.ele("w:left", { "w:w": "108", "w:type": "dxa" });
  tblCellMar.ele("w:bottom", { "w:w": "0", "w:type": "dxa" });
  tblCellMar.ele("w:right", { "w:w": "108", "w:type": "dxa" });

  // Heading styles
  for (let i = 1; i <= 6; i++) {
    const headingStyle = styles.ele("w:style", {
      "w:type": "paragraph",
      "w:styleId": `Heading${i}`,
    });
    headingStyle.ele("w:name", { "w:val": `heading ${i}` });
    headingStyle.ele("w:basedOn", { "w:val": "Normal" });
    headingStyle.ele("w:next", { "w:val": "Normal" });
    headingStyle.ele("w:link", { "w:val": `Heading${i}Char` });
    headingStyle.ele("w:uiPriority", { "w:val": "9" });
    headingStyle.ele("w:qFormat");

    const pPr = headingStyle.ele("w:pPr");
    pPr.ele("w:keepNext");
    pPr.ele("w:keepLines");
    pPr.ele("w:spacing", { "w:before": "240", "w:after": "0" });
    pPr.ele("w:outlineLvl", { "w:val": (i - 1).toString() });

    const rPr = headingStyle.ele("w:rPr");
    rPr.ele("w:rFonts", {
      "w:ascii": "Calibri Light",
      "w:hAnsi": "Calibri Light",
      "w:eastAsia": "Yu Gothic Light",
      "w:cs": "Times New Roman",
    });

    // Font size decreases with heading level
    const fontSize = Math.max(22, 48 - i * 6);
    rPr.ele("w:sz", { "w:val": fontSize.toString() });
    rPr.ele("w:szCs", { "w:val": fontSize.toString() });

    // Primary color for headings
    rPr.ele("w:color", { "w:val": "2F5496" });
  }

  // List paragraph style
  const listStyle = styles.ele("w:style", {
    "w:type": "paragraph",
    "w:styleId": "ListParagraph",
  });
  listStyle.ele("w:name", { "w:val": "List Paragraph" });
  listStyle.ele("w:basedOn", { "w:val": "Normal" });
  listStyle.ele("w:uiPriority", { "w:val": "34" });
  listStyle.ele("w:qFormat");
  const listPPr = listStyle.ele("w:pPr");
  listPPr.ele("w:ind", { "w:left": "720" });
  listPPr.ele("w:contextualSpacing");

  return styles.end({ prettyPrint: true });
}

function buildNumberingXml(documentData: IDocumentData): string | null {
  const paragraphs = documentData.body?.paragraphs || [];
  const lists = (documentData as any).lists || {};
  const hasLists = paragraphs.some((p: any) => p.bullet);

  if (!hasLists) return null;

  const numbering = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("w:numbering", {
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:wpc":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:r":
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:wp14":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "xmlns:wp":
      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xmlns:w10": "urn:schemas-microsoft-com:office:word",
    "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "xmlns:wpg":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "xmlns:wpi":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "xmlns:wps":
      "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc:Ignorable": "w14 w15 wp14",
  });

  // Collect unique list IDs from paragraphs
  const uniqueListIds = new Set<string>();
  paragraphs.forEach((p: any) => {
    if (p.bullet?.listId) {
      uniqueListIds.add(p.bullet.listId);
    }
  });

  const listIdArray = Array.from(uniqueListIds);

  // Create abstract numbering definitions
  listIdArray.forEach((listId, index) => {
    const abstractNumId = index.toString();
    const listDef = lists[listId];
    const isBullet =
      listDef?.listType === "BULLET_LIST" ||
      listDef?.listType === 2 ||
      listDef?.listType === "BULLET";

    const abstractNum = numbering.ele("w:abstractNum", {
      "w:abstractNumId": abstractNumId,
      "w15:restartNumberingAfterBreak": "0",
    });

    abstractNum.ele("w:nsid", { "w:val": generateNsid() });
    abstractNum.ele("w:multiLevelType", { "w:val": "hybridMultilevel" });

    // Create 9 levels
    for (let lvl = 0; lvl < 9; lvl++) {
      const level = abstractNum.ele("w:lvl", { "w:ilvl": lvl.toString() });

      const levelDef = listDef?.nestingLevel?.[lvl];
      const startNum = levelDef?.startNumber || 1;

      level.ele("w:start", { "w:val": startNum.toString() });

      if (isBullet) {
        level.ele("w:numFmt", { "w:val": "bullet" });

        // Bullet character based on level
        const bulletChars = ["•", "○", "■", "●", "◦", "▪", "►", "◊", "※"];
        level.ele("w:lvlText", {
          "w:val": bulletChars[lvl % bulletChars.length],
        });
        level.ele("w:lvlJc", { "w:val": "left" });

        // Indentation
        const pPr = level.ele("w:pPr");
        const indent = 720 + lvl * 360;
        pPr.ele("w:ind", {
          "w:left": indent.toString(),
          "w:hanging": "360",
        });

        // Bullet font
        const rPr = level.ele("w:rPr");
        rPr.ele("w:rFonts", {
          "w:ascii": "Symbol",
          "w:hAnsi": "Symbol",
          "w:hint": "default",
        });
      } else {
        // Numbered list
        const numFormats = [
          "decimal",
          "lowerLetter",
          "lowerRoman",
          "decimal",
          "lowerLetter",
          "lowerRoman",
        ];
        level.ele("w:numFmt", { "w:val": numFormats[lvl % numFormats.length] });
        level.ele("w:lvlText", { "w:val": `%${lvl + 1}.` });
        level.ele("w:lvlJc", { "w:val": "left" });

        const pPr = level.ele("w:pPr");
        const indent = 720 + lvl * 360;
        pPr.ele("w:ind", {
          "w:left": indent.toString(),
          "w:hanging": "360",
        });
      }
    }

    // Link numId to abstractNumId
    const num = numbering.ele("w:num", { "w:numId": listId });
    num.ele("w:abstractNumId", { "w:val": abstractNumId });
  });

  return numbering.end({ prettyPrint: true });
}

function generateNsid(): string {
  const chars = "0123456789ABCDEF";
  let result = "";
  for (let i = 0; i < 8; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function buildRelsXml(
  documentData: IDocumentData,
  imageInfos: ImageInfo[],
  hasLists: boolean
): string {
  const rels = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("Relationships", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
  });

  // ✅ FIX: Use sequential rId counter to avoid duplicates
  let rIdCounter = 1;

  // Always add styles first
  rels.ele("Relationship", {
    Id: `rId${rIdCounter++}`,
    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    Target: "styles.xml",
  });

  // Add numbering if lists exist
  if (hasLists) {
    rels.ele("Relationship", {
      Id: `rId${rIdCounter++}`,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
      Target: "numbering.xml",
    });
  }

  // ✅ Update image rIds to match actual counter
  imageInfos.forEach((imgInfo) => {
    imgInfo.rId = `rId${rIdCounter}`; // Update the rId to match sequential order
    rels.ele("Relationship", {
      Id: imgInfo.rId,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      Target: `media/${imgInfo.filename}`,
    });
    rIdCounter++;
  });

  console.log(`[DOCX Export] Generated ${rIdCounter - 1} relationships`);
  return rels.end({ prettyPrint: true });
}

function buildContentTypesXml(
  documentData: IDocumentData,
  imageInfos: ImageInfo[]
): string {
  const contentTypes = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("Types", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
  });

  // Default types
  contentTypes.ele("Default", {
    Extension: "rels",
    ContentType: "application/vnd.openxmlformats-package.relationships+xml",
  });
  contentTypes.ele("Default", {
    Extension: "xml",
    ContentType: "application/xml",
  });

  // Image extensions
  const imageExtensions = new Set(imageInfos.map((img) => img.extension));
  imageExtensions.forEach((ext) => {
    let contentType = "image/png";
    if (ext === "jpg" || ext === "jpeg") contentType = "image/jpeg";
    else if (ext === "gif") contentType = "image/gif";
    else if (ext === "png") contentType = "image/png";
    else if (ext === "bmp") contentType = "image/bmp";
    else if (ext === "tiff" || ext === "tif") contentType = "image/tiff";

    contentTypes.ele("Default", {
      Extension: ext,
      ContentType: contentType,
    });
  });

  // Override types
  contentTypes.ele("Override", {
    PartName: "/word/document.xml",
    ContentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  });
  contentTypes.ele("Override", {
    PartName: "/word/styles.xml",
    ContentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  });

  // Numbering
  const hasLists =
    documentData.body?.paragraphs?.some((p: any) => p.bullet) || false;
  if (hasLists) {
    contentTypes.ele("Override", {
      PartName: "/word/numbering.xml",
      ContentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    });
  }

  return contentTypes.end({ prettyPrint: true });
}

function buildRootRels(): string {
  const rels = create({
    version: "1.0",
    encoding: "UTF-8",
    standalone: "yes",
  }).ele("Relationships", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
  });

  rels.ele("Relationship", {
    Id: "rId1",
    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    Target: "word/document.xml",
  });

  return rels.end({ prettyPrint: true });
}

function mapRgbToHighlight(rgb: string): string | null {
  const hex = rgb.replace("#", "").toUpperCase();

  const colorMap: Record<string, string> = {
    FFFF00: "yellow",
    "00FF00": "green",
    "00FFFF": "cyan",
    FF00FF: "magenta",
    "0000FF": "blue",
    FF0000: "red",
    "000080": "darkBlue",
    "008080": "darkCyan",
    "008000": "darkGreen",
    "800080": "darkMagenta",
    "800000": "darkRed",
    "808000": "darkYellow",
    "808080": "darkGray",
    C0C0C0: "lightGray",
    "000000": "black",
    FFFFFF: "white",
  };

  return colorMap[hex] || null;
}

function getImageExtension(source: string): string {
  const src = source.toLowerCase();
  if (src.includes("image/png") || src.includes(".png")) return "png";
  if (
    src.includes("image/jpeg") ||
    src.includes("image/jpg") ||
    src.includes(".jpg") ||
    src.includes(".jpeg")
  )
    return "jpg";
  if (src.includes("image/gif") || src.includes(".gif")) return "gif";
  if (src.includes("image/bmp") || src.includes(".bmp")) return "bmp";
  if (src.includes("image/webp") || src.includes(".webp")) return "png"; // Convert webp to png
  return "png"; // Default
}

// Export alias for server-side usage
export const buildDocx = buildDocxFromUniverData;
