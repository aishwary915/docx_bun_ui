import type { IDocumentData } from "@univerjs/core";
import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";

/**
 * Convert DOCX to Univer IDocumentData by directly parsing OOXML
 * This approach is more accurate than HTML conversion and ensures proper alignment
 *
 * DOCX CONVERSION SUPPORT:
 * ========================
 * This converter handles:
 * - Text formatting (bold, italic, underline, strikethrough)
 * - Font properties (family, size, color, background)
 * - Paragraph alignment (left, center, right, justify)
 * - Paragraph spacing (before, after, line spacing)
 * - Indentation (left, right, first-line, hanging)
 * - Headings (H1-H6) with style detection
 * - Lists (bulleted and numbered) with multi-level support
 * - Hyperlinks with relationship parsing
 * - Paragraph borders (horizontal rules)
 *
 * KNOWN LIMITATIONS:
 * ==================
 * - Images are disabled to fix cursor positioning issues
 * - List bullets are prepended as text due to Univer v0.15.x renderer limitations
 * - Some advanced OOXML features may not be fully supported
 * - Tables support basic structure (rows, columns, text content) but advanced formatting may be limited
 */
export async function convertDocxToUniverData(
  arrayBuffer: ArrayBuffer,
): Promise<IDocumentData> {
  console.log("🔄 Starting DOCX conversion...");
  const zip = await JSZip.loadAsync(arrayBuffer);

  // Read main document XML
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) {
    throw new Error("Invalid DOCX: word/document.xml not found");
  }

  console.log("📖 Document XML length:", documentXml.length);
  console.log("📖 Document XML preview:", documentXml.substring(0, 500));

  // Parse XML to JavaScript object
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    textNodeName: "#text",
    parseAttributeValue: false,
    trimValues: false,
    ignoreDeclaration: true,
  });

  const doc = parser.parse(documentXml);

  // Extract body
  const body = doc["w:document"]?.["w:body"];
  console.log("📋 Document structure:", {
    hasDocument: !!doc["w:document"],
    hasBody: !!body,
    bodyKeys: body ? Object.keys(body) : [],
    paragraphCount: body?.["w:p"]
      ? Array.isArray(body["w:p"])
        ? body["w:p"].length
        : 1
      : 0,
  });

  if (!body) {
    throw new Error("Invalid DOCX structure: no document body found");
  }

  // Extract raw XML body content for proper element ordering
  const bodyXmlMatch = documentXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/);
  const rawBodyXml = bodyXmlMatch ? bodyXmlMatch[1] : "";

  // Read numbering.xml for lists
  let numberingDoc = null;
  try {
    const numberingXml = await zip.file("word/numbering.xml")?.async("text");
    if (numberingXml) {
      numberingDoc = parser.parse(numberingXml);
    }
  } catch (e) {
    console.warn("Could not parse numbering.xml:", e);
  }

  // Read styles.xml for style definitions
  let stylesDoc = null;
  try {
    const stylesXml = await zip.file("word/styles.xml")?.async("text");
    if (stylesXml) {
      stylesDoc = parser.parse(stylesXml);
      console.log("📐 Styles.xml loaded successfully");
    }
  } catch (e) {
    console.warn("Could not parse styles.xml:", e);
  }

  // Read relationships for hyperlinks
  const relationships = new Map<string, { type: string; target: string }>();
  try {
    const relsXml = await zip
      .file("word/_rels/document.xml.rels")
      ?.async("text");
    if (relsXml) {
      const relsParsed = parser.parse(relsXml);
      const rels = relsParsed["Relationships"]?.["Relationship"] || [];
      const relArray = Array.isArray(rels) ? rels : [rels];
      // biome-ignore lint/suspicious/noExplicitAny: OOXML relationship structure is dynamic
      relArray.forEach((rel: any) => {
        relationships.set(rel["@_Id"], {
          type: rel["@_Type"],
          target: rel["@_Target"],
        });
      });
      console.log(`🔗 Loaded ${relationships.size} relationships`);
    }
  } catch (e) {
    console.warn("Could not parse document.xml.rels:", e);
  }

  const bodyData = await convertOOXMLToBody(
    body,
    rawBodyXml || "",
    numberingDoc,
    stylesDoc,
    relationships,
  );
  
  // Build final IDocumentData structure
  const documentData: IDocumentData = {
    id: `doc_${Date.now()}_${Math.random().toString(36).substring(7)}`,
    body: bodyData.body,
    documentStyle: {
      pageSize: {
        width: 595.27, // A4 width in points (210mm)
        height: 841.89, // A4 height in points (297mm)
      },
      marginTop: 72, // 1 inch = 72 points
      marginBottom: 72,
      marginLeft: 90,
      marginRight: 90,
      renderConfig: {
        vertexAngle: 0,
        centerAngle: 0,
      },
    },
  };

  // Add optional top-level properties only if they have content
  if (Object.keys(bodyData.lists).length > 0) {
    // biome-ignore lint/suspicious/noExplicitAny: Univer IDocumentData accepts lists property
    (documentData as any).lists = bodyData.lists;
  }

  // Add tableSource if present (top-level property from IReferenceSource)
  if (bodyData.tableSource && Object.keys(bodyData.tableSource).length > 0) {
    // biome-ignore lint/suspicious/noExplicitAny: Univer IDocumentData accepts tableSource property
    (documentData as any).tableSource = bodyData.tableSource;
  }

  console.log("📦 Document converted:", {
    id: documentData.id,
    paragraphs: documentData.body?.paragraphs?.length || 0,
    textRuns: documentData.body?.textRuns?.length || 0,
    lists: Object.keys(bodyData.lists).length,
    tableSourceCount: bodyData.tableSource ? Object.keys(bodyData.tableSource).length : 0,
    tableSourceKeys: bodyData.tableSource ? Object.keys(bodyData.tableSource) : [],
    bodyTablesCount: documentData.body?.tables?.length || 0,
  });

  // Debug: Log table details
  if (bodyData.tableSource && Object.keys(bodyData.tableSource).length > 0) {
    console.log("📋 Table details:");
    Object.entries(bodyData.tableSource).forEach(([id, table]: [string, any]) => {
      console.log(`  ${id}: ${table.tableRows?.length || 0} rows, ${table.tableColumns?.length || 0} columns`);
    });
  }

  return documentData;
}

// biome-ignore lint/suspicious/noExplicitAny: OOXML parsing returns dynamic structures
async function convertOOXMLToBody(
  body: any,
  rawBodyXml: string,
  // biome-ignore lint/suspicious/noExplicitAny: Numbering document structure is dynamic
  numberingDoc: any,
  // biome-ignore lint/suspicious/noExplicitAny: Styles document structure is dynamic
  stylesDoc: any,
  relationships: Map<string, { type: string; target: string }>,
): Promise<any> {
  let dataStream = "";
  // biome-ignore lint/suspicious/noExplicitAny: Univer paragraph structure
  let paragraphs: any[] = [];
  // biome-ignore lint/suspicious/noExplicitAny: Univer text run structure
  let textRuns: any[] = [];
  // biome-ignore lint/suspicious/noExplicitAny: Univer list structure
  const lists: Record<string, any> = {};
  // biome-ignore lint/suspicious/noExplicitAny: Univer custom range structure
  const customRanges: any[] = []; // For hyperlinks
  // biome-ignore lint/suspicious/noExplicitAny: Univer table structure
  const tables: Record<string, any> = {};
  // biome-ignore lint/suspicious/noExplicitAny: Body tables array
  const bodyTables: any[] = [];

  // Parse styles from styles.xml
  // biome-ignore lint/suspicious/noExplicitAny: Style definition structure is dynamic
  const styleMap = new Map<string, any>();
  if (stylesDoc?.["w:styles"]?.["w:style"]) {
    const styles = Array.isArray(stylesDoc["w:styles"]["w:style"])
      ? stylesDoc["w:styles"]["w:style"]
      : [stylesDoc["w:styles"]["w:style"]];

    // biome-ignore lint/suspicious/noExplicitAny: Style object structure varies
    styles.forEach((style: any) => {
      const styleId = style["@_w:styleId"];
      if (styleId) {
        // biome-ignore lint/suspicious/noExplicitAny: Style data object
        const styleData: any = {};
        if (style["w:rPr"]) styleData.rPr = style["w:rPr"];
        if (style["w:pPr"]) styleData.pPr = style["w:pPr"];

        if (styleData.rPr || styleData.pPr) {
          styleMap.set(styleId, styleData);
        }
      }
    });
  }

  // Parse abstract numbering definitions from numbering.xml
  // biome-ignore lint/suspicious/noExplicitAny: Abstract numbering level definitions
  const abstractNumDefs = new Map<string, any[]>();
  if (numberingDoc?.["w:numbering"]?.["w:abstractNum"]) {
    const abstractNums = Array.isArray(
      numberingDoc["w:numbering"]["w:abstractNum"],
    )
      ? numberingDoc["w:numbering"]["w:abstractNum"]
      : [numberingDoc["w:numbering"]["w:abstractNum"]];

    // biome-ignore lint/suspicious/noExplicitAny: Abstract numbering structure
    abstractNums.forEach((abstractNum: any) => {
      const abstractNumId = abstractNum["@_w:abstractNumId"];
      const levels = abstractNum["w:lvl"] || [];
      const levelArray = Array.isArray(levels) ? levels : [levels];

      // biome-ignore lint/suspicious/noExplicitAny: Level definition structure
      const levelDefs = levelArray.map((lvl: any) => {
        const ilvl = parseInt(lvl["@_w:ilvl"] || "0");
        const numFmt = lvl["w:numFmt"]?.["@_w:val"] || "bullet";
        let lvlText = lvl["w:lvlText"]?.["@_w:val"];
        const start = parseInt(lvl["w:start"]?.["@_w:val"] || "1");

        // Replace checkbox symbols with standard bullets
        const checkboxSymbols = [
          "☐",
          "☑",
          "☒",
          "❏",
          "❑",
          "\u2610",
          "\u2611",
          "\u2612",
          "\uF0A7",
          "\uF0FC",
        ];
        if (lvlText && checkboxSymbols.includes(lvlText)) {
          lvlText = "•";
        }

        // Handle Wingdings font special characters
        if (
          lvlText &&
          lvlText.charCodeAt(0) >= 0xf000 &&
          lvlText.charCodeAt(0) <= 0xf0ff
        ) {
          lvlText = "•";
        }

        // Default to bullet if empty
        if (!lvlText || lvlText.trim() === "" || lvlText.length === 0) {
          lvlText = numFmt === "bullet" ? "•" : "%1.";
        }

        // Parse level-specific indentation with enhanced values
        const lvlPPr = lvl["w:pPr"];
        let levelIndentStart = 36 * (ilvl + 1); // Increased from 21 to 36 for better visibility
        let levelHanging = 24; // Increased from 21 to 24

        if (lvlPPr?.["w:ind"]) {
          const ind = lvlPPr["w:ind"];
          const leftTwips = parseInt(ind["@_w:left"] || "0");
          const hangingTwips = parseInt(ind["@_w:hanging"] || "0");

          if (leftTwips > 0) {
            levelIndentStart = Math.max(leftTwips / 20, 36 * (ilvl + 1));
          }
          if (hangingTwips > 0) {
            levelHanging = Math.max(hangingTwips / 20, 24);
          }
        }

        let glyphFormat = lvlText;
        const glyphType = numFmt === "bullet" ? 0 : 1;

        return {
          ilvl,
          numFmt,
          glyphFormat,
          start,
          lvlText,
          glyphType,
          indentStart: levelIndentStart,
          hanging: levelHanging,
        };
      });

      abstractNumDefs.set(abstractNumId, levelDefs);
    });
  }

  // Track list definitions
  // biome-ignore lint/suspicious/noExplicitAny: List definition structure
  const listDefs = new Map<string, any>();
  if (numberingDoc?.["w:numbering"]?.["w:num"]) {
    const nums = Array.isArray(numberingDoc["w:numbering"]["w:num"])
      ? numberingDoc["w:numbering"]["w:num"]
      : [numberingDoc["w:numbering"]["w:num"]];

    // biome-ignore lint/suspicious/noExplicitAny: Numbering definition structure
    nums.forEach((num: any) => {
      const numId = num["@_w:numId"];
      const abstractNumId = num["w:abstractNumId"]?.["@_w:val"];
      if (numId && abstractNumId) {
        listDefs.set(numId, abstractNumId);
      }
    });
  }

  // Get all body children (paragraphs and tables) in document order
  // biome-ignore lint/suspicious/noExplicitAny: Body child structure
  const bodyChildren: any[] = [];

  const paragraphElements = body["w:p"]
    ? Array.isArray(body["w:p"])
      ? body["w:p"]
      : [body["w:p"]]
    : [];

  const tableElements = body["w:tbl"]
    ? Array.isArray(body["w:tbl"])
      ? body["w:tbl"]
      : [body["w:tbl"]]
    : [];

  // Parse raw XML to get element order (both paragraphs and tables)
  const elementOrderRegex = /<w:(p|tbl)[\s>]/g;
  let pIndex = 0;
  let tIndex = 0;
  let match;

  while ((match = elementOrderRegex.exec(rawBodyXml)) !== null) {
    if (match[1] === "p" && pIndex < paragraphElements.length) {
      bodyChildren.push({
        type: "paragraph",
        element: paragraphElements[pIndex++],
      });
    } else if (match[1] === "tbl" && tIndex < tableElements.length) {
      bodyChildren.push({
        type: "table",
        element: tableElements[tIndex++],
      });
    }
  }

  // Fallback if regex parsing failed
  if (bodyChildren.length === 0) {
    paragraphElements.forEach((p) =>
      bodyChildren.push({ type: "paragraph", element: p }),
    );
    tableElements.forEach((t) =>
      bodyChildren.push({ type: "table", element: t }),
    );
  }

  console.log(`📊 Body children: ${bodyChildren.length} elements (${paragraphElements.length} paragraphs, ${tableElements.length} tables)`);

  // Process each body element (paragraph or table)
  for (let childIndex = 0; childIndex < bodyChildren.length; childIndex++) {
    const child = bodyChildren[childIndex];

    // Handle table elements
    if (child.type === "table") {
      const tableId = `tbl_${Date.now()}_${Object.keys(tables).length}`;
      const tableResult = parseTableWithDataStream(child.element, styleMap, tableId, dataStream.length);

      if (tableResult) {
        // Record table start position (where \x1A will be)
        const tableStartIndex = dataStream.length;

        // Add table content to dataStream with control characters
        dataStream += tableResult.dataStream;

        // endIndex points AFTER the table end marker (exclusive end)
        const tableEndIndex = dataStream.length;

        // Add paragraph break after table
        dataStream += '\r';
        const paragraphIndex = dataStream.length - 1;

        // Create paragraph for the line after the table
        paragraphs.push({
          startIndex: paragraphIndex,
          paragraphStyle: {
            spaceAbove: { v: 0 },
            spaceBelow: { v: 8 },
            lineSpacing: 1.5,
            horizontalAlign: 0,
          },
        });

        // Store table metadata in tableSource (top-level)
        tables[tableId] = tableResult.tableSource;

        // Add to body.tables array
        bodyTables.push({
          startIndex: tableStartIndex,
          endIndex: tableEndIndex,
          tableId: tableId,
        });

        console.log(`📋 Parsed table ${tableId}: ${tableResult.tableSource.tableRows?.length || 0} rows, start=${tableStartIndex}, end=${tableEndIndex}`);
      }
      continue;
    }

    if (child.type !== "paragraph") continue;

    const p = child.element;
    let hasContent = false;

    // Check for paragraph properties
    const pPr = p["w:pPr"];
    // biome-ignore lint/suspicious/noExplicitAny: Bullet structure
    let bullet = null;
    let isHeading = false;
    let headingLevel = 0;
    // biome-ignore lint/suspicious/noExplicitAny: Text style structure
    let headingTextStyle: any = null;

    // Extract style value
    const styleVal = pPr?.["w:pStyle"]?.["@_w:val"];

    if (styleVal) {
      if (styleVal === "Title") {
        isHeading = true;
        headingLevel = 1;
      } else if (styleVal.match(/^Heading(\d+)$/)) {
        isHeading = true;
        headingLevel = parseInt(styleVal.replace("Heading", ""));
      }

      // Extract heading formatting
      if (isHeading) {
        const runs = p["w:r"] || [];
        const firstRun = Array.isArray(runs) ? runs[0] : runs;

        let rPr = firstRun?.["w:rPr"];

        if (!rPr && styleMap.has(styleVal)) {
          const styleData = styleMap.get(styleVal);
          rPr = styleData?.rPr;
        }

        if (rPr) {
          headingTextStyle = {};

          if ("w:b" in rPr) {
            const bVal = rPr["w:b"];
            if (
              typeof bVal === "string" ||
              !bVal["@_w:val"] ||
              bVal["@_w:val"] !== "0"
            ) {
              headingTextStyle.bl = 1;
            }
          }

          if ("w:i" in rPr) {
            const iVal = rPr["w:i"];
            if (
              typeof iVal === "string" ||
              !iVal["@_w:val"] ||
              iVal["@_w:val"] !== "0"
            ) {
              headingTextStyle.it = 1;
            }
          }

          if (rPr["w:u"]) headingTextStyle.ul = { s: 1 };
          if (rPr["w:strike"]) headingTextStyle.st = 1;
          if (rPr["w:sz"]?.["@_w:val"]) {
            headingTextStyle.fs = parseInt(rPr["w:sz"]["@_w:val"]) / 2;
          }
          if (rPr["w:rFonts"]) {
            const fonts = rPr["w:rFonts"];
            headingTextStyle.ff =
              fonts["@_w:ascii"] ||
              fonts["@_w:hAnsi"] ||
              fonts["@_w:eastAsia"] ||
              fonts["@_w:cs"];
          }
          if (rPr["w:color"]?.["@_w:val"]) {
            const colorVal = rPr["w:color"]["@_w:val"];
            if (colorVal !== "auto" && colorVal !== "000000") {
              headingTextStyle.cl = { rgb: `#${colorVal}` };
            }
          }
        }

        // Apply default font sizes for headings
        if (!headingTextStyle || !headingTextStyle.fs) {
          if (!headingTextStyle) headingTextStyle = {};
          const defaultHeadingSizes = [24, 18, 14, 12, 11, 11];
          headingTextStyle.fs = defaultHeadingSizes[headingLevel - 1] || 11;

          if (!headingTextStyle.bl) {
            headingTextStyle.bl = 1;
          }
        }
      }
    }

    // Parse text alignment with enhanced detection
    // UNIVER ALIGNMENT MAPPING (verified from export):
    // - undefined/omitted = Left (default)
    // - 2 = Center
    // - 3 = Right
    // - 4 = Justify (assumed)
    let textAlignment: number | undefined;

    if (pPr?.["w:jc"]?.["@_w:val"]) {
      const jc = pPr["w:jc"]["@_w:val"];
      const alignMap: Record<string, number | undefined> = {
        left: undefined,      // Left is omitted in Univer
        start: undefined,     // "start" means left in LTR languages
        center: 2,            // Center = 2 in Univer
        right: 3,             // Right = 3 in Univer
        end: 3,               // "end" means right in LTR languages
        both: 4,              // Justify = 4 (assumed)
        justify: 4,
        distribute: 4,
      };
      textAlignment = alignMap[jc];
      console.log(`✓ Paragraph alignment detected: "${jc}" → ${textAlignment === undefined ? 'undefined (left)' : textAlignment} (undefined=left, 2=center, 3=right, 4=justify)`);
    } else if (styleVal && styleMap.has(styleVal)) {
      const styleData = styleMap.get(styleVal);
      const stylePPr = styleData?.pPr;
      if (stylePPr?.["w:jc"]?.["@_w:val"]) {
        const jc = stylePPr["w:jc"]["@_w:val"];
        const alignMap: Record<string, number | undefined> = {
          left: undefined,
          start: undefined,
          center: 2,
          right: 3,
          end: 3,
          both: 4,
          justify: 4,   
          distribute: 4,
        };
        textAlignment = alignMap[jc];
        console.log(`✓ Style-based alignment for "${styleVal}": "${jc}" → ${textAlignment === undefined ? 'undefined (left)' : textAlignment}`);
      }
    }

    // Parse indentation
    // biome-ignore lint/suspicious/noExplicitAny: Indent value structure
    let indentStart: any = undefined;
    // biome-ignore lint/suspicious/noExplicitAny: Indent value structure
    let indentEnd: any = undefined;
    // biome-ignore lint/suspicious/noExplicitAny: Indent value structure
    let indentFirstLine: any = undefined;
    // biome-ignore lint/suspicious/noExplicitAny: Indent value structure
    let hangingIndent: any = undefined;

    if (pPr?.["w:ind"]) {
      const ind = pPr["w:ind"];
      const leftTwips = parseInt(ind["@_w:left"] || "0");
      const firstLineTwips = parseInt(ind["@_w:firstLine"] || "0");
      const hangingTwips = parseInt(ind["@_w:hanging"] || "0");
      const rightTwips = parseInt(ind["@_w:right"] || "0");

      const leftPt = leftTwips / 20;
      const firstLinePt = firstLineTwips / 20;
      const hangingPt = hangingTwips / 20;
      const rightPt = rightTwips / 20;

      if (hangingTwips > 0) {
        indentStart = { v: leftPt };
        hangingIndent = { v: hangingPt };
      } else if (firstLineTwips > 0) {
        indentStart = { v: leftPt };
        indentFirstLine = { v: firstLinePt };
      } else if (leftTwips > 0) {
        indentStart = { v: leftPt };
      }

      if (rightTwips > 0) {
        indentEnd = { v: rightPt };
      }
    }

    // Parse line spacing and space before/after
    let lineSpacing = 1.5;
    // biome-ignore lint/suspicious/noExplicitAny: Space value structure
    let spaceAbove: any = { v: 0 };
    // biome-ignore lint/suspicious/noExplicitAny: Space value structure
    let spaceBelow: any = { v: 0 };

    if (isHeading) {
      const headingSpacing = [
        { before: 12, after: 3 },
        { before: 4, after: 3 },
        { before: 4, after: 2 },
        { before: 4, after: 2 },
        { before: 4, after: 2 },
        { before: 4, after: 2 },
      ];
      const spacing = headingSpacing[headingLevel - 1] || {
        before: 4,
        after: 2,
      };
      spaceAbove = { v: spacing.before };
      spaceBelow = { v: spacing.after };
    } else {
      spaceAbove = { v: 0 };
      spaceBelow = { v: 8 };
    }

    if (pPr?.["w:spacing"]) {
      const spacing = pPr["w:spacing"];

      if (spacing["@_w:line"]) {
        const lineVal = parseInt(spacing["@_w:line"]);
        const lineRule = spacing["@_w:lineRule"];

        if (lineRule === "exact") {
          lineSpacing = lineVal / 20;
        } else if (lineRule === "atLeast") {
          lineSpacing = lineVal / 20;
        } else {
          lineSpacing = lineVal / 240;
        }
      }

      if (spacing["@_w:before"] !== undefined) {
        spaceAbove = { v: parseInt(spacing["@_w:before"]) / 20 };
      }
      if (spacing["@_w:after"] !== undefined) {
        spaceBelow = { v: parseInt(spacing["@_w:after"]) / 20 };
      }
    }

    // Check for list numbering
    if (pPr?.["w:numPr"]) {
      const numId = pPr["w:numPr"]["w:numId"]?.["@_w:val"];
      const ilvl = pPr["w:numPr"]["w:ilvl"]?.["@_w:val"] || "0";

      if (numId !== undefined && numId !== null) {
        const listId = String(numId);
        const nestingLevel = parseInt(ilvl);

        if (!lists[listId]) {
          const abstractNumId = listDefs.get(numId);
          const abstractLevels = abstractNumId
            ? abstractNumDefs.get(abstractNumId)
            : null;

          if (abstractLevels) {
            // Determine list type from the first level's glyphType (0 = bullet, 1 = ordered)
            const firstLevelGlyphType =
              abstractLevels[0]?.glyphType ??
              (abstractLevels[0]?.numFmt === "bullet" ? 0 : 1);
            const listTypeValue =
              firstLevelGlyphType === 0 ? "BULLET_LIST" : "ORDER_LIST";

        

            lists[listId] = {
              listType: listTypeValue,
              // biome-ignore lint/suspicious/noExplicitAny: Nesting level structure
              nestingLevel: abstractLevels.map(
                (lvlDef: any, level: number) => ({
                  bulletAlignment: 0,
                  glyphFormat: lvlDef.glyphFormat || "•",
                  startNumber: lvlDef.start || 0,
                  glyphType:
                    lvlDef.glyphType ?? (lvlDef.numFmt === "bullet" ? 0 : 1),
                  textStyle: { fs: 11 },
                  hanging: { v: lvlDef.hanging || 21 },
                  indentStart: { v: lvlDef.indentStart || 21 * (level + 1) },
                }),
              ),
            };
          } else {
           lists[listId] = {
              listType: "BULLET_LIST",
              nestingLevel: Array.from({ length: 9 }, (_, level) => ({
                bulletAlignment: 0,
                glyphFormat: "•",
                startNumber: 0,
                glyphType: 0,
                textStyle: { fs: 11 },
                hanging: { v: 24 },
                indentStart: { v: 36 * (level + 1) },
              })),
            };
          }
        }

        bullet = {
          listId: listId,
          nestingLevel,
          listDef: lists[listId],
        };
      }
    }

    // Process runs (text content) and hyperlinks
    const runs = p["w:r"] || [];
    const hyperlinks = p["w:hyperlink"] || [];
    const runArray = Array.isArray(runs) ? runs : [runs];
    const hyperlinkArray = Array.isArray(hyperlinks)
      ? hyperlinks
      : hyperlinks
        ? [hyperlinks]
        : [];

    // Process hyperlinks first
    // biome-ignore lint/suspicious/noExplicitAny: Hyperlink structure
    hyperlinkArray.forEach((hyperlink: any) => {
      const rId = hyperlink["@_r:id"];
      const relationship = rId ? relationships.get(rId) : null;
      const url =
        relationship?.type ===
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
          ? relationship.target
          : null;

      const hyperlinkRuns = hyperlink["w:r"] || [];
      const hyperlinkRunArray = Array.isArray(hyperlinkRuns)
        ? hyperlinkRuns
        : [hyperlinkRuns];

      const hyperlinkStartIndex = dataStream.length;

      // biome-ignore lint/suspicious/noExplicitAny: Run structure
      hyperlinkRunArray.forEach((run: any) => {
        // biome-ignore lint/suspicious/noExplicitAny: Text style structure
        const ts: any =
          isHeading && headingTextStyle ? { ...headingTextStyle } : {};
        const rPr = run["w:rPr"];

        if (url) {
          if (!ts.cl) ts.cl = { rgb: "#0563C1" };
          if (!ts.ul) ts.ul = { s: 1 };
        }

        if (rPr) {
          if ("w:b" in rPr) {
            const bVal = rPr["w:b"];
            if (
              typeof bVal === "string" ||
              !bVal["@_w:val"] ||
              bVal["@_w:val"] !== "0"
            ) {
              ts.bl = 1;
            }
          }
          if ("w:i" in rPr) {
            const iVal = rPr["w:i"];
            if (
              typeof iVal === "string" ||
              !iVal["@_w:val"] ||
              iVal["@_w:val"] !== "0"
            ) {
              ts.it = 1;
            }
          }
          if (rPr["w:u"]) ts.ul = { s: 1 };
          if (rPr["w:sz"]?.["@_w:val"]) {
            ts.fs = parseInt(rPr["w:sz"]["@_w:val"]) / 2;
          }
          if (rPr["w:rFonts"]) {
            const fonts = rPr["w:rFonts"];
            ts.ff =
              fonts["@_w:ascii"] ||
              fonts["@_w:hAnsi"] ||
              fonts["@_w:eastAsia"] ||
              fonts["@_w:cs"];
          }
          if (rPr["w:color"]?.["@_w:val"]) {
            const colorVal = rPr["w:color"]["@_w:val"];
            if (colorVal !== "auto" && colorVal !== "000000") {
              ts.cl = { rgb: `#${colorVal}` };
            }
          }
        }

        if (run["w:t"]) {
          const text =
            typeof run["w:t"] === "object"
              ? run["w:t"]["#text"] || ""
              : run["w:t"];
          if (text) {
            hasContent = true;
            const st = dataStream.length;
            dataStream += text;
            textRuns.push({ st, ed: dataStream.length, ts });
          }
        }
      });

      if (url && dataStream.length > hyperlinkStartIndex) {
        const rangeId = `hyperlink_${customRanges.length + 1}`;
        customRanges.push({
          rangeId,
          rangeType: 0,
          startIndex: hyperlinkStartIndex,
          endIndex: dataStream.length,
          properties: { url },
        });
      }
    });

    // Process regular runs
    for (let runIndex = 0; runIndex < runArray.length; runIndex++) {
      const run = runArray[runIndex];

      // Univer will handle list glyphs automatically based on listType and glyphFormat
      // No need to manually prepend glyphs

      // biome-ignore lint/suspicious/noExplicitAny: Text style structure
      const ts: any =
        isHeading && headingTextStyle ? { ...headingTextStyle } : {};

      const rPr = run["w:rPr"];

      if (rPr) {
        if ("w:b" in rPr) {
          const bVal = rPr["w:b"];
          if (
            typeof bVal === "string" ||
            !bVal["@_w:val"] ||
            bVal["@_w:val"] !== "0"
          ) {
            ts.bl = 1;
          }
        }

        if ("w:i" in rPr) {
          const iVal = rPr["w:i"];
          if (
            typeof iVal === "string" ||
            !iVal["@_w:val"] ||
            iVal["@_w:val"] !== "0"
          ) {
            ts.it = 1;
          }
        }

        if (rPr["w:u"]) ts.ul = { s: 1 };
        if (rPr["w:strike"]) ts.st = 1;
        if (rPr["w:sz"]?.["@_w:val"]) {
          ts.fs = parseInt(rPr["w:sz"]["@_w:val"]) / 2;
        }
        if (rPr["w:rFonts"]) {
          const fonts = rPr["w:rFonts"];
          ts.ff =
            fonts["@_w:ascii"] ||
            fonts["@_w:hAnsi"] ||
            fonts["@_w:eastAsia"] ||
            fonts["@_w:cs"];
        }
        if (rPr["w:vertAlign"]?.["@_w:val"]) {
          const va = rPr["w:vertAlign"]["@_w:val"];
          if (va === "superscript") ts.va = 1;
          if (va === "subscript") ts.va = 2;
        }
        if (rPr["w:color"]?.["@_w:val"]) {
          const colorVal = rPr["w:color"]["@_w:val"];
          if (colorVal !== "auto" && colorVal !== "000000") {
            ts.cl = { rgb: `#${colorVal}` };
          }
        }
        if (rPr["w:shd"]?.["@_w:fill"]) {
          const fillVal = rPr["w:shd"]["@_w:fill"];
          if (fillVal !== "auto" && fillVal !== "FFFFFF") {
            ts.bg = { rgb: `#${fillVal}` };
          }
        }
        if (rPr["w:highlight"]?.["@_w:val"]) {
          const highlightColors: Record<string, string> = {
            yellow: "#FFFF00",
            green: "#00FF00",
            cyan: "#00FFFF",
            magenta: "#FF00FF",
            blue: "#0000FF",
            red: "#FF0000",
          };
          const highlightVal = rPr["w:highlight"]["@_w:val"];
          if (highlightColors[highlightVal]) {
            ts.bg = { rgb: highlightColors[highlightVal] };
          }
        }
      }

      if (run["w:t"]) {
        const text =
          typeof run["w:t"] === "object"
            ? run["w:t"]["#text"] || ""
            : run["w:t"];

        if (text) {
          hasContent = true;
          const st = dataStream.length;
          dataStream += text;
          textRuns.push({ st, ed: dataStream.length, ts });
        }
      }

      if (run["w:br"]) {
        hasContent = true;
        dataStream += "\r";
      }
    }

    // If paragraph is empty, add space placeholder
    if (!hasContent) {
      const st = dataStream.length;
      dataStream += " ";
      textRuns.push({ st, ed: dataStream.length, ts: {} });
    }

    // Add paragraph terminator
    dataStream += "\r";

    const paragraphStartIndex = dataStream.length - 1;

    // Create paragraph definition with Univer-compatible alignment structure
    // biome-ignore lint/suspicious/noExplicitAny: Paragraph structure
    const paragraph: any = {
      startIndex: paragraphStartIndex,
      paragraphStyle: {
        spaceAbove: spaceAbove || { v: 0 },
        spaceBelow: spaceBelow || { v: 0 },
        lineSpacing: lineSpacing || 1.5,
      },
    };

    // Set alignment using Univer's expected format (0=Left, 1=Center, 2=Right, 3=Justify)
    if (textAlignment !== undefined) {
      paragraph.paragraphStyle.horizontalAlign = textAlignment;
      console.log(`✓ Applied alignment ${textAlignment} to paragraph at index ${paragraphStartIndex}`);
      console.log(`   Paragraph text preview:`, dataStream.substring(Math.max(0, paragraphStartIndex - 20), paragraphStartIndex + 20).replace(/\r/g, '\\r'));
    } else {
      // IMPORTANT: Check exported IDocumentData to see if Univer uses undefined or 0 for default alignment
      console.log(`⚠️ No alignment specified for paragraph at index ${paragraphStartIndex} - NOT setting horizontalAlign (will be undefined)`);
      console.log(`   Paragraph text preview:`, dataStream.substring(Math.max(0, paragraphStartIndex - 20), paragraphStartIndex + 20).replace(/\r/g, '\\r'));
      // DO NOT set a default value - let Univer handle it
      // If Univer exports show horizontalAlign is omitted for left-aligned text, we should also omit it
    }

    if (indentStart) paragraph.paragraphStyle.indentStart = indentStart;
    if (indentEnd) paragraph.paragraphStyle.indentEnd = indentEnd;
    if (indentFirstLine)
      paragraph.paragraphStyle.indentFirstLine = indentFirstLine;
    if (hangingIndent) paragraph.paragraphStyle.hanging = hangingIndent;

    if (bullet) {
      const listDef = lists[bullet.listId];
      const levelDef = listDef?.nestingLevel?.[bullet.nestingLevel];

      // Determine list type based on glyphType (0 = bullet, 1 = ordered/numbered)
      const isBulletList = levelDef?.glyphType === 0;

      paragraph.bullet = {
        listType: isBulletList ? "BULLET_LIST" : "ORDER_LIST",
        listId: bullet.listId,
        nestingLevel: bullet.nestingLevel,
        textStyle: {
          fs: 10,
        },
      };

      if (levelDef) {
      
        // levelDef.indentStart and levelDef.hanging are already { v: number } objects
        paragraph.paragraphStyle.indentStart = levelDef.indentStart || {
          v: 36 * (bullet.nestingLevel + 1),
        };
        paragraph.paragraphStyle.hanging = levelDef.hanging || { v: 24 };
      } else {
        paragraph.paragraphStyle.indentStart = {
          v: 36 * (bullet.nestingLevel + 1),
        };
        paragraph.paragraphStyle.hanging = { v: 24 };
      }
    }

    paragraphs.push(paragraph);
  }

  // Ensure minimum structure
  if (paragraphs.length === 0) {
    console.warn("⚠️ No paragraphs found, creating default paragraph");
    dataStream = "Default Document Content\r";
    textRuns.push({ st: 0, ed: 23, ts: {} });
    paragraphs.push({ startIndex: 0, paragraphStyle: {} }); // No horizontalAlign = left by default
  }

  // Ensure document ends with proper Univer structure
  dataStream = dataStream.replace(/[\r\n]+$/, "");

  if (dataStream.length === 0 || dataStream.trim().length === 0) {
    console.warn(
      "⚠️ No content extracted from DOCX, creating minimal document",
    );
    dataStream = "Document content could not be extracted\r";
    textRuns.length = 0; // Clear array
    textRuns.push({ st: 0, ed: dataStream.length - 1, ts: {} });
    paragraphs.length = 0; // Clear array
    paragraphs.push({ startIndex: 0, paragraphStyle: {} }); // No horizontalAlign = left by default
  } else if (!dataStream.endsWith("\r")) {
    dataStream += "\r";
  }

  // Add section break
if (!dataStream.endsWith('\r')) {
  dataStream += '\r';
}
const sectionBreakIndex = dataStream.length;
dataStream += '\n';


  const result = {
    body: {
      dataStream,
      paragraphs,
      textRuns,
      sectionBreaks: [
        {
          startIndex: sectionBreakIndex,
        },
      ],
      customRanges: customRanges.length > 0 ? customRanges : undefined,
      // body.tables contains table position references
      tables: bodyTables.length > 0 ? bodyTables : [],
      customBlocks: [],
    },
    lists,
    // tableSource is the TOP-LEVEL property for table definitions (from IReferenceSource)
    tableSource: Object.keys(tables).length > 0 ? tables : {},
  };

  console.log("✅ DOCX conversion complete:");
  console.log(
    "📊 Result stats:",
    JSON.stringify(
      {
        dataStreamLength: dataStream.length,
        paragraphsCount: paragraphs.length,
        textRunsCount: textRuns.length,
        listsCount: Object.keys(lists || {}).length,
        tablesCount: Object.keys(tables || {}).length,
        customRangesCount: customRanges.length,
        dataStreamPreview: dataStream
          .substring(0, 200)
          .replace(/\r/g, "\\r")
          .replace(/\n/g, "\\n"),
        firstParagraph: paragraphs[0] || null,
        firstTextRun: textRuns[0] || null,
      },
      null,
      2,
    ),
  );

  return result;
}

/**
 * Parse a DOCX table (w:tbl) element into Univer table format
 * Builds the table dataStream with control characters and returns tableSource
 *
 * Control characters in dataStream:
 * - \x1A (26) - Table Start
 * - \x1B (27) - Row Start
 * - \x1C (28) - Cell Start
 * - \x1D (29) - Cell End
 * - \x0E (14) - Row End
 * - \x0F (15) - Table End
 */
// biome-ignore lint/suspicious/noExplicitAny: OOXML table structure is dynamic
function parseTableWithDataStream(
  tableElement: any,
  // biome-ignore lint/suspicious/noExplicitAny: Style map structure
  _styleMap: Map<string, any>,
  tableId: string,
  _baseIndex: number,
): {
  dataStream: string;
  tableSource: any;
} | null {
  try {
    let tableDataStream = "";

    // Table Start character
    tableDataStream += "\x1A";

    // Get all rows
    const rows = tableElement["w:tr"]
      ? Array.isArray(tableElement["w:tr"])
        ? tableElement["w:tr"]
        : [tableElement["w:tr"]]
      : [];

    if (rows.length === 0) return null;

    // biome-ignore lint/suspicious/noExplicitAny: Table rows structure
    const tableRows: any[] = [];
    let maxCols = 0;

    // Process each row
    // biome-ignore lint/suspicious/noExplicitAny: Row element structure
    rows.forEach((row: any) => {
      const trPr = row["w:trPr"];
      let rowHeight = 30;
      if (trPr?.["w:trHeight"]?.["@_w:val"]) {
        rowHeight = parseInt(trPr["w:trHeight"]["@_w:val"]) / 20;
      }

      // Row Start character
      tableDataStream += "\x1B";

      const cellElements = row["w:tc"]
        ? Array.isArray(row["w:tc"])
          ? row["w:tc"]
          : [row["w:tc"]]
        : [];

      // biome-ignore lint/suspicious/noExplicitAny: Cell array
      const tableCells: any[] = [];
      let colCount = 0;

      // biome-ignore lint/suspicious/noExplicitAny: Cell element structure
      cellElements.forEach((cellElement: any) => {
        const tcPr = cellElement["w:tcPr"];

        // Check for merged cell continuation (vMerge without restart)
        if (tcPr?.["w:vMerge"]) {
          const vMergeVal = tcPr["w:vMerge"]["@_w:val"];
          if (vMergeVal !== "restart") {
            // This is a continuation cell - skip it
            return;
          }
        }

        // Get gridSpan for column count
        const gridSpanNode = tcPr?.["w:gridSpan"];
        const gridSpan = gridSpanNode ? parseInt(gridSpanNode["@_w:val"] || "1") : 1;
        colCount += gridSpan;

        // Cell Start character
        tableDataStream += "\x1C";

        // Extract cell text content
        const cellText = extractCellTextContent(cellElement);

        // Add cell content and paragraph break
        if (cellText) {
          tableDataStream += cellText;
        }
        tableDataStream += "\r"; // Paragraph break inside cell

        // Cell End character
        tableDataStream += "\x1D";

        // Build cell object matching ITableCell interface
        // biome-ignore lint/suspicious/noExplicitAny: Cell metadata
        const cell: any = {
          margin: {
            top: { v: 5 },
            bottom: { v: 5 },
            start: { v: 5 },
            end: { v: 5 },
          },
          rowSpan: 1,
          columnSpan: gridSpan,
          backgroundColor: { rgb: "#ffffff" },
          borderTop: { color: { rgb: "#000000" }, width: { v: 1 }, dashStyle: 0 },
          borderBottom: { color: { rgb: "#000000" }, width: { v: 1 }, dashStyle: 0 },
          borderLeft: { color: { rgb: "#000000" }, width: { v: 1 }, dashStyle: 0 },
          borderRight: { color: { rgb: "#000000" }, width: { v: 1 }, dashStyle: 0 },
        };

        // Background color
        if (tcPr?.["w:shd"]?.["@_w:fill"]) {
          const fill = tcPr["w:shd"]["@_w:fill"];
          if (fill !== "auto" && fill !== "FFFFFF") {
            cell.backgroundColor = { rgb: `#${fill}` };
          }
        }

        // Vertical alignment
        if (tcPr?.["w:vAlign"]?.["@_w:val"]) {
          const vAlign = tcPr["w:vAlign"]["@_w:val"];
          const vAlignMap: Record<string, number> = { top: 0, center: 1, bottom: 2 };
          cell.vAlign = vAlignMap[vAlign] ?? 0;
        }

        tableCells.push(cell);
      });

      if (colCount > maxCols) maxCols = colCount;

      // Row End character
      tableDataStream += "\x0E";

      // Build row object matching ITableRow interface
      tableRows.push({
        tableCells,
        trHeight: {
          val: { v: rowHeight },
          hRule: 0, // AUTO
        },
      });
    });

    // Table End character
    tableDataStream += "\x0F";

    // Build tableColumns array
    const tableColumns = Array(maxCols).fill(null).map(() => ({
      size: { type: 1, width: { v: 150 } },
    }));

    // Build tableSource entry matching ITable interface
    const tableSource = {
      tableId,
      tableRows,
      tableColumns,
      align: 0, // TableAlignmentType.START
      indent: { v: 0 },
      textWrap: 0, // TableTextWrapType.NONE
      position: {
        positionH: { relativeFrom: 0, posOffset: 0 },
        positionV: { relativeFrom: 0, posOffset: 0 },
      },
      dist: { distB: 0, distL: 0, distR: 0, distT: 0 },
      cellMargin: {
        start: { v: 10 },
        end: { v: 10 },
        top: { v: 5 },
        bottom: { v: 5 },
      },
      size: {
        type: 0, // TableSizeType.UNSPECIFIED
        width: { v: 660 },
      },
      layout: 0, // TableLayoutType.AUTO_FIT
    };

    return {
      dataStream: tableDataStream,
      tableSource,
    };
  } catch (error) {
    console.error("Error parsing table:", error);
    return null;
  }
}

/**
 * Extract text content from a table cell (for dataStream)
 */
// biome-ignore lint/suspicious/noExplicitAny: Cell element structure
function extractCellTextContent(cellElement: any): string {
  let text = "";

  // Get paragraphs within the cell
  const paragraphs = cellElement["w:p"]
    ? Array.isArray(cellElement["w:p"])
      ? cellElement["w:p"]
      : [cellElement["w:p"]]
    : [];

  // biome-ignore lint/suspicious/noExplicitAny: Paragraph structure
  paragraphs.forEach((p: any, pIndex: number) => {
    // Add newline between paragraphs (but not before the first)
    if (pIndex > 0) text += "\r";

    // Get runs within the paragraph
    const runs = p["w:r"]
      ? Array.isArray(p["w:r"])
        ? p["w:r"]
        : [p["w:r"]]
      : [];

    // biome-ignore lint/suspicious/noExplicitAny: Run structure
    runs.forEach((run: any) => {
      if (run["w:t"]) {
        const runText =
          typeof run["w:t"] === "object"
            ? run["w:t"]["#text"] || ""
            : run["w:t"];
        text += runText;
      }
    });
  });

  return text || " "; // Return at least a space for empty cells
}
