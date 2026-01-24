/**
 * DOCX Exporter for Univer Documents
 * Converts Univer document data back to DOCX format
 */

import {
  Document,
  Paragraph,
  TextRun,
  AlignmentType,
  HeadingLevel,
  convertInchesToTwip,
  Packer,
} from "docx";
import FileSaver from "file-saver";

const { saveAs } = FileSaver;

/**
 * Export Univer document snapshot to DOCX format
 */
export async function exportUniverToDocx(
  snapshot: any,
  filename: string = "document.docx"
): Promise<void> {
  console.log("📤 Starting DOCX export...");
  console.log("📋 Snapshot:", snapshot);

  try {
    const body = snapshot.body || snapshot;
    const dataStream = body.dataStream || "";
    const paragraphs = body.paragraphs || [];
    const textRuns = body.textRuns || [];

    console.log(`📝 Document stats:`, {
      dataStreamLength: dataStream.length,
      paragraphsCount: paragraphs.length,
      textRunsCount: textRuns.length,
    });

    // Convert Univer paragraphs to docx paragraphs
    const docxParagraphs: Paragraph[] = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      const startIndex = para.startIndex;
      const nextIndex =
        i < paragraphs.length - 1 ? paragraphs[i + 1].startIndex : dataStream.length;

      // Get paragraph text
      const paragraphText = dataStream.substring(startIndex, nextIndex).replace(/\r$/, "");

      // Find text runs for this paragraph
      const paraTextRuns = textRuns.filter(
        (run: any) => run.st >= startIndex && run.ed <= nextIndex
      );

      // Get paragraph style
      const style = para.paragraphStyle || {};
      const alignment = getAlignmentType(style.horizontalAlign || style.align);

      // Create text runs with formatting
      const runs: TextRun[] = [];

      if (paraTextRuns.length > 0) {
        for (const run of paraTextRuns) {
          const runText = dataStream.substring(run.st, run.ed);
          const textStyle = run.ts || {};

          runs.push(
            new TextRun({
              text: runText,
              bold: textStyle.bl === 1,
              italics: textStyle.it === 1,
              underline: textStyle.ul ? { type: "single" } : undefined,
              strike: textStyle.st === 1,
              size: textStyle.fs ? textStyle.fs * 2 : undefined, // Convert pt to half-pt
              color: textStyle.cl?.rgb?.replace("#", "") || undefined,
              font: textStyle.ff || undefined,
            })
          );
        }
      } else {
        // No specific text runs, use plain text
        if (paragraphText) {
          runs.push(new TextRun({ text: paragraphText }));
        }
      }

      // Determine if this is a heading
      let heading: typeof HeadingLevel[keyof typeof HeadingLevel] | undefined;
      if (style.headingLevel) {
        heading = (HeadingLevel as any)[`HEADING_${style.headingLevel}`];
      }

      // Create paragraph
      docxParagraphs.push(
        new Paragraph({
          children: runs.length > 0 ? runs : [new TextRun({ text: "" })],
          heading: heading,
          alignment: alignment,
          spacing: {
            before: style.spaceAbove?.v
              ? convertInchesToTwip(style.spaceAbove.v / 72)
              : undefined,
            after: style.spaceBelow?.v
              ? convertInchesToTwip(style.spaceBelow.v / 72)
              : undefined,
            line: style.lineSpacing
              ? style.lineSpacing * 240
              : undefined,
          },
          indent: {
            start: style.indentStart?.v
              ? convertInchesToTwip(style.indentStart.v / 72)
              : undefined,
            end: style.indentEnd?.v
              ? convertInchesToTwip(style.indentEnd.v / 72)
              : undefined,
            firstLine: style.indentFirstLine?.v
              ? convertInchesToTwip(style.indentFirstLine.v / 72)
              : undefined,
            hanging: style.hanging?.v
              ? convertInchesToTwip(style.hanging.v / 72)
              : undefined,
          },
          bullet: para.bullet
            ? {
                level: para.bullet.nestingLevel || 0,
              }
            : undefined,
        })
      );
    }

    // Create document
    const doc = new Document({
      sections: [
        {
          children: docxParagraphs.length > 0 ? docxParagraphs : [new Paragraph({})],
        },
      ],
    });

    // Generate DOCX file
    console.log("🔨 Generating DOCX file...");
    const blob = await Packer.toBlob(doc);

    // Download file
    saveAs(blob, filename);
    console.log("✅ DOCX export complete!");
  } catch (error) {
    console.error("❌ DOCX export failed:", error);
    throw error;
  }
}

/**
 * Convert Univer alignment value to docx AlignmentType
 */
function getAlignmentType(align: number | undefined): typeof AlignmentType[keyof typeof AlignmentType] | undefined {
  if (align === undefined) return undefined;

  const alignMap: Record<number, typeof AlignmentType[keyof typeof AlignmentType]> = {
    0: AlignmentType.LEFT,
    1: AlignmentType.CENTER,
    2: AlignmentType.RIGHT,
    3: AlignmentType.JUSTIFIED,
  };

  return alignMap[align];
}
