/**
 * Server-Side DOCX Export API Route (React Router v7)
 *
 * This endpoint handles DOCX document generation on the server side.
 * It receives document data (IDocumentData) from the client and uses
 * the docx-builder to generate a proper DOCX file with full fidelity.
 *
 * Benefits of server-side export:
 * - Captures user edits properly
 * - Better handling of complex formatting
 * - No browser memory limitations
 * - Consistent export quality
 */

import { buildDocx } from "~/utils/docx-builder";

export async function action({ request }: { request: Request }) {
  try {
    console.log("[Server Export] Received export request");

    // Parse the document data from request body
    const body = await request.json();
    const { documentData, filename } = body;
    console.log("🚀 ~ action ~ documentData:", JSON.stringify(documentData));
console.log("🚀 ~ action ~ documentData textRuns:", JSON.stringify(documentData.body?.textRuns));
console.log("🚀 ~ action ~ documentData paragraphs:", JSON.stringify(documentData.body?.paragraphs));
console.log("🚀 ~ action ~ documentData tables:", JSON.stringify(documentData.body?.tables));
console.log("🚀 ~ action ~ documentData drawings:", JSON.stringify(documentData.drawings));
console.log("🚀 ~ action ~ documentData dataStream:", documentData.body?.dataStream ? documentData.body.dataStream.substring(0, 100) + '...' : 'empty');
    if (!documentData) {
      return Response.json({ error: "Missing document data" }, { status: 400 });
    }

    console.log("[Server Export] Building DOCX for document:", documentData.id);
    console.log("[Server Export] Document stats:", {
      paragraphs: documentData.body?.paragraphs?.length || 0,
      textRuns: documentData.body?.textRuns?.length || 0,
      tables: documentData.body?.tables
        ? Object.keys(documentData.body.tables).length
        : 0,
      drawings: documentData.drawings
        ? Object.keys(documentData.drawings).length
        : 0,
    });
    console.log("[Server Export] DataStream preview:", {
      length: documentData.body?.dataStream?.length || 0,
      first200: documentData.body?.dataStream?.substring(0, 200) || "empty",
      last50: documentData.body?.dataStream?.substring(Math.max(0, (documentData.body?.dataStream?.length || 0) - 50)) || "empty",
      charCodes: documentData.body?.dataStream?.split('').map((c: string, i: number) => 
        i < 100 ? `${c}(${c.charCodeAt(0).toString(16)})` : null
      ).filter(Boolean).join(' ') || "empty"
    });

    // Build the DOCX file using our custom builder
    const docxBlob = await buildDocx(documentData);

    // Convert Blob to Buffer for response
    const arrayBuffer = await docxBlob.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    console.log(
      "[Server Export] DOCX generated, size:",
      buffer.length,
      "bytes"
    );

    // Prepare filename
    const exportFilename =
      filename || `document_${new Date().toISOString().split("T")[0]}.docx`;

    // Return the file as a download
    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${exportFilename}"`,
        "Content-Length": buffer.length.toString(),
      },
    });
  } catch (error) {
    console.error("[Server Export] Error:", error);
    return Response.json(
      {
        error: "Export failed",
        message: error instanceof Error ? error.message : "Unknown error",
      },
      { status: 500 }
    );
  }
}

// Optional: Handle OPTIONS for CORS if needed
export async function loader() {
  return new Response(null, {
    status: 405,
    headers: {
      Allow: "POST",
    },
  });
}
