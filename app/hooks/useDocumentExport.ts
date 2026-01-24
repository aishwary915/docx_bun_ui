/**
 * Hook for exporting Univer documents with full formatting preservation
 * Supports both DOCX (server-side) and XLSX (client-side) exports
 */

import { useState } from "react";
import { exportWorkbookToXlsx } from "~/utils/xlsx-exporter";

interface UseDocumentExportOptions {
  filename?: string;
  onSuccess?: () => void;
  onError?: (error: Error) => void;
}

export function useDocumentExport(options: UseDocumentExportOptions = {}) {
  const [isExporting, setIsExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

  /**
   * Export a Univer document to DOCX format (server-side)
   */
  const exportDocx = async (documentData: any) => {
    try {
      setIsExporting(true);
      setExportError(null);

      console.log("[Export Hook] Starting DOCX export...");

      const filename = options.filename || "document.docx";

      const response = await fetch("/api/export", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          documentData,
          filename,
        }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(
          errorData.message ||
            `Server error: ${response.status} ${response.statusText}`
        );
      }

      const blob = await response.blob();

      // Download the file
      const FileSaver = await import("file-saver");
      FileSaver.saveAs(blob, filename);

      console.log("[Export Hook] ✓ DOCX export successful");
      options.onSuccess?.();
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : "Export failed";
      console.error("[Export Hook] DOCX export failed:", error);
      setExportError(errorMessage);
      options.onError?.(
        error instanceof Error ? error : new Error(errorMessage)
      );
    } finally {
      setIsExporting(false);
    }
  };

  /**
   * Export a Univer workbook to XLSX format (client-side)
   */
  const exportXlsx = async (workbookData: any) => {
    try {
      setIsExporting(true);
      setExportError(null);

      console.log("[Export Hook] Starting XLSX export...");

      const filename = options.filename || "spreadsheet.xlsx";
      await exportWorkbookToXlsx(workbookData, filename);

      console.log("[Export Hook] ✓ XLSX export successful");
      options.onSuccess?.();
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : "Export failed";
      console.error("[Export Hook] XLSX export failed:", error);
      setExportError(errorMessage);
      options.onError?.(
        error instanceof Error ? error : new Error(errorMessage)
      );
    } finally {
      setIsExporting(false);
    }
  };

  return {
    exportDocx,
    exportXlsx,
    isExporting,
    exportError,
  };
}

/**
 * Example usage in a component:
 *
 * const { exportDocx, isExporting } = useDocumentExport({
 *   filename: "my-document.docx",
 *   onSuccess: () => toast.success("Exported successfully!"),
 *   onError: (error) => toast.error(error.message),
 * });
 *
 * const handleExport = () => {
 *   const snapshot = univerAPI.getActiveDocument().getSnapshot();
 *   exportDocx(snapshot);
 * };
 */
