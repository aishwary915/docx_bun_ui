/**
 * Example: Univer Sheets component with XLSX export
 * Shows how to integrate client-side XLSX export with formatting preservation
 */

import { useEffect, useRef, useState } from "react";
import { Button } from "~/components/ui/button";
import { Download } from "lucide-react";
import { useDocumentExport } from "~/hooks/useDocumentExport";

export function UniverSheetsExample() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerAPIRef = useRef<any>(null);
  const [loading, setLoading] = useState(true);

  const { exportXlsx, isExporting } = useDocumentExport({
    filename: "spreadsheet.xlsx",
    onSuccess: () => console.log("✓ Export complete"),
    onError: (error) => console.error("✗ Export failed:", error),
  });

  useEffect(() => {
    if (!containerRef.current) return;

    const initUniver = async () => {
      try {
        // Import Univer Sheets packages
        const [presetsModule, presetSheetsCore, presetLocale] =
          await Promise.all([
            import("@univerjs/presets"),
            import("@univerjs/preset-sheets-core"),
            import("@univerjs/preset-sheets-core/locales/en-US"),
          ]);

        const { createUniver, LocaleType } = presetsModule;
        const { UniverSheetsCorePreset } = presetSheetsCore;
        const locale = presetLocale.default || presetLocale;

        if (!containerRef.current) return;

        // Initialize Univer Sheets
        const { univerAPI, univer } = createUniver({
          locale: LocaleType.EN_US,
          locales: {
            [LocaleType.EN_US]: locale,
          },
          presets: [
            UniverSheetsCorePreset({
              container: containerRef.current,
            }),
          ],
        });

        // Create a sample workbook
        univerAPI.createWorkbook({
          name: "Sample Spreadsheet",
          sheetOrder: ["sheet-1"],
          sheets: {
            "sheet-1": {
              id: "sheet-1",
              name: "Sheet1",
              cellData: {
                0: {
                  0: { v: "Product", s: "header" },
                  1: { v: "Price", s: "header" },
                  2: { v: "Quantity", s: "header" },
                  3: { v: "Total", s: "header" },
                },
                1: {
                  0: { v: "Apple" },
                  1: { v: 1.5 },
                  2: { v: 10 },
                  3: { f: "=B2*C2" }, // Formula example
                },
                2: {
                  0: { v: "Banana" },
                  1: { v: 0.8 },
                  2: { v: 15 },
                  3: { f: "=B3*C3" },
                },
              },
            },
          },
          styles: {
            header: {
              bl: 1, // Bold
              bg: { rgb: "#4A90E2" }, // Blue background
              cl: { rgb: "#FFFFFF" }, // White text
            },
          },
        });

        univerAPIRef.current = univerAPI;
        setLoading(false);
      } catch (error) {
        console.error("Failed to initialize Univer:", error);
        setLoading(false);
      }
    };

    initUniver();

    return () => {
      // Cleanup
      if (univerAPIRef.current) {
        univerAPIRef.current.dispose?.();
      }
    };
  }, []);

  const handleExport = async () => {
    if (!univerAPIRef.current) return;

    const workbook = univerAPIRef.current.getActiveWorkbook();
    if (!workbook) return;

    // Get snapshot with all formatting
    const snapshot = workbook.getSnapshot();

    // Export to XLSX
    await exportXlsx(snapshot);
  };

  return (
    <div className="flex h-screen flex-col">
      <div className="flex items-center justify-between border-b p-4">
        <h1 className="text-xl font-bold">Spreadsheet Editor</h1>
        <Button
          onClick={handleExport}
          disabled={loading || isExporting}
          variant="default"
        >
          <Download className="mr-2 h-4 w-4" />
          {isExporting ? "Exporting..." : "Export XLSX"}
        </Button>
      </div>
      <div
        ref={containerRef}
        className="flex-1"
        style={{ width: "100%", height: "calc(100vh - 80px)" }}
      />
    </div>
  );
}

/**
 * Usage in a route:
 *
 * // app/routes/spreadsheet.tsx
 * export default function SpreadsheetRoute() {
 *   return <UniverSheetsExample />;
 * }
 */
