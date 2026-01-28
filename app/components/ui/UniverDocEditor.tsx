// Import styles for preset mode
import "@univerjs/preset-docs-core/lib/index.css";
import { Download, Upload } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import { CustomQuickInsertPlugin } from "~/plugins/CustomQuickInsertPlugin";
import { HorizontalLineSpacingPlugin } from "~/plugins/HorizontalLineSpacingPlugin";
import { convertDocxToUniverData } from "~/utils/docx-converter";
import CustomQuickInsertMenu from "../CustomQuickInsertMenu";


interface UniverDocEditorProps {
  initialFile?: File;
}

export function UniverDocEditor({ initialFile }: UniverDocEditorProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [importing, setImporting] = useState(false);
  // biome-ignore lint/suspicious/noExplicitAny: IDocumentData type from converter
  const [documentData, setDocumentData] = useState<any>(null);
  const [fileName, setFileName] = useState<string>("");
  const timeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const lifecycleDisposeRef = useRef<{ dispose: () => void } | null>(null);
  // biome-ignore lint/suspicious/noExplicitAny: Univer API type is complex
  const univerAPIRef = useRef<any>(null);
  // biome-ignore lint/suspicious/noExplicitAny: Univer instance type is complex
  const univerInstanceRef = useRef<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Quick insert menu state
  const [menuVisible, setMenuVisible] = useState(false);
  const [menuPosition, setMenuPosition] = useState<
    { x: number; y: number } | undefined
  >();
  const quickInsertServiceRef = useRef<any>(null);

  // Load initial file if provided
  useEffect(() => {
    if (initialFile) {
      console.log(`📂 Loading initial file: ${initialFile.name}`);
      setFileName(initialFile.name);
      setImporting(true);
      setLoading(true);

      initialFile.arrayBuffer().then(async (arrayBuffer) => {
        try {
          const convertedData = await convertDocxToUniverData(arrayBuffer);
          console.log("✅ Initial file converted successfully");
          setDocumentData(convertedData);
        } catch (err) {
          console.error("❌ Failed to import initial file:", err);
          setError(
            err instanceof Error ? err.message : "Failed to import document",
          );
          setLoading(false);
          setImporting(false);
        }
      });
    }
  }, [initialFile]);

  // biome-ignore lint/correctness/useExhaustiveDependencies: documentData accessed in nested function
  useEffect(() => {
    if (!containerRef.current) return;

    // Dynamically import Univer packages using PRESET MODE
    const initUniver = async () => {
      try {
        const [presetsModule, presetDocsCore, presetLocale] = await Promise.all(
          [
            import("@univerjs/presets"),
            import("@univerjs/preset-docs-core"),
            import("@univerjs/preset-docs-core/locales/en-US"),
          ],
        );

        const { createUniver, LocaleType } = presetsModule;
        const { UniverDocsCorePreset } = presetDocsCore;
        const UniverPresetDocsCoreEnUS = presetLocale.default || presetLocale;

        if (!containerRef.current) {
          throw new Error("Container ref is not available");
        }

        // Initialize Univer using PRESET MODE
        const { univerAPI, univer } = createUniver({
          locale: LocaleType.EN_US,
          locales: {
            [LocaleType.EN_US]: UniverPresetDocsCoreEnUS,
          },
          presets: [
            UniverDocsCorePreset({
              container: containerRef.current,
            }),
          ],
        });

        // Register the CustomQuickInsertPlugin after Univer is created
        univer.registerPlugin(CustomQuickInsertPlugin);
        console.log("✅ CustomQuickInsertPlugin registered successfully");

        // Register the HorizontalLineSpacingPlugin for improved horizontal line spacing
        univer.registerPlugin(HorizontalLineSpacingPlugin);
        console.log("✅ HorizontalLineSpacingPlugin registered successfully");

        univerAPIRef.current = univerAPI;
        univerInstanceRef.current = univer;

        // Setup quick insert menu event listeners
        const handleQuickInsertShow = (event: CustomEvent) => {
          console.log(
            "[UniverDocEditor] Quick insert show event:",
            event.detail,
          );

          // Get service if not already available
          if (!quickInsertServiceRef.current) {
            quickInsertServiceRef.current = (
              window as any
            ).__customQuickInsertService;
          }

          const { bounds } = event.detail;
          if (bounds) {
            const univerContainer = containerRef.current;

            if (univerContainer) {
              const canvasElement = univerContainer.querySelector(
                "canvas"
              ) as HTMLCanvasElement | null;

              if (canvasElement) {
                const canvasRect = canvasElement.getBoundingClientRect();

                // Find ANY scrollable parent element
                let scrollTop = 0;
                let scrollLeft = 0;
                let scrollableElement: HTMLElement | null = null;

                // Check all parent elements for scroll
                let element: HTMLElement | null = canvasElement.parentElement;
                while (element && element !== document.body) {
                  const hasVerticalScroll = element.scrollHeight > element.clientHeight;
                  const hasHorizontalScroll = element.scrollWidth > element.clientWidth;
                  const computedStyle = window.getComputedStyle(element);
                  const overflowY = computedStyle.overflowY;
                  const overflowX = computedStyle.overflowX;

                  if (
                    (hasVerticalScroll && (overflowY === 'auto' || overflowY === 'scroll')) ||
                    (hasHorizontalScroll && (overflowX === 'auto' || overflowX === 'scroll'))
                  ) {
                    scrollableElement = element;
                    scrollTop = element.scrollTop;
                    scrollLeft = element.scrollLeft;
                    console.log("[UniverDocEditor] Found scrollable container:", {
                      element: element.className,
                      scrollTop,
                      scrollLeft,
                      scrollHeight: element.scrollHeight,
                      clientHeight: element.clientHeight
                    });
                    break;
                  }
                  element = element.parentElement;
                }

                console.log("[UniverDocEditor] Canvas rect:", canvasRect);
                console.log("[UniverDocEditor] Raw bounds:", bounds);
                console.log("[UniverDocEditor] Final scroll offset:", { scrollTop, scrollLeft });

                // Calculate viewport coordinates (bounds are in document space, convert to viewport)
                // bounds.bottom is absolute document position, subtract scroll to get viewport position
                const viewportX = bounds.left + canvasRect.left - scrollLeft;
                const viewportY = bounds.bottom + canvasRect.top - scrollTop + 5;

                console.log(
                  "[UniverDocEditor] Viewport position:",
                  viewportX,
                  viewportY
                );

                setMenuPosition({ x: viewportX, y: viewportY });
                setMenuVisible(true);
              } else {
                console.warn("[UniverDocEditor] Canvas element not found");
                setMenuPosition({ x: bounds.left, y: bounds.bottom + 5 });
                setMenuVisible(true);
              }
            }
          }
        };

        const handleQuickInsertClose = () => {
          console.log("[UniverDocEditor] Quick insert close event");
          setMenuVisible(false);
        };

        window.addEventListener(
          "univer:quick-insert-show",
          handleQuickInsertShow as EventListener,
        );
        window.addEventListener(
          "univer:quick-insert-close",
          handleQuickInsertClose as EventListener,
        );

        const createDocument = () => {
          try {
            setError(null);

            if (documentData) {
              console.log("📄 Creating document with imported data");
              console.log(
                "🔍 Document data structure:",
                JSON.stringify(
                  {
                    hasBody: !!documentData.body,
                    dataStreamLength:
                      documentData.body?.dataStream?.length || 0,
                    paragraphsCount: documentData.body?.paragraphs?.length || 0,
                    textRunsCount: documentData.body?.textRuns?.length || 0,
                    dataStreamPreview:
                      documentData.body?.dataStream?.substring(0, 100) ||
                      "empty",
                    firstParagraph: documentData.body?.paragraphs?.[0] || null,
                    firstTextRun: documentData.body?.textRuns?.[0] || null,
                  },
                  null,
                  2,
                ),
              );

              // Ensure the document data has valid content
              if (
                !documentData.body ||
                !documentData.body.dataStream ||
                documentData.body.dataStream.length === 0
              ) {
                console.warn(
                  "⚠️ Document data appears empty, creating with sample content",
                );
                univerAPI.createUniverDoc({
                  body: {
                    dataStream: "Sample Document Content\r\n",
                    paragraphs: [
                      { startIndex: 0, paragraphStyle: { horizontalAlign: 0 } },
                      {
                        startIndex: 24,
                        paragraphStyle: { horizontalAlign: 0 },
                      },
                    ],
                    textRuns: [{ st: 0, ed: 23, ts: {} }],
                    sectionBreaks: [{ startIndex: 24 }],
                  },
                });
              } else {
                univerAPI.createUniverDoc(documentData);
              }
            } else {
              console.log("📄 Creating blank document");
              univerAPI.createUniverDoc({});
            }

            setTimeout(() => {
              const activeDoc = univerAPI.getActiveDocument();
              if (activeDoc) {
                console.log("✓ Document created:", activeDoc.getId());
                const snapshot = activeDoc.getSnapshot();
                console.log(
                  "📋 Document snapshot:",
                  JSON.stringify(
                    {
                      hasBody: !!snapshot?.body,
                      dataStreamLength: snapshot?.body?.dataStream?.length || 0,
                      paragraphsCount: snapshot?.body?.paragraphs?.length || 0,
                      dataStreamContent:
                        snapshot?.body?.dataStream?.substring(0, 200) ||
                        "empty",
                      firstParagraph: snapshot?.body?.paragraphs?.[0] || null,
                    },
                    null,
                    2,
                  ),
                );
              }
            }, 100);

            setLoading(false);
            setImporting(false);
          } catch (e) {
            console.error("Failed to create document:", e);
            setError("Failed to create document");
            setLoading(false);
            setImporting(false);
          }
        };

        let contentCreated = false;

        const createContentOnce = () => {
          if (!contentCreated) {
            contentCreated = true;
            createDocument();
          }
        };

        // Force-apply alignment after document creation using Univer commands
        // This ensures alignment triggers proper layout recalculation and render invalidation
        const applyAlignmentFix = async () => {
          const api = univerAPIRef.current;
          if (!api) return;

          const doc = api.getActiveDocument();
          if (!doc) return;

          const snapshot = doc.getSnapshot();
          const paragraphs = snapshot.body?.paragraphs ?? [];

          console.log(`🔧 Applying alignment fix to ${paragraphs.length} paragraphs`);

          try {
            // Alignment is already applied during import in docx-converter.ts
            // No need to re-apply alignment via SetParagraphAlignCommand
            // (SetParagraphAlignCommand doesn't exist in @univerjs/docs v0.15.x)

            paragraphs.forEach((p: any, index: number) => {
              const alignment = p.paragraphStyle?.horizontalAlign;
              if (alignment !== undefined && alignment !== 0) {
                // Alignment is already in the data from import
                console.log(`  ✓ Paragraph ${index} at ${p.startIndex}: has alignment ${alignment}`);
              }
            });

            console.log("✅ Alignment check completed");
          } catch (err) {
            console.error("❌ Error checking alignment:", err);
          }
        };

        const disposable = univerAPI.addEvent(
          univerAPI.Event.LifeCycleChanged,
          // biome-ignore lint/suspicious/noExplicitAny: Univer lifecycle event structure
          ({ stage }: any) => {
            console.log("Univer lifecycle stage:", stage);

            if (stage === univerAPI.Enum.LifecycleStages.Rendered) {
              createContentOnce();

              // Apply alignment fix after document is rendered
              setTimeout(() => {
                applyAlignmentFix();
              }, 100);

              setTimeout(() => {
                const activeDoc = univerAPI.getActiveDocument();
                if (activeDoc) {
                  console.log(
                    "Active document at Rendered:",
                    activeDoc.getId(),
                  );
                }

                // Focus the editor
                const focusEditor = () => {
                  const selectors = [
                    '[contenteditable="true"]',
                    ".univer-editor",
                    '[role="textbox"]',
                  ];

                  for (const selector of selectors) {
                    const editorElement = containerRef.current?.querySelector(
                      selector,
                    ) as HTMLElement;
                    if (editorElement) {
                      editorElement.focus();
                      editorElement.click();
                      console.log("Editor focused using selector:", selector);
                      return true;
                    }
                  }
                  return false;
                };

                if (!focusEditor()) {
                  setTimeout(() => {
                    focusEditor();
                  }, 200);
                }
              }, 300);
            }

            if (stage === univerAPI.Enum.LifecycleStages.Steady) {
              if (disposable) {
                disposable.dispose();
                lifecycleDisposeRef.current = null;
              }
            }
          },
        );
        lifecycleDisposeRef.current = disposable || null;

        timeoutRef.current = setTimeout(() => {
          if (!contentCreated) {
            console.log("Creating document via fallback timeout");
            createContentOnce();
          }
          if (lifecycleDisposeRef.current) {
            lifecycleDisposeRef.current.dispose();
            lifecycleDisposeRef.current = null;
          }
        }, 1500);
      } catch (e) {
        console.error("Failed to initialize Univer:", e);
        setError("Failed to initialize editor");
        setLoading(false);
      }
    };

    initUniver();

    return () => {
      if (timeoutRef.current) {
        clearTimeout(timeoutRef.current);
      }
      if (lifecycleDisposeRef.current) {
        lifecycleDisposeRef.current.dispose();
        lifecycleDisposeRef.current = null;
      }
      if (univerInstanceRef.current) {
        univerInstanceRef.current.dispose();
        univerInstanceRef.current = null;
        univerAPIRef.current = null;
      }
      // Cleanup event listeners
      window.removeEventListener("univer:quick-insert-show", () => { });
      window.removeEventListener("univer:quick-insert-close", () => { });
    };
  }, [documentData]);

  // Handle DOCX file import
  const handleFileImport = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.name.endsWith(".docx") && !file.name.endsWith(".doc")) {
      setError("Please select a valid DOCX file");
      return;
    }

    setImporting(true);
    setLoading(true);
    setError(null);
    setFileName(file.name);

    try {
      console.log(`📂 Importing DOCX file: ${file.name}`);

      const arrayBuffer = await file.arrayBuffer();
      const convertedData = await convertDocxToUniverData(arrayBuffer);

      console.log("✅ DOCX converted successfully");
      setDocumentData(convertedData);
    } catch (err) {
      console.error("❌ Failed to import DOCX:", err);
      setError(
        err instanceof Error ? err.message : "Failed to import document",
      );
      setLoading(false);
      setImporting(false);
    } finally {
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  const triggerFileImport = () => {
    fileInputRef.current?.click();
  };

  const handleExport = async () => {
    try {
      console.log("[DOCX Export] Starting SERVER-SIDE export...");

      if (!univerAPIRef.current) {
        setError("Editor not initialized");
        return;
      }

      const activeDoc = univerAPIRef.current.getActiveDocument();
      if (!activeDoc) {
        setError("No active document");
        return;
      }

      // CRITICAL FIX: Force document to commit any pending edits before snapshot
      // This ensures getSnapshot() includes the very latest user changes
      console.log("[DOCX Export] Flushing pending edits...");

      // Trigger a blur event to force Univer to commit any buffered text
      const editorElement = document.querySelector('.univer-render-canvas');
      if (editorElement) {
        (editorElement as HTMLElement).blur();
        // Small delay to let Univer process the blur and commit changes
        await new Promise(resolve => setTimeout(resolve, 100));
      }

      const snapshot = activeDoc.getSnapshot();

      // Log full IDocumentData contents
      console.log("[DOCX Export] ========== FULL IDocumentData Contents ==========");
      console.log("[DOCX Export] Complete snapshot object:", snapshot);
      console.log("[DOCX Export] Pretty printed:", JSON.stringify(snapshot, null, 2));

      console.log("[DOCX Export] Quick summary:", {
        textRuns: snapshot.body?.textRuns?.length || 0,
        dataStreamLength: snapshot.body?.dataStream?.length || 0,
        paragraphs: snapshot.body?.paragraphs?.length || 0,
        dataStreamPreview: snapshot.body?.dataStream?.substring(0, 200) || "empty",
        dataStreamEndPreview: snapshot.body?.dataStream?.substring(Math.max(0, (snapshot.body?.dataStream?.length || 0) - 50)) || "empty",
      });
      console.log("[DOCX Export] ===================================================");

      // Prepare filename
      const exportFileName = fileName
        ? `${fileName.replace(/\.[^/.]+$/, "")}_edited.docx`
        : "document_edited.docx";

      console.log("[DOCX Export] Sending to server for processing...");

      // Send document data to server for DOCX generation
      const response = await fetch("/api/export", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          documentData: snapshot,
          filename: exportFileName,
        }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(
          errorData.message ||
          `Server error: ${response.status} ${response.statusText}`
        );
      }

      console.log("[DOCX Export] Receiving file from server...");

      // Get the blob from server response
      const blob = await response.blob();

      console.log("[DOCX Export] File received, size:", blob.size, "bytes");

      // Download using file-saver
      const FileSaver = await import("file-saver");
      FileSaver.saveAs(blob, exportFileName);

      console.log(
        "[DOCX Export] ✓ Document exported successfully via SERVER as",
        exportFileName
      );
    } catch (err) {
      console.error("[DOCX Export] Failed:", err);
      setError(
        "Failed to export document: " +
        (err instanceof Error ? err.message : String(err))
      );
    }
  };

  const handleMenuSelect = (menu: any) => {
    console.log("[UniverDocEditor] Menu selected:", menu);

    // Get service if not already available
    const service =
      quickInsertServiceRef.current ||
      (window as any).__customQuickInsertService;

    if (service) {
      quickInsertServiceRef.current = service;
      service.emitMenuSelected(menu);
    } else {
      console.warn(
        "[UniverDocEditor] Service not available for menu selection",
      );
    }

    setMenuVisible(false);
  };

  const handleMenuClose = () => {
    console.log("[UniverDocEditor] Menu closed by user");

    // Get service if not already available
    const service =
      quickInsertServiceRef.current ||
      (window as any).__customQuickInsertService;

    if (service) {
      quickInsertServiceRef.current = service;
      service.closePopup();
    }

    setMenuVisible(false);
  };

  return (
    <div className="relative flex h-full flex-col">
      <input
        ref={fileInputRef}
        type="file"
        accept=".docx,.doc"
        onChange={handleFileImport}
        className="hidden"
      />

      {/* Quick Insert Menu */}
      {menuVisible && (
        <CustomQuickInsertMenu
          service={
            quickInsertServiceRef.current ||
            (window as any).__customQuickInsertService
          }
          visible={menuVisible}
          position={menuPosition}
          onSelect={handleMenuSelect}
          onClose={handleMenuClose}
        />
      )}

      {/* Action buttons */}
      {!error && !loading && (
        <div className="absolute top-4 right-4 z-20 flex gap-3">
          <button
            type="button"
            onClick={triggerFileImport}
            disabled={importing}
            className="flex items-center gap-2 rounded-md bg-blue-600 px-4 py-2 text-sm font-medium text-white shadow-md transition-colors hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <Upload className="h-4 w-4" />
            {importing ? "Importing..." : "Import DOCX"}
          </button>
          <button
            type="button"
            onClick={handleExport}
            className="flex items-center gap-2 rounded-md bg-emerald-600 px-4 py-2 text-sm font-medium text-white shadow-md transition-colors hover:bg-emerald-700"
          >
            <Download className="h-4 w-4" />
            Export
          </button>
        </div>
      )}

      {error ? (
        <div className="flex h-full flex-col items-center justify-center p-4">
          <h3 className="mb-2 font-semibold text-lg">
            {error.includes("import")
              ? "Import Failed"
              : "Unable to initialize editor"}
          </h3>
          <p className="mb-4 text-center text-muted-foreground text-sm">
            {error}
          </p>
          {error.includes("import") && (
            <button
              type="button"
              onClick={() => setError(null)}
              className="rounded-md bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
            >
              Try Again
            </button>
          )}
        </div>
      ) : (
        <>
          {(loading || importing) && (
            <div className="absolute inset-0 z-10 flex items-center justify-center bg-background/50">
              <div className="text-center">
                <div className="mx-auto mb-2 h-8 w-8 animate-spin rounded-full border-primary border-b-2" />
                <p className="text-muted-foreground text-sm">
                  {importing
                    ? "Importing document..."
                    : "Initializing editor..."}
                </p>
              </div>
            </div>
          )}
          <div
            key={documentData?.id || "blank"}
            className="flex-1"
            ref={containerRef}
            style={{ overflow: "hidden", position: "relative" }}
          />
        </>
      )}
    </div>
  );
}
