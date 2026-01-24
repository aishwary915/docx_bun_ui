# Document Export System

This project implements high-fidelity document export for both DOCX and XLSX formats using Univer document models.

## Architecture

### DOCX Export (Server-Side)

**Why Server-Side?**
- No browser memory limitations for large documents
- Better handling of complex OOXML structures
- Consistent export quality across all clients
- Captures all user edits properly

**Flow:**
1. Client captures Univer document snapshot
2. Sends snapshot JSON to `/export-docx` API route
3. Server uses `docx-builder.ts` to generate OOXML XML
4. Server creates DOCX ZIP file with full formatting
5. Returns file as download to client

**Files:**
- `app/routes/export-docx.ts` - API route handler
- `app/utils/docx-builder.ts` - OOXML XML generator
- `app/components/ui/UniverDocEditor.tsx` - Client integration

### XLSX Export (Client-Side)

**Why Client-Side?**
- Spreadsheets are typically smaller than documents
- Real-time export without server round-trip
- Better for frequent exports during editing
- Uses browser's native ZIP capabilities

**Flow:**
1. Captures Univer workbook snapshot
2. Builds OOXML XML structures client-side
3. Creates XLSX ZIP using JSZip
4. Downloads using file-saver

**Files:**
- `app/utils/xlsx-exporter.ts` - XLSX export utility
- Client components can import and use directly

## Features Preserved

### DOCX Export Supports:
- ✅ Text formatting (bold, italic, underline, strikethrough)
- ✅ Font families and sizes
- ✅ Text colors and highlights
- ✅ Paragraph alignment (left, center, right, justified)
- ✅ Line spacing and paragraph spacing
- ✅ Indentation (first line, hanging, left, right)
- ✅ Bullet and numbered lists (nested)
- ✅ Images (embedded with proper sizing)
- ✅ Superscript and subscript
- ✅ Page margins and size
- ✅ Document structure validation

### XLSX Export Supports:
- ✅ Cell styles (bold, italic, underline, strikethrough)
- ✅ Font families and sizes
- ✅ Text colors and background colors
- ✅ Cell borders (all sides)
- ✅ Cell alignment (horizontal and vertical)
- ✅ Formulas (preserved with calculated values)
- ✅ Merged cells
- ✅ Column widths and row heights
- ✅ Hidden rows and columns
- ✅ Multiple sheets with proper ordering
- ✅ Shared strings (optimized file size)

## Usage

### DOCX Export

```typescript
// In your component
const handleExport = async () => {
  const snapshot = univerAPI.getActiveDocument().getSnapshot();
  
  const response = await fetch("/export-docx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      documentData: snapshot,
      filename: "my-document.docx",
    }),
  });

  if (response.ok) {
    const blob = await response.blob();
    saveAs(blob, "my-document.docx");
  }
};
```

### XLSX Export

```typescript
import { exportWorkbookToXlsx } from "~/utils/xlsx-exporter";

// In your component
const handleExport = async () => {
  const snapshot = univerAPI.getActiveWorkbook().getSnapshot();
  await exportWorkbookToXlsx(snapshot, "my-spreadsheet.xlsx");
};
```

## Technical Details

### DOCX OOXML Structure

The DOCX exporter creates a valid OOXML package:

```
document.docx (ZIP)
├── [Content_Types].xml     # MIME types
├── _rels/
│   └── .rels               # Root relationships
└── word/
    ├── document.xml        # Main document content
    ├── styles.xml          # Style definitions
    ├── numbering.xml       # List definitions
    ├── _rels/
    │   └── document.xml.rels  # Document relationships
    └── media/              # Embedded images
        ├── image1.png
        └── image2.jpg
```

### XLSX OOXML Structure

The XLSX exporter creates a valid OOXML package:

```
workbook.xlsx (ZIP)
├── [Content_Types].xml     # MIME types
├── _rels/
│   └── .rels               # Root relationships
└── xl/
    ├── workbook.xml        # Workbook structure
    ├── styles.xml          # Cell styles
    ├── sharedStrings.xml   # Shared text values
    ├── _rels/
    │   └── workbook.xml.rels  # Workbook relationships
    └── worksheets/
        ├── sheet1.xml
        └── sheet2.xml
```

### OOXML Validation

The DOCX builder includes validation to ensure:
- DataStream ends with `\r\n` (required by Univer)
- All paragraph startIndex values are within dataStream
- All textRun positions are valid
- CustomRanges point to correct markers (0x1A for tables/images)

### Performance Considerations

**DOCX:**
- Server-side processing prevents browser memory issues
- Large documents (>100 pages) export reliably
- No client-side performance impact during export

**XLSX:**
- Client-side processing is efficient for typical spreadsheets
- Shared strings reduce file size for repeated text
- Streaming not currently implemented (future enhancement)

## Error Handling

Both exporters include comprehensive error handling:

```typescript
try {
  await exportWorkbookToXlsx(snapshot, filename);
} catch (error) {
  console.error("[XLSX Export] Failed:", error);
  alert("Export failed: " + error.message);
}
```

Server-side exports return proper HTTP status codes:
- 200: Success
- 400: Invalid document data
- 500: Server processing error

## Dependencies

```json
{
  "jszip": "^3.10.1",          // ZIP file creation
  "file-saver": "^2.0.5",      // Client-side file download
  "xmlbuilder2": "^3.0.0",     // OOXML XML generation
  "@univerjs/core": "^0.15.3"  // Document models
}
```

## Future Enhancements

- [ ] CSV export for spreadsheets
- [ ] PDF export (via server-side rendering)
- [ ] Table support in DOCX export
- [ ] Chart support in XLSX export
- [ ] Streaming for very large files
- [ ] Export progress indicators
- [ ] Batch export multiple documents
- [ ] Custom templates for DOCX
- [ ] Conditional formatting in XLSX

## Troubleshooting

### "Invalid document structure" error
- Check that dataStream ends with `\r\n`
- Verify all textRun positions are within dataStream length
- Ensure customRanges point to valid positions

### Missing formatting in export
- Check that snapshot includes `styles` object
- Verify textRuns have proper `ts` (text style) properties
- Confirm paragraphStyle properties are set

### File download doesn't start
- Check browser console for CORS errors
- Verify `/export-docx` route is accessible
- Check that blob is created successfully

### XLSX formulas not working
- Ensure formula strings are properly escaped
- Include both formula `f` and value `v` properties
- Test in Excel/LibreOffice for compatibility
