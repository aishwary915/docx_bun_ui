# Fix: List Elements Not Deletable in Univer Editor

## Problem

When importing DOCX files with ordered or unordered lists, users were unable to delete list elements in the Univer document editor. The cursor would not properly select or remove list items.

## Root Cause

Three bugs in `app/utils/docx-converter.ts` were causing malformed data structures in the generated Univer document model:

1. **Double-wrapped indent objects** - Indent values were being wrapped twice, creating invalid `{ v: { v: number } }` structures
2. **Inconsistent property formats** - Some list properties used plain numbers while others used `{ v: number }` objects
3. **Mismatched `listType` values** - Document-level list definitions used numeric IDs instead of type strings

## Changes Made

### 1. Fixed Double-Wrapping of Indent Values

**Location:** Lines 910-915

The `levelDef.indentStart` and `levelDef.hanging` properties are already `{ v: number }` objects from the list definition. The code was incorrectly wrapping them again.

**Before:**
```typescript
if (levelDef) {
  paragraph.paragraphStyle.indentStart = {
    v: levelDef.indentStart || 36 * (bullet.nestingLevel + 1),
  };
  paragraph.paragraphStyle.hanging = {
    v: levelDef.hanging || 24,
  };
}
```

**After:**
```typescript
if (levelDef) {
  // levelDef.indentStart and levelDef.hanging are already { v: number } objects
  paragraph.paragraphStyle.indentStart = levelDef.indentStart || {
    v: 36 * (bullet.nestingLevel + 1),
  };
  paragraph.paragraphStyle.hanging = levelDef.hanging || { v: 24 };
}
```

### 2. Fixed Inconsistent Property Formats in Fallback List Definition

**Location:** Lines 622-635

The fallback list definition (used when no abstract numbering levels are found) was using plain numbers instead of the expected `{ v: number }` object format.

**Before:**
```typescript
lists[listId] = {
  listType: listId,
  nestingLevel: Array.from({ length: 9 }, (_, level) => ({
    bulletAlignment: 0,
    glyphFormat: "•",
    startNumber: 0,
    glyphType: 0,
    textStyle: { fs: 11 },
    hanging: 24,                      // Plain number
    indentStart: 36 * (level + 1),    // Plain number
  })),
};
```

**After:**
```typescript
lists[listId] = {
  listType: "BULLET_LIST",
  nestingLevel: Array.from({ length: 9 }, (_, level) => ({
    bulletAlignment: 0,
    glyphFormat: "•",
    startNumber: 0,
    glyphType: 0,
    textStyle: { fs: 11 },
    hanging: { v: 24 },                      // Object format
    indentStart: { v: 36 * (level + 1) },    // Object format
  })),
};
```

### 3. Fixed `listType` Value Mismatch

**Location:** Lines 598-621

The document-level `lists` dictionary was setting `listType` to the numeric list ID (e.g., `"1"`, `"2"`), but `paragraph.bullet.listType` uses string constants (`"BULLET_LIST"` or `"ORDER_LIST"`). This mismatch caused inconsistencies in how Univer interpreted list types.

**Before:**
```typescript
lists[listId] = {
  listType: listId,  // Was using the numeric ID like "1", "2"
  // ...
};
```

**After:**
```typescript
// Determine list type from the first level's glyphType (0 = bullet, 1 = ordered)
const firstLevelGlyphType =
  abstractLevels[0]?.glyphType ??
  (abstractLevels[0]?.numFmt === "bullet" ? 0 : 1);
const listTypeValue =
  firstLevelGlyphType === 0 ? "BULLET_LIST" : "ORDER_LIST";

lists[listId] = {
  listType: listTypeValue,  // Now uses "BULLET_LIST" or "ORDER_LIST"
  // ...
};
```

## Technical Details

### Univer Document Model Structure

Univer expects list-related properties to follow specific formats:

- **Indent values** must be objects: `{ v: number }` (where `v` is the value in points)
- **`listType`** should be a string constant: `"BULLET_LIST"` or `"ORDER_LIST"`
- **Document-level `lists`** and **paragraph `bullet`** properties must use consistent type values

### Data Flow

```
DOCX numbering.xml → abstractNumDefs → lists[listId] → paragraph.bullet
                                            ↓
                              paragraph.paragraphStyle.indentStart
                              paragraph.paragraphStyle.hanging
```

The fixes ensure data consistency throughout this flow.

## Testing

To verify the fix:

1. Import a DOCX file containing bulleted lists
2. Import a DOCX file containing numbered lists
3. Verify that list items can be:
   - Selected with cursor
   - Deleted with backspace/delete keys
   - Modified (text editing)
   - Have their list formatting removed

## Files Changed

- `app/utils/docx-converter.ts`

---

# Reference: Text Alignment Default Behavior

## Why Text Defaults to Left Alignment

When opening converted DOCX files in Univer, paragraphs may appear left-aligned by default. This is intentional behavior that mirrors Microsoft Word's defaults.

### How Alignment is Determined

The converter checks for alignment in this order:

1. **Explicit paragraph alignment** - The `w:jc` element in the paragraph's properties
2. **Style-based alignment** - Alignment defined in the paragraph's referenced style
3. **Special case for Title** - "Title" style paragraphs default to center
4. **Fallback default** - Left alignment (value `0`)

### Alignment Values

Univer uses numeric values for `horizontalAlign`:

| Value | Alignment |
|-------|-----------|
| 0     | Left      |
| 1     | Center    |
| 2     | Right     |
| 3     | Justify   |

### Code Location

The default is set in `app/utils/docx-converter.ts` at line 884:

```typescript
// Set alignment using Univer's expected format (0=Left, 1=Center, 2=Right, 3=Justify)
if (textAlignment !== undefined) {
  paragraph.paragraphStyle.horizontalAlign = textAlignment;
} else {
  paragraph.paragraphStyle.horizontalAlign = 0; // Default to left
}
```

### Changing the Default

To change the default alignment, modify line 884:

```typescript
// Examples:
paragraph.paragraphStyle.horizontalAlign = 0; // Left (current default)
paragraph.paragraphStyle.horizontalAlign = 1; // Center
paragraph.paragraphStyle.horizontalAlign = 2; // Right
paragraph.paragraphStyle.horizontalAlign = 3; // Justify
```

### Why Left is the Default

- Matches Microsoft Word's behavior for paragraphs without explicit alignment
- Most body text in documents is left-aligned
- Provides predictable behavior when source DOCX lacks alignment metadata
