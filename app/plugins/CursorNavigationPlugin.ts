import type { DocumentDataModel, IDisposable } from "@univerjs/core";
import {
  CommandType,
  ICommandService,
  Inject,
  Injector,
  IUniverInstanceService,
  Plugin,
  UniverInstanceType,
} from "@univerjs/core";
import { DocSelectionManagerService } from "@univerjs/docs";
import { IShortcutService, KeyCode, MetaKeys } from "@univerjs/ui";

// KeyCodes
const KEY_HOME = 36;
const KEY_END = 35;
const KEY_A = 65;

// ============================================================================
// Helper Functions for Word and Line Boundary Detection
// ============================================================================

/**
 * Find the previous word boundary from the given position.
 * Uses Intl.Segmenter when available for proper word boundary detection.
 */
function findPrevWordBoundary(dataStream: string, position: number): number {
  if (position <= 0) return 0;

  // Use Intl.Segmenter for proper word segmentation if available
  if (typeof Intl !== "undefined" && (Intl as any).Segmenter) {
    try {
      const segmenter = new (Intl as any).Segmenter(undefined, {
        granularity: "word",
      });
      const segments = Array.from(
        segmenter.segment(dataStream.substring(0, position))
      );

      // Find the last word-like segment before position
      let lastWordEnd = 0;
      for (const segment of segments as any[]) {
        if (segment.isWordLike) {
          const segEnd = segment.index + segment.segment.length;
          if (segEnd < position) {
            lastWordEnd = segment.index;
          } else if (segment.index < position) {
            // We're in the middle of a word, go to start of this word
            return segment.index;
          }
        }
      }

      return lastWordEnd;
    } catch (e) {
      // Fall back to regex-based approach
    }
  }

  // Fallback: regex-based word boundary detection
  // Skip any whitespace immediately before position
  let pos = position - 1;
  while (pos > 0 && /\s/.test(dataStream[pos])) {
    pos--;
  }

  // Find the start of the current/previous word
  while (pos > 0 && /\w/.test(dataStream[pos - 1])) {
    pos--;
  }

  return pos;
}

/**
 * Find the next word boundary from the given position.
 * Uses Intl.Segmenter when available for proper word boundary detection.
 */
function findNextWordBoundary(dataStream: string, position: number): number {
  const maxPos = dataStream.length - 1; // -1 to account for document end marker
  if (position >= maxPos) return maxPos;

  // Use Intl.Segmenter for proper word segmentation if available
  if (typeof Intl !== "undefined" && (Intl as any).Segmenter) {
    try {
      const segmenter = new (Intl as any).Segmenter(undefined, {
        granularity: "word",
      });
      const segments = Array.from(
        segmenter.segment(dataStream.substring(position))
      );

      // Find the first word-like segment after position
      for (const segment of segments as any[]) {
        if (segment.isWordLike && segment.index > 0) {
          // Found a word that starts after our position
          return Math.min(
            position + segment.index + segment.segment.length,
            maxPos
          );
        }
        if (segment.isWordLike && segment.index === 0) {
          // We're at the start of a word, skip to end of this word
          return Math.min(position + segment.segment.length, maxPos);
        }
      }

      return maxPos;
    } catch (e) {
      // Fall back to regex-based approach
    }
  }

  // Fallback: regex-based word boundary detection
  let pos = position;

  // Skip any whitespace at position
  while (pos < maxPos && /\s/.test(dataStream[pos])) {
    pos++;
  }

  // Skip the current word
  while (pos < maxPos && /\w/.test(dataStream[pos])) {
    pos++;
  }

  return Math.min(pos, maxPos);
}

/**
 * Find the start of the line/paragraph containing the given position.
 */
function findLineStart(
  dataStream: string,
  paragraphs: Array<{ startIndex: number }>,
  position: number
): number {
  // Find the paragraph containing this position
  let paragraphStart = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const nextP = paragraphs[i + 1];

    if (p.startIndex <= position && (!nextP || nextP.startIndex > position)) {
      // This paragraph contains the position
      // The content starts after the previous paragraph's newline
      if (i === 0) {
        paragraphStart = 0;
      } else {
        // Paragraph startIndex points to the paragraph marker (usually \r)
        // Content starts after that
        paragraphStart = p.startIndex + 1;
      }
      break;
    }
  }

  // Also look for newline characters to handle cases where paragraph info is incomplete
  const textBefore = dataStream.substring(0, position);
  const lastNewline = Math.max(
    textBefore.lastIndexOf("\n"),
    textBefore.lastIndexOf("\r")
  );

  if (lastNewline >= paragraphStart) {
    return lastNewline + 1;
  }

  return paragraphStart;
}

/**
 * Find the end of the line/paragraph containing the given position.
 */
function findLineEnd(
  dataStream: string,
  paragraphs: Array<{ startIndex: number }>,
  position: number
): number {
  const maxPos = dataStream.length - 1; // -1 to account for document end marker

  // Find the paragraph containing this position
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const nextP = paragraphs[i + 1];

    if (p.startIndex <= position && (!nextP || nextP.startIndex > position)) {
      // This paragraph contains the position
      // The paragraph ends at the next paragraph's startIndex - 1, or at the end
      if (nextP) {
        return nextP.startIndex;
      }
      break;
    }
  }

  // Also look for newline characters
  const textAfter = dataStream.substring(position);
  const nextNewlineR = textAfter.indexOf("\r");
  const nextNewlineN = textAfter.indexOf("\n");

  let nextNewline = -1;
  if (nextNewlineR >= 0 && nextNewlineN >= 0) {
    nextNewline = Math.min(nextNewlineR, nextNewlineN);
  } else if (nextNewlineR >= 0) {
    nextNewline = nextNewlineR;
  } else if (nextNewlineN >= 0) {
    nextNewline = nextNewlineN;
  }

  if (nextNewline >= 0) {
    return Math.min(position + nextNewline, maxPos);
  }

  return maxPos;
}

/* ================= COMMANDS ================= */

/**
 * Command to move cursor to the previous word boundary
 */
export const MoveCursorPrevWordCommand = {
  id: "custom.doc.command.move-cursor-prev-word",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const currentPos = selection.collapsed
      ? selection.startOffset
      : Math.min(selection.startOffset, selection.endOffset);

    // Find previous word boundary
    const newPos = findPrevWordBoundary(dataStream, currentPos);

    // Update cursor position
    selectionManager.replaceDocRanges([
      {
        startOffset: newPos,
        endOffset: newPos,
      },
    ]);

    return true;
  },
};

/**
 * Command to move cursor to the next word boundary
 */
export const MoveCursorNextWordCommand = {
  id: "custom.doc.command.move-cursor-next-word",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const currentPos = selection.collapsed
      ? selection.startOffset
      : Math.max(selection.startOffset, selection.endOffset);

    // Find next word boundary
    const newPos = findNextWordBoundary(dataStream, currentPos);

    // Update cursor position
    selectionManager.replaceDocRanges([
      {
        startOffset: newPos,
        endOffset: newPos,
      },
    ]);

    return true;
  },
};

/**
 * Command to extend selection to the previous word boundary
 */
export const SelectPrevWordCommand = {
  id: "custom.doc.command.select-prev-word",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;

    // When extending selection, we move the end that's being extended
    const anchorPos = selection.collapsed
      ? selection.startOffset
      : selection.startOffset;
    const activePos = selection.collapsed
      ? selection.startOffset
      : selection.endOffset;

    // Find previous word boundary from active position
    const newActivePos = findPrevWordBoundary(
      dataStream,
      Math.min(anchorPos, activePos)
    );

    // Update selection
    selectionManager.replaceDocRanges([
      {
        startOffset: Math.min(newActivePos, Math.max(anchorPos, activePos)),
        endOffset: Math.max(newActivePos, Math.max(anchorPos, activePos)),
      },
    ]);

    return true;
  },
};

/**
 * Command to extend selection to the next word boundary
 */
export const SelectNextWordCommand = {
  id: "custom.doc.command.select-next-word",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;

    // When extending selection, we move the end that's being extended
    const anchorPos = selection.startOffset;
    const activePos = selection.endOffset;

    // Find next word boundary from the rightmost position
    const newActivePos = findNextWordBoundary(
      dataStream,
      Math.max(anchorPos, activePos)
    );

    // Update selection
    selectionManager.replaceDocRanges([
      {
        startOffset: Math.min(anchorPos, newActivePos),
        endOffset: Math.max(anchorPos, newActivePos),
      },
    ]);

    return true;
  },
};

/**
 * Command to move cursor to the start of the current line/paragraph
 */
export const MoveCursorLineStartCommand = {
  id: "custom.doc.command.move-cursor-line-start",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const paragraphs = body.paragraphs || [];
    const currentPos = selection.startOffset;

    // Find line/paragraph start
    const newPos = findLineStart(dataStream, paragraphs, currentPos);

    // Update cursor position
    selectionManager.replaceDocRanges([
      {
        startOffset: newPos,
        endOffset: newPos,
      },
    ]);

    return true;
  },
};

/**
 * Command to move cursor to the end of the current line/paragraph
 */
export const MoveCursorLineEndCommand = {
  id: "custom.doc.command.move-cursor-line-end",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const paragraphs = body.paragraphs || [];
    const currentPos = selection.endOffset;

    // Find line/paragraph end
    const newPos = findLineEnd(dataStream, paragraphs, currentPos);

    // Update cursor position
    selectionManager.replaceDocRanges([
      {
        startOffset: newPos,
        endOffset: newPos,
      },
    ]);

    return true;
  },
};

/**
 * Command to extend selection to the start of the current line/paragraph
 */
export const SelectLineStartCommand = {
  id: "custom.doc.command.select-line-start",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const paragraphs = body.paragraphs || [];

    // Find line start from the leftmost selected position
    const leftPos = Math.min(selection.startOffset, selection.endOffset);
    const rightPos = Math.max(selection.startOffset, selection.endOffset);
    const lineStart = findLineStart(dataStream, paragraphs, leftPos);

    // Update selection
    selectionManager.replaceDocRanges([
      {
        startOffset: lineStart,
        endOffset: rightPos,
      },
    ]);

    return true;
  },
};

/**
 * Command to extend selection to the end of the current line/paragraph
 */
export const SelectLineEndCommand = {
  id: "custom.doc.command.select-line-end",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const paragraphs = body.paragraphs || [];

    // Find line end from the rightmost selected position
    const leftPos = Math.min(selection.startOffset, selection.endOffset);
    const rightPos = Math.max(selection.startOffset, selection.endOffset);
    const lineEnd = findLineEnd(dataStream, paragraphs, rightPos);

    // Update selection
    selectionManager.replaceDocRanges([
      {
        startOffset: leftPos,
        endOffset: lineEnd,
      },
    ]);

    return true;
  },
};

/**
 * Command to delete the previous word (from cursor to previous word boundary)
 * macOS: Option + Backspace
 * Windows/Linux: Ctrl + Backspace
 */
export const DeletePrevWordCommand = {
  id: "custom.doc.command.delete-prev-word",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);
    const commandService = accessor.get(ICommandService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const currentPos = selection.collapsed
      ? selection.startOffset
      : Math.min(selection.startOffset, selection.endOffset);

    // If there's a selection, just delete it
    if (!selection.collapsed) {
      // Use Univer's built-in delete command for selections
      return commandService.executeCommand("doc.command.delete-left");
    }

    // Find previous word boundary
    const prevWordPos = findPrevWordBoundary(dataStream, currentPos);

    if (prevWordPos === currentPos) return false;

    // Select the text from prev word boundary to current position, then delete
    selectionManager.replaceDocRanges([
      {
        startOffset: prevWordPos,
        endOffset: currentPos,
      },
    ]);

    // Execute delete command to remove the selected text
    return commandService.executeCommand("doc.command.delete-left");
  },
};

/**
 * Command to delete the next word (from cursor to next word boundary)
 * macOS: Option + Delete (Fn+Backspace)
 * Windows/Linux: Ctrl + Delete
 */
export const DeleteNextWordCommand = {
  id: "custom.doc.command.delete-next-word",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);
    const commandService = accessor.get(ICommandService);

    const selection = selectionManager.getActiveTextRange();
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!(selection && currentDoc)) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const dataStream = body.dataStream;
    const currentPos = selection.collapsed
      ? selection.startOffset
      : Math.max(selection.startOffset, selection.endOffset);

    // If there's a selection, just delete it
    if (!selection.collapsed) {
      // Use Univer's built-in delete command for selections
      return commandService.executeCommand("doc.command.delete-right");
    }

    // Find next word boundary
    const nextWordPos = findNextWordBoundary(dataStream, currentPos);

    if (nextWordPos === currentPos) return false;

    // Select the text from current position to next word boundary, then delete
    selectionManager.replaceDocRanges([
      {
        startOffset: currentPos,
        endOffset: nextWordPos,
      },
    ]);

    // Execute delete command to remove the selected text
    return commandService.executeCommand("doc.command.delete-right");
  },
};

export const SelectAllCommand = {
  id: "custom.doc.command.select-all",
  type: CommandType.COMMAND,
  handler: async (accessor: any) => {
    const selectionManager = accessor.get(DocSelectionManagerService);
    const univerInstanceService = accessor.get(IUniverInstanceService);

    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC
      );

    if (!currentDoc) return false;

    const body = currentDoc.getBody();
    if (!body) return false;

    const paragraphs = body.paragraphs || [];

    if (!paragraphs.length) return false;

    const startOffset = 0;
    const lastParagraph = paragraphs[paragraphs.length - 1];
    const endOffset =
      (lastParagraph.startIndex ?? 0) + (lastParagraph.textRun?.length ?? 0);

    if (endOffset <= 0) return false;

    selectionManager.replaceDocRanges([
      {
        startOffset,
        endOffset,
      },
    ]);

    return true;
  },
};

/* ================= PLUGIN ================= */

export class CursorNavigationPlugin extends Plugin {
  static override pluginName = "CURSOR_NAVIGATION_PLUGIN";
  static override type = UniverInstanceType.UNIVER_DOC;

  private _disposables: IDisposable[] = [];

  constructor(
    _config: unknown,
    @Inject(Injector) protected readonly _injector: Injector
  ) {
    super();
  }

  override onReady(): void {
    this._registerCommands();
    this._registerShortcuts();
    this._setupSelectAllListener();
    console.log(
      "✅ CursorNavigationPlugin: Word/line navigation, word deletion, and select-all shortcuts registered"
    );
  }

  override onDestroy(): void {
    this._disposables.forEach((d) => d.dispose());
    this._disposables = [];
  }

  private _registerCommands(): void {
    const commandService = this._injector.get(ICommandService);

    const commands = [
      MoveCursorPrevWordCommand,
      MoveCursorNextWordCommand,
      SelectPrevWordCommand,
      SelectNextWordCommand,
      MoveCursorLineStartCommand,
      MoveCursorLineEndCommand,
      SelectLineStartCommand,
      SelectLineEndCommand,
      DeletePrevWordCommand,
      DeleteNextWordCommand,
      SelectAllCommand,
    ];

    for (const command of commands) {
      this._disposables.push(commandService.registerCommand(command));
    }
  }

  private _registerShortcuts(): void {
    const shortcutService = this._injector.get(IShortcutService);

    const isMac =
      typeof navigator !== "undefined" &&
      /Mac|iPod|iPhone|iPad/.test(navigator.platform);

    // ===================================
    // Navigation Shortcuts
    // ===================================

    // Move cursor to previous word
    this._disposables.push(
      shortcutService.registerShortcut({
        id: MoveCursorPrevWordCommand.id,
        description: "Move cursor to previous word",
        binding: isMac
          ? KeyCode.ARROW_LEFT | MetaKeys.ALT
          : KeyCode.ARROW_LEFT | MetaKeys.CTRL_COMMAND,
        priority: 100,
      })
    );

    // Move cursor to next word
    this._disposables.push(
      shortcutService.registerShortcut({
        id: MoveCursorNextWordCommand.id,
        description: "Move cursor to next word",
        binding: isMac
          ? KeyCode.ARROW_RIGHT | MetaKeys.ALT
          : KeyCode.ARROW_RIGHT | MetaKeys.CTRL_COMMAND,
        priority: 100,
      })
    );

    // Select previous word
    this._disposables.push(
      shortcutService.registerShortcut({
        id: SelectPrevWordCommand.id,
        description: "Select previous word",
        binding: isMac
          ? KeyCode.ARROW_LEFT | MetaKeys.ALT | MetaKeys.SHIFT
          : KeyCode.ARROW_LEFT | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
        priority: 100,
      })
    );

    // Select next word
    this._disposables.push(
      shortcutService.registerShortcut({
        id: SelectNextWordCommand.id,
        description: "Select next word",
        binding: isMac
          ? KeyCode.ARROW_RIGHT | MetaKeys.ALT | MetaKeys.SHIFT
          : KeyCode.ARROW_RIGHT | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
        priority: 100,
      })
    );

    // Line Start
    this._disposables.push(
      shortcutService.registerShortcut({
        id: MoveCursorLineStartCommand.id,
        description: "Move to line start",
        binding: isMac ? KeyCode.ARROW_LEFT | MetaKeys.CTRL_COMMAND : KEY_HOME,
        priority: 100,
      })
    );

    // Line End
    this._disposables.push(
      shortcutService.registerShortcut({
        id: MoveCursorLineEndCommand.id,
        description: "Move to line end",
        binding: isMac ? KeyCode.ARROW_RIGHT | MetaKeys.CTRL_COMMAND : KEY_END,
        priority: 100,
      })
    );

    // Select Line Start
    this._disposables.push(
      shortcutService.registerShortcut({
        id: SelectLineStartCommand.id,
        description: "Select to line start",
        binding: isMac
          ? KeyCode.ARROW_LEFT | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT
          : KEY_HOME | MetaKeys.SHIFT,
        priority: 100,
      })
    );

    // Select Line End
    this._disposables.push(
      shortcutService.registerShortcut({
        id: SelectLineEndCommand.id,
        description: "Select to line end",
        binding: isMac
          ? KeyCode.ARROW_RIGHT | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT
          : KEY_END | MetaKeys.SHIFT,
        priority: 100,
      })
    );

    // Delete Previous Word
    this._disposables.push(
      shortcutService.registerShortcut({
        id: DeletePrevWordCommand.id,
        description: "Delete previous word",
        binding: isMac
          ? KeyCode.BACKSPACE | MetaKeys.ALT
          : KeyCode.BACKSPACE | MetaKeys.CTRL_COMMAND,
        priority: 100,
      })
    );

    // Delete Next Word
    this._disposables.push(
      shortcutService.registerShortcut({
        id: DeleteNextWordCommand.id,
        description: "Delete next word",
        binding: isMac
          ? KeyCode.DELETE | MetaKeys.ALT
          : KeyCode.DELETE | MetaKeys.CTRL_COMMAND,
        priority: 100,
      })
    );

    // Select All
    this._disposables.push(
      shortcutService.registerShortcut({
        id: SelectAllCommand.id,
        description: "Select all content",
        binding: KEY_A | MetaKeys.CTRL_COMMAND,
        mac: KEY_A | MetaKeys.CTRL_COMMAND,
        win: KEY_A | MetaKeys.CTRL_COMMAND,
        linux: KEY_A | MetaKeys.CTRL_COMMAND,
        priority: 999,
      })
    );
  }

  private _setupSelectAllListener(): void {
    const commandService = this._injector.get(ICommandService);

    const handleKeyDown = (e: KeyboardEvent) => {
      const isA = e.key === "a" || e.key === "A" || e.code === "KeyA";
      const isCmdOrCtrl = e.metaKey || e.ctrlKey;

      if (!(isCmdOrCtrl && isA)) return;

      const target = e.target as HTMLElement;
      const tag = target?.tagName?.toLowerCase();

      const isNativeInput =
        tag === "input" || tag === "textarea" || target?.isContentEditable;

      const isUniver =
        target?.closest?.(".univer-doc-editor") ||
        target?.closest?.(".univer-app-container") ||
        tag === "canvas";

      if (!isUniver || isNativeInput) return;

      e.preventDefault();
      e.stopPropagation();

      commandService.executeCommand(SelectAllCommand.id);
    };

    window.addEventListener("keydown", handleKeyDown, true);

    this._disposables.push({
      dispose: () => window.removeEventListener("keydown", handleKeyDown, true),
    });

    console.log("✅ Select All listener attached");
  }
}
