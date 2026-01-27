/**
 * Horizontal Line Spacing Plugin for Univer Docs
 *
 * This plugin adds visual spacing around horizontal lines by:
 * 1. Modifying the horizontal line padding value
 * 2. Inserting an empty paragraph before the horizontal line for spacing
 */

import {
    Disposable,
    ICommandService,
    Inject,
    IUniverInstanceService,
    UniverInstanceType,
    Plugin,
    Injector,
} from "@univerjs/core";
import { DocSelectionManagerService } from "@univerjs/docs";

const BREAK_LINE_COMMAND_ID = "doc.command.break-line";
const INSERT_PARAGRAPH_COMMAND_ID = "doc.command.break-line";

// Spacing configuration
const HORIZONTAL_LINE_PADDING = 30;

export class HorizontalLineSpacingService extends Disposable {
    constructor(
        @Inject(ICommandService) private readonly _commandService: ICommandService,
        @Inject(IUniverInstanceService) private readonly _univerInstanceService: IUniverInstanceService,
        @Inject(DocSelectionManagerService) private readonly _docSelectionManagerService: DocSelectionManagerService,
    ) {
        super();
        this._init();
    }

    private _init(): void {
        // Intercept the break-line command BEFORE execution
        const beforeListener = this._commandService.beforeCommandExecuted((commandInfo) => {
            if (commandInfo.id === BREAK_LINE_COMMAND_ID) {
                const params = commandInfo.params as any;

                if (params?.horizontalLine) {
                    // Increase padding
                    params.horizontalLine.padding = HORIZONTAL_LINE_PADDING;

                    console.log("[HorizontalLineSpacingPlugin] Modified horizontal line params");
                }
            }
        });

        // After horizontal line is inserted, add an empty line for spacing
        const afterListener = this._commandService.onCommandExecuted((commandInfo) => {
            if (commandInfo.id === BREAK_LINE_COMMAND_ID) {
                const params = commandInfo.params as any;

                if (params?.horizontalLine) {
                    console.log("[HorizontalLineSpacingPlugin] Horizontal line inserted, adding spacing paragraph");

                    // Insert an empty paragraph after the horizontal line for visual spacing
                    setTimeout(() => {
                        this._insertSpacingParagraph();
                    }, 50);
                }
            }
        });

        this.disposeWithMe(beforeListener);
        this.disposeWithMe(afterListener);

        console.log("[HorizontalLineSpacingPlugin] Initialized");
    }

    private _insertSpacingParagraph(): void {
        try {
            // Insert a simple line break (empty paragraph) for spacing
            this._commandService.executeCommand(BREAK_LINE_COMMAND_ID, {});
            console.log("[HorizontalLineSpacingPlugin] Spacing paragraph inserted");
        } catch (e) {
            console.error("[HorizontalLineSpacingPlugin] Error inserting spacing:", e);
        }
    }
}

export class HorizontalLineSpacingPlugin extends Plugin {
    static override pluginName = "HORIZONTAL_LINE_SPACING_PLUGIN";
    static override type = UniverInstanceType.UNIVER_DOC;

    private _service: HorizontalLineSpacingService | null = null;

    constructor(
        _config: unknown,
        @Inject(Injector) protected readonly _injector: Injector,
    ) {
        super();
    }

    override onReady(): void {
        this._service = this._injector.createInstance(HorizontalLineSpacingService);
    }

    override onDestroy(): void {
        this._service?.dispose();
        this._service = null;
    }
}
