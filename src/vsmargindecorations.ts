"use strict";

import { DecorationOptions, Range, TextEditor, Position, window, ThemeColor, TextDocument, workspace, TextEditorDecorationType } from "vscode";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { ESourceFormat } from "./externalfeatures";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { VSWorkspaceFolders } from "./vscobolfolders";
import { VSCodeSourceHandlerLite } from "./vscodesourcehandler";
import { SourceFormat } from "./sourceformat";
import { TextLanguage, VSExtensionUtils } from "./vsextutis";
import { ColourTagHandler } from "./vscolourcomments";
import { fileformatStrategy, ICOBOLSettings } from "./iconfiguration";
import { VSExternalFeatures } from "./vsexternalfeatures";

const defaultTrailingSpacesDecoration: TextEditorDecorationType = window.createTextEditorDecorationType({
    light: {
        color: new ThemeColor("editorLineNumber.foreground"),
        backgroundColor: new ThemeColor("editor.background"),
        textDecoration: "solid"
    },
    dark: {
        color: new ThemeColor("editorLineNumber.foreground"),
        backgroundColor: new ThemeColor("editor.background"),
        textDecoration: "solid"
    }
});

interface MarginRange {
    start: number;
    end: number;
    maxTagLength: number;
}

export class VSmargindecorations extends ColourTagHandler {
    private tags = new Map<string, TextEditorDecorationType>();

    constructor() {
        super();
        this.setupTags();
    }

    public setupTags(): void {
        super.setupTags("columns_tags", this.tags);
    }

    private isEnabledViaWorkspace4jcl(settings: ICOBOLSettings): boolean {
        if (!VSWorkspaceFolders.get(settings)) return false;
        return workspace.getConfiguration("jcleditor").get<boolean>("margin") ?? true;
    }

    private async updateJCLDecorations(doc: TextDocument, activeTextEditor: TextEditor, decoration: TextEditorDecorationType) {
        const decorations: DecorationOptions[] = [];
        for (let i = 0; i < doc.lineCount; i++) {
            const line = doc.lineAt(i).text;
            if (line.length > 72) {
                decorations.push({ range: new Range(new Position(i, 72), new Position(i, line.length)) });
            }
        }
        activeTextEditor.setDecorations(decoration, decorations);
    }

    private addDecorationsForColumnRange(
        line: string,
        lineIndex: number,
        rangeDef: MarginRange,
        declsMap: Map<string, DecorationOptions[]>,
        defaultDecorations: DecorationOptions[],
        enableTags: boolean
    ) {
        const rangeStart = rangeDef.start;
        const rangeEnd = Math.min(line.length, rangeDef.end);
        const baseDecoration: DecorationOptions = { range: new Range(new Position(lineIndex, rangeStart), new Position(lineIndex, rangeEnd)) };
        let useDefault = true;

        if (enableTags) {
            const text = line.slice(rangeStart, rangeEnd);
            for (const [tag] of this.tags) {
                if (tag.length > rangeDef.maxTagLength) continue;
                const tagIndex = text.indexOf(tag);
                if (tagIndex !== -1) {
                    const items = declsMap.get(tag);
                    if (!items) continue;
                    useDefault = false;
                    const tagRange = new Range(
                        new Position(lineIndex, rangeStart + tagIndex),
                        new Position(lineIndex, rangeStart + tagIndex + tag.length)
                    );
                    items.push({ range: tagRange });

                    // add left and right colors if necessary
                    if (tagIndex > 0) defaultDecorations.push({ range: new Range(new Position(lineIndex, rangeStart), new Position(lineIndex, rangeStart + tagIndex)) });
                    if (tagIndex + tag.length < rangeDef.maxTagLength) defaultDecorations.push({ range: new Range(new Position(lineIndex, rangeStart + tagIndex + tag.length), new Position(lineIndex, rangeStart + rangeDef.maxTagLength)) });
                }
            }
        }

        if (useDefault) {
            defaultDecorations.push(baseDecoration);
        }
    }

    public async updateDecorations(activeTextEditor: TextEditor | undefined) {
        if (!activeTextEditor) return;

        const doc = activeTextEditor.document;
        const defaultDecorations: DecorationOptions[] = [];
        const configHandler = VSCOBOLConfiguration.get_resource_settings(doc, VSExternalFeatures);
        const textLanguage = VSExtensionUtils.isSupportedLanguage(doc);

        if (textLanguage === TextLanguage.Unknown) {
            activeTextEditor.setDecorations(defaultTrailingSpacesDecoration, defaultDecorations);
            return;
        }

        if (textLanguage === TextLanguage.JCL) {
            if (!this.isEnabledViaWorkspace4jcl(configHandler)) {
                activeTextEditor.setDecorations(defaultTrailingSpacesDecoration, defaultDecorations);
            } else {
                await this.updateJCLDecorations(doc, activeTextEditor, defaultTrailingSpacesDecoration);
            }
            return;
        }

        if (!configHandler.margin) {
            activeTextEditor.setDecorations(defaultTrailingSpacesDecoration, defaultDecorations);
            return;
        }

        const declsMap = new Map<string, DecorationOptions[]>();
        for (const tag of this.tags.keys()) declsMap.set(tag, []);

        let sf: ESourceFormat = ESourceFormat.unknown;
        switch (configHandler.fileformat_strategy) {
            case fileformatStrategy.AlwaysFixed: sf = ESourceFormat.fixed; break;
            case fileformatStrategy.AlwaysVariable: sf = ESourceFormat.variable; break;
            default:
                const gcp = VSCOBOLSourceScanner.getCachedObject(doc, configHandler);
                sf = gcp ? gcp.sourceFormat : SourceFormat.get(new VSCodeSourceHandlerLite(doc), configHandler);
                break;
        }

        if ([ESourceFormat.free, ESourceFormat.unknown, ESourceFormat.terminal].includes(sf)) {
            activeTextEditor.setDecorations(defaultTrailingSpacesDecoration, defaultDecorations);
            return;
        }

        const decSequenceNumber = true;
        const decArea73_80 = sf !== ESourceFormat.variable && configHandler.margin_identification_area;
        const maxLineLength = configHandler.editor_maxTokenizationLineLength;

        const columnRanges: MarginRange[] = [];
        if (decSequenceNumber) columnRanges.push({ start: 0, end: 6, maxTagLength: 6 });
        if (decArea73_80) columnRanges.push({ start: 72, end: 80, maxTagLength: 8 });

        for (let i = 0; i < doc.lineCount; i++) {
            const line = doc.lineAt(i).text;
            if (line.length > maxLineLength) continue;
            const containsTab = line.indexOf("\t");

            for (const rangeDef of columnRanges) {
                if (containsTab === -1 || (rangeDef.start === 0 ? containsTab >= 6 : containsTab > 80)) {
                    this.addDecorationsForColumnRange(line, i, rangeDef, declsMap, defaultDecorations, configHandler.enable_columns_tags);
                }
            }
        }

        for (const [tag, decls] of declsMap) {
            const typeDecl = this.tags.get(tag);
            if (typeDecl) activeTextEditor.setDecorations(typeDecl, decls);
        }
        activeTextEditor.setDecorations(defaultTrailingSpacesDecoration, defaultDecorations);
    }
}

export const vsMarginHandler = new VSmargindecorations();
