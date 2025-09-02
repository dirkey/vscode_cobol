import * as vscode from "vscode";
import path from "path";
import { SourceScannerUtils, COBOLTokenStyle } from "./cobolsourcescanner";
import { cobolRegistersDictionary, cobolStorageKeywordDictionary, getCOBOLKeywordDictionary } from "./keywords/cobolKeywords";
import { VSLogger } from "./vslogger";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { InMemoryGlobalSymbolCache } from "./globalcachehelper";
import { COBOLFileUtils } from "./fileutils";
import { VSWorkspaceFolders } from "./vscobolfolders";
import { ICOBOLSettings, intellisenseStyle } from "./iconfiguration";
import { COBOLFileSymbol } from "./cobolglobalcache";
import { VSCOBOLFileUtils } from "./vsfileutils";
import { IExternalFeatures } from "./externalfeatures";
import { ExtensionDefaults } from "./extensionDefaults";
import { VSCustomIntelliseRules } from "./vscustomrules";
import { SplitTokenizer } from "./splittoken";
import { VSExtensionUtils } from "./vsextutis";
import { setMicroFocusSuppressFileAssociationsPrompt } from "./vscommon_commands";
import { VSCOBOLEditorConfiguration } from "./vsconfiguration";
import { workspace } from "vscode";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

let commandTerminal: vscode.Terminal | undefined = undefined;
const commandTerminalName = "COBOL Application";

export enum FoldAction {
    PerformTargets = 1,
    ConstantsOrVariables = 2,
    Keywords = 3
}

export enum AlignStyle {
    First = 1,
    Left = 2,
    Center = 3,
    Right = 4
}

export class VSCOBOLUtils {
    private static getGlobPattern(config: ICOBOLSettings, extensions: string[]): string {
        const prefix = config.maintain_metadata_recursive_search ? "**/*." : "*.";
        const filteredExts = extensions.filter(ext => ext.length > 0).join(",");
        return `${prefix}{${filteredExts}}`;
    }

    private static getProgramGlobPattern(config: ICOBOLSettings): string {
        return VSCOBOLUtils.getGlobPattern(config, config.program_extensions);
    }

    private static getCopyBookGlobPattern(config: ICOBOLSettings): string {
        return VSCOBOLUtils.getGlobPattern(config, config.copybookexts);
    }

    public static getCopyBookGlobPatternForPartialName(config: ICOBOLSettings, partialFilename = "*"): string {
        const extensions = config.copybookexts
            .filter(ext => ext.length > 0)
            .map(ext => `.${ext}`)
            .join(",");

        return `**/${partialFilename}{${extensions}}`;
    }

    static prevWorkSpaceUri: vscode.Uri | undefined = undefined;

    private static async populateDefaultSymbols(
        settings: ICOBOLSettings,
        reset: boolean,
        cache: Map<string, string>,
        getGlob: (config: ICOBOLSettings) => string,
        keyTransform: (fileName: string) => string = (f) => f.toLowerCase()
    ): Promise<void> {
        const wsFolders = VSWorkspaceFolders.get(settings);
        if (!wsFolders) return;

        const workspaceFile = vscode.workspace.workspaceFile;
        if (!workspaceFile) {
            cache.clear();
            return;
        }

        // Reset or skip if already cached
        if (reset) {
            VSCOBOLUtils.prevWorkSpaceUri = undefined;
        } else if (workspaceFile.fsPath === VSCOBOLUtils.prevWorkSpaceUri?.fsPath) {
            return;
        }

        VSCOBOLUtils.prevWorkSpaceUri = workspaceFile;

        const globPattern = getGlob(settings);
        const uris = await vscode.workspace.findFiles(globPattern);

        for (const uri of uris) {
            const fullPath = uri.fsPath;
            const fileName = path.basename(fullPath, path.extname(fullPath));
            const key = keyTransform(fileName);

            if (!cache.has(key)) {
                cache.set(key, fullPath);
            }
        }
    }

    // Usage for callable symbols
    static populateDefaultCallableSymbols(settings: ICOBOLSettings, reset: boolean) {
        void VSCOBOLUtils.populateDefaultSymbols(
            settings,
            reset,
            InMemoryGlobalSymbolCache.defaultCallableSymbols,
            VSCOBOLUtils.getProgramGlobPattern,
            (f) => f.toLowerCase()
        );
    }

    // Usage for copybooks
    static populateDefaultCopyBooks(settings: ICOBOLSettings, reset: boolean) {
        void VSCOBOLUtils.populateDefaultSymbols(
            settings,
            reset,
            InMemoryGlobalSymbolCache.defaultCopybooks,
            VSCOBOLUtils.getCopyBookGlobPattern
        );
    }

    private static typeToArray(types: string[], prefix: string, typeMap: Map<string, COBOLFileSymbol[]>) {
        for (const [typeKey, symbols] of typeMap.entries()) {
            if (!symbols) continue;

            for (const symbol of symbols) {
                types.push(`${prefix},${typeKey},${symbol.filename},${symbol.linenum}`);
            }
        }
    }

    public static clearGlobalCache(settings: ICOBOLSettings): void {
        // only update if we have a workspace
        if (VSWorkspaceFolders.get(settings) === undefined) {
            return;
        }

        const editorConfig = VSCOBOLEditorConfiguration.getEditorConfig();
        editorConfig.update("metadata_symbols", [], false);
        editorConfig.update("metadata_entrypoints", [], false);
        editorConfig.update("metadata_types", [], false);
        editorConfig.update("metadata_files", [], false);
        editorConfig.update("metadata_knowncopybooks", [], false);
        InMemoryGlobalSymbolCache.isDirty = false;
    }

    public static saveGlobalCacheToWorkspace(settings: ICOBOLSettings, update = true): void {
        // Only proceed if workspace exists and caching is enabled
        if (!VSWorkspaceFolders.get(settings) || vscode.workspace.workspaceFile === undefined || !settings.maintain_metadata_cache) {
            return;
        }

        if (!InMemoryGlobalSymbolCache.isDirty) return;

        try {
            // Callable symbols excluding default callables
            const symbols = Array.from(InMemoryGlobalSymbolCache.callableSymbols.entries())
                .flatMap(([symbolName, fileSymbols]) =>
                    !symbolName || !fileSymbols ? [] :
                    fileSymbols
                        .filter(() => !InMemoryGlobalSymbolCache.defaultCallableSymbols.has(symbolName))
                        .map(s => s.linenum !== 0 ? `${symbolName},${s.filename},${s.linenum}` : `${symbolName},${s.filename}`)
                );

            // Entry points
            const entrypoints = Array.from(InMemoryGlobalSymbolCache.entryPoints.entries())
                .flatMap(([entryName, fileSymbols]) =>
                    !entryName || !fileSymbols ? [] :
                    fileSymbols.map(s => `${entryName},${s.filename},${s.linenum}`)
                );

            // Known copybooks
            const knownCopybooks = Array.from(InMemoryGlobalSymbolCache.knownCopybooks.keys());

            // Types, interfaces, enums
            const types: string[] = [];
            VSCOBOLUtils.typeToArray(types, "T", InMemoryGlobalSymbolCache.types);
            VSCOBOLUtils.typeToArray(types, "I", InMemoryGlobalSymbolCache.interfaces);
            VSCOBOLUtils.typeToArray(types, "E", InMemoryGlobalSymbolCache.enums);

            // Modified source files
            const files = Array.from(InMemoryGlobalSymbolCache.sourceFilenameModified.values())
                .map(cws => `${cws.lastModifiedTime},${cws.workspaceFilename}`);

            // Update workspace configuration if requested
            if (update) {
                const editorConfig = VSCOBOLEditorConfiguration.getEditorConfig();
                editorConfig.update("metadata_symbols", symbols, false);
                editorConfig.update("metadata_entrypoints", entrypoints, false);
                editorConfig.update("metadata_types", types, false);
                editorConfig.update("metadata_files", files, false);
                editorConfig.update("metadata_knowncopybooks", knownCopybooks, false);

                InMemoryGlobalSymbolCache.isDirty = false;
            }
        } catch (e) {
            VSLogger.logException("Failed to update metadata", e as Error);
        }
    }

    public static inCopybookdirs(config: ICOBOLSettings, copybookdir: string): boolean {
        for (const ext of config.copybookdirs) {
            if (ext === copybookdir) {
                return true;
            }
        }

        return false;
    }

    private static adjustLineInternal(toCursor: boolean): void {
        const editor = vscode.window.activeTextEditor;
        if (!editor) return;

        const sel = editor.selection;
        const lineText = editor.document.lineAt(sel.start.line).text;
        const trimmedLine = lineText.trimStart();

        const startOfLine = new vscode.Position(sel.start.line, 0);
        const endOfLine = new vscode.Position(sel.start.line, lineText.length);
        const range = new vscode.Range(startOfLine, endOfLine);

        const newLine = toCursor
            ? " ".repeat(sel.start.character) + trimmedLine
            : trimmedLine;

        editor.edit(editBuilder => {
            editBuilder.replace(range, newLine);
        }).then(() => {
            editor.selection = toCursor
                ? sel // restore cursor for indentToCursor
                : new vscode.Selection(startOfLine, startOfLine); // move to start for leftAdjustLine
        });
    }

    public static indentToCursor(): void {
        VSCOBOLUtils.adjustLineInternal(true);
    }

    public static leftAdjustLine(): void {
        VSCOBOLUtils.adjustLineInternal(false);
    }

    public static transposeSelection(editor: vscode.TextEditor, editBuilder: vscode.TextEditorEdit): void {
        for (const selection of editor.selections) {
            // Handle single-character cursor (no selection)
            if (selection.isEmpty) {
                const pos = selection.active;
                const lineEnd = editor.document.lineAt(pos.line).range.end;

                // Skip if at the start or end of the line
                if (pos.character > 0 && pos.isBefore(lineEnd)) {
                    const nextCharRange = new vscode.Range(pos, pos.translate(0, 1));
                    const nextChar = editor.document.getText(nextCharRange);

                    editBuilder.delete(nextCharRange);
                    const prevPos = pos.translate(0, -1);
                    editBuilder.insert(prevPos, nextChar);
                }
            } else {
                // Handle a selection (non-empty)
                const start = selection.start;
                const end = selection.end;

                const firstCharRange = new vscode.Range(start, start.translate(0, 1));
                const firstChar = editor.document.getText(firstCharRange);

                editBuilder.delete(firstCharRange);
                editBuilder.insert(end, firstChar);
            }
        }
    }

    public static extractSelectionTo(editor: vscode.TextEditor, isParagraph: boolean): void {
        const selection = editor.selection;
        const selectedRange = new vscode.Range(selection.start, selection.end);
        const selectedText = editor.document.getText(selectedRange);

        vscode.window.showInputBox({
            prompt: isParagraph ? "New paragraph name?" : "New section name?",
            validateInput: (input: string): string | undefined => {
                if (!input || input.includes(" ")) {
                    return "Invalid paragraph or section name.";
                }
                return undefined;
            }
        }).then(name => {
            if (!name) return; // user canceled or invalid

            editor.edit(editBuilder => {
                // Determine the minimum indentation in the selected text
                const lines = selectedText.split(/\r?\n/);
                const minIndent = lines
                    .filter(line => line.trim().length > 0)
                    .reduce((min, line) => Math.min(min, line.match(/^ */)![0].length), Infinity);

                // Re-indent the selected text under the new paragraph/section
                const adjustedText = lines.map(line => {
                    if (line.trim().length === 0) return line; // keep empty lines
                    return "           " + line.slice(minIndent); // 11 spaces for COBOL block
                }).join("\n");

                const replacementText =
    `       PERFORM ${name}

        ${name}${isParagraph ? '.' : ' SECTION.'}
    ${adjustedText}
        .`;

                editBuilder.replace(selectedRange, replacementText);
            });
        });
    }

    private static pad(num: number, size: number): string {
        return num.toString().padStart(size, "0");
    }

    public static getMFUnitAnsiColorConfig(): boolean {
        const editorConfig = VSCOBOLEditorConfiguration.getEditorConfig();
        return editorConfig.get<boolean>("mfunit.diagnostic.color") ?? false;
    }

    public static resequenceColumnNumbers(
        activeEditor: vscode.TextEditor | undefined, 
        startValue: number, 
        increment: number
    ): void {
        if (!activeEditor) {
            return;
        }

        const edits = new vscode.WorkspaceEdit();
        const uri = activeEditor.document.uri;

        for (let lineNumber = 0; lineNumber < activeEditor.document.lineCount; lineNumber++) {
            const line = activeEditor.document.lineAt(lineNumber);
            if (line.text.length > 6) {
                const range = new vscode.Range(
                    new vscode.Position(lineNumber, 0),
                    new vscode.Position(lineNumber, 6)
                );
                const paddedValue = VSCOBOLUtils.pad(startValue + lineNumber * increment, 6);
                edits.replace(uri, range, paddedValue);
            }
        }

        vscode.workspace.applyEdit(edits);
    }

    public static removeColumnNumbers(activeEditor: vscode.TextEditor): void {
        const uri = activeEditor.document.uri;
        const edits: vscode.TextEdit[] = [];

        for (let lineNumber = 0; lineNumber < activeEditor.document.lineCount; lineNumber++) {
            const line = activeEditor.document.lineAt(lineNumber);
            if (line.text.length > 6) {
                const range = new vscode.Range(
                    new vscode.Position(lineNumber, 0),
                    new vscode.Position(lineNumber, 6)
                );
                edits.push(vscode.TextEdit.replace(range, "      "));
            }
        }

        if (edits.length > 0) {
            const workspaceEdit = new vscode.WorkspaceEdit();
            workspaceEdit.set(uri, edits);
            vscode.workspace.applyEdit(workspaceEdit);
        }
    }

    public static removeIdentificationArea(activeEditor: vscode.TextEditor): void {
        const uri = activeEditor.document.uri;
        const edits: vscode.TextEdit[] = [];

        for (let lineNumber = 0; lineNumber < activeEditor.document.lineCount; lineNumber++) {
            const line = activeEditor.document.lineAt(lineNumber);
            if (line.text.length > 73) {
                const range = new vscode.Range(
                    new vscode.Position(lineNumber, 72),
                    new vscode.Position(lineNumber, line.text.length)
                );
                edits.push(vscode.TextEdit.delete(range));
            }
        }

        if (edits.length > 0) {
            const workspaceEdit = new vscode.WorkspaceEdit();
            workspaceEdit.set(uri, edits);
            vscode.workspace.applyEdit(workspaceEdit);
        }
    }

    public static removeComments(activeEditor: vscode.TextEditor): void {
        const uri = activeEditor.document.uri;
        const edits: vscode.TextEdit[] = [];

        // Comment patterns: [regex, removeToLineEnd]
        const commentPatterns: [RegExp, boolean][] = [
            [/\*>/i, true],       // inline comments starting with *>
            [/^.{6}\*/i, false]   // comments in column 7
        ];

        for (let lineNumber = 0; lineNumber < activeEditor.document.lineCount; lineNumber++) {
            const line = activeEditor.document.lineAt(lineNumber).text;

            for (const [regex, removeToLineEnd] of commentPatterns) {
                const match = regex.exec(line);
                if (match) {
                    const startPos = new vscode.Position(lineNumber, match.index);
                    const endPos = removeToLineEnd
                        ? new vscode.Position(lineNumber, line.length)
                        : new vscode.Position(lineNumber + 1, 0);

                    edits.push(vscode.TextEdit.delete(new vscode.Range(startPos, endPos)));
                    break; // stop after the first matching pattern per line
                }
            }
        }

        if (edits.length > 0) {
            const workspaceEdit = new vscode.WorkspaceEdit();
            workspaceEdit.set(uri, edits);
            vscode.workspace.applyEdit(workspaceEdit);
        }
    }

    private static isValidKeywordOrStorageKeyword(languageId: string, keyword: string): boolean {
        const keywordLower = keyword.toLowerCase();
        const keywordDict = getCOBOLKeywordDictionary(languageId);

        return keywordDict.has(keywordLower) 
            || cobolStorageKeywordDictionary.has(keywordLower) 
            || cobolRegistersDictionary.has(keywordLower);
    }

    public static foldTokenLine(text: string, current: ICOBOLSourceScanner | undefined, action: FoldAction, foldConstantsToUpper: boolean, languageid: string, settings: ICOBOLSettings, defaultFoldStyle: intellisenseStyle): string {
        let newtext = text;
        const args: string[] = [];

        SplitTokenizer.splitArgument(text, args);
        const textLower = text.toLowerCase();
        let lastPos = args.length > 1 ? textLower.indexOf(args[0].toLowerCase()) : 0;
        let foldstyle: intellisenseStyle = defaultFoldStyle; //settings.intellisense_style;
        for (let ic = 0; ic < args.length; ic++) {
            let arg = args[ic];

            if (arg.startsWith("\"") && arg.endsWith("\"")) {
                lastPos += arg.length;
                continue;
            }

            if (arg.endsWith(".")) {
                arg = arg.substr(0, arg.length - 1);
            }

            const argLower = arg.toLowerCase();
            const ipos = textLower.indexOf(argLower, lastPos);
            let actionIt = false;

            switch (action) {
                case FoldAction.PerformTargets:
                    if (current !== undefined) {
                        actionIt = current.sections.has(argLower);
                        if (actionIt === false) {
                            actionIt = current.paragraphs.has(argLower);
                            foldstyle = VSCustomIntelliseRules.Default.findCustomIStyle(settings, argLower, foldstyle);
                        }
                    }
                    break;

                case FoldAction.ConstantsOrVariables:
                    if (current !== undefined) {
                        actionIt = current.constantsOrVariables.has(argLower);
                        if (actionIt) {
                            foldstyle = VSCustomIntelliseRules.Default.findCustomIStyle(settings, arg, foldstyle);
                            if (foldConstantsToUpper) {
                                const cvars = current.constantsOrVariables.get(argLower);
                                if (cvars !== undefined) {
                                    for (const cvar of cvars) {
                                        if (cvar.tokenType === COBOLTokenStyle.Constant) {
                                            foldstyle = intellisenseStyle.UpperCase;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;

                case FoldAction.Keywords:
                    actionIt = VSCOBOLUtils.isValidKeywordOrStorageKeyword(languageid, argLower);
                    if (actionIt) {
                        foldstyle = VSCustomIntelliseRules.Default.findCustomIStyle(settings, argLower, foldstyle);
                    }
                    break;
            }

            if (actionIt && foldstyle !== undefined) {
                switch (foldstyle) {
                    case intellisenseStyle.LowerCase:
                        {
                            if (argLower !== arg) {
                                const tmpline = newtext.substr(0, ipos) + argLower + newtext.substr(ipos + arg.length);
                                newtext = tmpline;
                                lastPos += arg.length;
                            }
                        }
                        break;
                    case intellisenseStyle.UpperCase:
                        {
                            const argUpper = arg.toUpperCase();
                            if (argUpper !== arg) {
                                const tmpline = newtext.substr(0, ipos) + argUpper + newtext.substr(ipos + arg.length);
                                newtext = tmpline;
                                lastPos += arg.length;
                            }
                        }
                        break;
                    case intellisenseStyle.CamelCase:
                        {
                            const camelArg = SourceScannerUtils.camelize(arg);
                            if (camelArg !== arg) {
                                const tmpline = newtext.substr(0, ipos) + camelArg + newtext.substr(ipos + arg.length);
                                newtext = tmpline;
                                lastPos += arg.length;
                            }
                        }
                        break;
                }
            }
        }

        // has it changed?
        if (newtext !== text) {
            return newtext;
        }
        return text;
    }

    public static foldToken(
        externalFeatures: IExternalFeatures,
        settings: ICOBOLSettings,
        activeEditor: vscode.TextEditor,
        action: FoldAction,
        languageid: string,
        defaultFoldStyle: intellisenseStyle): void {
        const uri = activeEditor.document.uri;

        const current: ICOBOLSourceScanner | undefined = VSCOBOLSourceScanner.getCachedObject(activeEditor.document, settings);
        if (current === undefined) {
            VSLogger.logMessage(`Unable to fold ${externalFeatures}, as it is has not been parsed`);
            return;
        }
        const file = current.sourceHandler;
        const edits = new vscode.WorkspaceEdit();

        // traverse all the lines
        for (let l = 0; l < file.getLineCount(); l++) {
            const text = file.getLine(l, false);
            if (text === undefined) {
                break;      // eof
            }

            const newtext = VSCOBOLUtils.foldTokenLine(text, current, action, settings.format_constants_to_uppercase, languageid, settings, defaultFoldStyle);

            // one edit per line to avoid the odd overlapping error
            if (newtext !== text) {
                const startPos = new vscode.Position(l, 0);
                const endPos = new vscode.Position(l, newtext.length);
                const range = new vscode.Range(startPos, endPos);
                edits.replace(uri, range, newtext);
            }
        }
        vscode.workspace.applyEdit(edits);
    }

    private static readonly storageAlignItems: string[] = [
        "picture",
        "pic",
        "usage",
        "binary-char",
        "binary-double",
        "binary-long",
        "binary-short",
        "boolean",
        "character",
        "comp-1",
        "comp-2",
        "comp-3",
        "comp-4",
        "comp-5",
        "comp-n",
        "comp-x",
        "comp",
        "computational-1",
        "computational-2",
        "computational-3",
        "computational-4",
        "computational-5",
        "computational-n",
        "computational-x",
        "computational",
        "conditional-value",
        "constant",
        "decimal",
        "external",
        "float-long",
        "float-short",
        "signed-int",
        "signed-long",
        "signed-short",
        "value",
        "as",
        "redefines",
        "renames",
        "pointer-32",
        "pointer"
    ];

    public static getStorageItemPosition(line: string): number {
        for (const storageAlignItem of this.storageAlignItems) {
            const pos = line.toLowerCase().indexOf(" " + storageAlignItem);
            if (pos !== -1) {
                const afterCharPos = 1 + pos + storageAlignItem.length;
                const afterChar = line.charAt(afterCharPos);
                if (afterChar === " " || afterChar === ".") {
                    return pos + 1;
                }
                // VSLogger.logMessage(`afterChar is [${afterChar}]`);
            }
        }
        return -1;
    }

    private static getAlignItemFromSelections(
        editor: vscode.TextEditor,
        selections: readonly vscode.Selection[],
        style: AlignStyle
    ): number {
        let firstPos = -1;
        let leftPos = -1;
        let rightPos = -1;

        for (const sel of selections) {
            for (let lineNum = sel.start.line; lineNum <= sel.end.line; lineNum++) {
                const lineText = editor.document.lineAt(lineNum).text.trimEnd();
                const pos = this.getStorageItemPosition(lineText);

                if (firstPos === -1) firstPos = pos;
                if (leftPos === -1 || pos < leftPos) leftPos = pos;
                if (pos > rightPos) rightPos = pos;
            }
        }

        switch (style) {
            case AlignStyle.First: return firstPos;
            case AlignStyle.Left: return leftPos;
            case AlignStyle.Right: return rightPos;
            case AlignStyle.Center: return Math.trunc((leftPos + rightPos) / 2);
        }

        return -1; // fallback if style is undefined
    }

    public static alignStorage(style: AlignStyle): void {
        const editor = vscode.window.activeTextEditor;
        if (!editor) return;

        const selections = editor.selections;
        const targetPos = VSCOBOLUtils.getAlignItemFromSelections(editor, selections, style);
        if (targetPos === -1) return;

        editor.edit(edits => {
            for (const sel of selections) {
                for (let lineNum = sel.start.line; lineNum <= sel.end.line; lineNum++) {
                    const lineText = editor.document.lineAt(lineNum).text;
                    const currentPos = this.getStorageItemPosition(lineText);
                    if (currentPos === -1 || currentPos === targetPos) continue;

                    // Align the storage area without creating multiple substrings
                    const leftLength = targetPos - 1;
                    let leftPart = lineText.slice(0, currentPos).trimEnd();
                    if (leftPart.length < leftLength) {
                        leftPart = leftPart.padEnd(leftLength, " ");
                    }

                    const rightPart = lineText.slice(currentPos).trimStart();
                    const newText = `${leftPart} ${rightPart}`;

                    edits.replace(
                        new vscode.Range(
                            new vscode.Position(lineNum, 0),
                            new vscode.Position(lineNum, lineText.length)
                        ),
                        newText
                    );
                }
            }
        });
    }

    public static async changeDocumentId(fromID: string, toID: string): Promise<void> {
        const editors = vscode.window.visibleTextEditors.filter(editor => editor.document.languageId === fromID);
        for (const editor of editors) {
            await vscode.languages.setTextDocumentLanguage(editor.document, toID);
        }
    }

    public static enforceFileExtensions(
        settings: ICOBOLSettings,
        activeEditor: vscode.TextEditor,
        externalFeatures: IExternalFeatures,
        verbose: boolean,
        requiredLanguage: string
    ): void {
        const filesConfig = vscode.workspace.getConfiguration("files");

        // Handle Micro Focus suppression prompt
        setMicroFocusSuppressFileAssociationsPrompt(
            settings,
            requiredLanguage === ExtensionDefaults.microFocusCOBOLLanguageId
        );

        // Get current file associations
        const filesAssociationsConfig = filesConfig.get<{ [key: string]: string }>("associations") ?? {};
        let updateRequired = false;

        const fileAssocMap = new Map<string, string>();
        let logHeaderPrinted = false;

        for (const [assoc, assocTo] of Object.entries(filesAssociationsConfig)) {
            fileAssocMap.set(assoc, assocTo);

            // Adjust Micro Focus associations if needed
            if (
                requiredLanguage !== ExtensionDefaults.microFocusCOBOLLanguageId &&
                assocTo === ExtensionDefaults.microFocusCOBOLLanguageId
            ) {
                filesAssociationsConfig[assoc] = requiredLanguage;
                updateRequired = true;
            }

            // Log associations once if verbose
            if (verbose && !logHeaderPrinted) {
                externalFeatures.logMessage("Active file associations:");
                logHeaderPrinted = true;
            }
            if (verbose) {
                externalFeatures.logMessage(` ${assoc} = ${assocTo}`);
            }
        }

        // Enforce required language for program extensions
        for (const ext of settings.program_extensions) {
            const key = `*.${ext}`;
            const assocTo = fileAssocMap.get(key);

            if (assocTo && assocTo !== requiredLanguage) {
                if (verbose && !VSExtensionUtils.isKnownCOBOLLanguageId(settings, assocTo)) {
                    externalFeatures.logMessage(` WARNING: ${ext} is associated with ${assocTo}`);
                }
                filesAssociationsConfig[key] = requiredLanguage;
                updateRequired = true;
            } else if (!assocTo) {
                filesAssociationsConfig[key] = requiredLanguage;
                updateRequired = true;
            }
        }

        // Update workspace or global configuration if needed
        if (updateRequired) {
            const target = VSWorkspaceFolders.get(settings)
                ? vscode.ConfigurationTarget.Workspace
                : vscode.ConfigurationTarget.Global;
            filesConfig.update("associations", filesAssociationsConfig, target);
        }
    }

    public static padTo72(): void {
        const editor = vscode.window.activeTextEditor;
        if (!editor) return;

        const lineNumber = editor.selection.start.line;
        const lineText = editor.document.lineAt(lineNumber).text;

        editor.edit(editBuilder => {
            const lineRange = new vscode.Range(
                new vscode.Position(lineNumber, 0),
                new vscode.Position(lineNumber, lineText.length)
            );

            const paddedLine = lineText.padEnd(72, " ");
            editBuilder.replace(lineRange, paddedLine);
        });
    }

    private static replaceSelectionWith(editor: vscode.TextEditor, transform: (text: string) => string) {
        const selection = editor.selection;
        const text = editor.document.getText(selection);
        const transformed = transform(text);

        editor.edit(editBuilder => {
            editBuilder.replace(selection, transformed);
        });
    }

    public static selectionToHEX(cobolify: boolean) {
        const editor = vscode.window.activeTextEditor;
        if (!editor) return;

        this.replaceSelectionWith(editor, text => VSCOBOLUtils.a2hex(text, cobolify));
    }

    public static selectionHEXToASCII() {
        const editor = vscode.window.activeTextEditor;
        if (!editor) return;

        this.replaceSelectionWith(editor, text => VSCOBOLUtils.hex2a(text));
    }

    public static selectionToNXHEX(cobolify: boolean) {
        const editor = vscode.window.activeTextEditor;
        if (!editor) return;

        this.replaceSelectionWith(editor, text => VSCOBOLUtils.a2nx(text, cobolify));
    }

    private static isControlCode(code: number) {
        return ((code >= 0x0000 && code <= 0x001f) || (code >= 0x007f && code <= 0x009f));
    }
    
    public static a2nx(str: string, cobolify: boolean): string {
        let hexString = "";

        for (const chr of str) {
            // Convert character to big-endian UTF-16 bytes
            const chrBuffer = Buffer.from(chr, "utf16le").reverse();

            for (const byte of chrBuffer) {
                hexString += byte.toString(16).toUpperCase().padStart(2, "0");
            }
        }

        return cobolify ? `NX"${hexString}"` : hexString;
    }

    public static nxhex2a(hexString: string): string {
        // Remove NX"..." or NX'...' wrapper if present
        if (
            hexString.length >= 4 &&
            hexString[0].toUpperCase() === "N" &&
            hexString[1].toUpperCase() === "X" &&
            (hexString[2] === '"' || hexString[2] === "'") &&
            (hexString.endsWith('"') || hexString.endsWith("'"))
        ) {
            hexString = hexString.slice(3, -1);
        }

        const buf = Buffer.from(hexString, "hex");
        let result = "";

        // Process in chunks to avoid argument limit in fromCharCode
        const CHUNK_SIZE = 0x8000; // 32k characters
        for (let i = 0; i < buf.length; i += CHUNK_SIZE * 2) {
            const chunkLength = Math.min(CHUNK_SIZE, (buf.length - i) / 2);
            const chars = new Array(chunkLength);
            for (let j = 0; j < chunkLength; j++) {
                chars[j] = buf.readUInt16BE(i + j * 2);
            }
            result += String.fromCharCode(...chars);
        }

        return result;
    }

    public static a2hex(str: string, cobolify: boolean): string {
        const chunkSize = 1024; // process in chunks to avoid memory issues
        const hexChunks: string[] = [];

        for (let offset = 0; offset < str.length; offset += chunkSize) {
            const chunk = str.slice(offset, offset + chunkSize);
            const hexChunk = Array.from(chunk, c =>
                c.charCodeAt(0).toString(16).padStart(2, "0").toUpperCase()
            ).join("");
            hexChunks.push(hexChunk);
        }

        const hexString = hexChunks.join("");
        return cobolify ? `X"${hexString}"` : hexString;
    }

    public static hex2a(hex: string): string {
        // Remove COBOL X"" or x'' wrappers
        if (hex.match(/^[Xx]["'][\s\S]*["']$/)) {
            hex = hex.slice(2, -1);
        }

        // If hex string has odd length, return as-is
        if (hex.length % 2 !== 0) {
            return hex;
        }

        const chars: string[] = [];
        for (let i = 0; i < hex.length; i += 2) {
            const hexPair = hex.slice(i, i + 2);
            const code = parseInt(hexPair, 16);

            if (this.isControlCode(code)) {
                chars.push(`[${hexPair}]`);
            } else {
                chars.push(String.fromCharCode(code));
            }
        }

        return chars.join("");
    }

    public static setupFilePaths(settings: ICOBOLSettings) {
        const invalidSearchDirectory: string[] = [];
        const fileSearchDirectory: string[] = settings.file_search_directory;
        const perfileCopybookDirs: string[] = settings.perfile_copybookdirs;
        const copybookDirs = settings.copybookdirs;
        const wsFolders = VSWorkspaceFolders.get(settings);

        // Process default copybook directories
        this.processCopybookDirs(copybookDirs, settings, fileSearchDirectory, invalidSearchDirectory, wsFolders);

        // Process workspace folders and nested copybook directories
        if (wsFolders) {
            this.processWorkspaceFolders(wsFolders, copybookDirs, settings, fileSearchDirectory, invalidSearchDirectory);
        }

        // Remove duplicates
        settings.file_search_directory = Array.from(new Set(fileSearchDirectory));
        settings.invalid_copybookdirs = Array.from(new Set(invalidSearchDirectory));

        // Logging
        this.logFilePaths(wsFolders, fileSearchDirectory, perfileCopybookDirs, invalidSearchDirectory);

        // Populate default symbols and copybooks if recursive search is enabled
        if (settings.maintain_metadata_recursive_search) {
            VSCOBOLUtils.populateDefaultCallableSymbols(settings, true);
            VSCOBOLUtils.populateDefaultCopyBooks(settings, true);
        }
    }

    /** Helper: Process top-level copybook directories */
    private static processCopybookDirs(
        copybookDirs: string[],
        settings: ICOBOLSettings,
        fileSearchDirectory: string[],
        invalidSearchDirectory: string[],
        wsFolders?: readonly vscode.WorkspaceFolder[]
    ) {
        VSCOBOLUtils.processCopybookAndWorkspaceDirs(copybookDirs, settings, fileSearchDirectory, invalidSearchDirectory, wsFolders);
    }

    private static processWorkspaceFolders(
        wsFolders: readonly vscode.WorkspaceFolder[],
        copybookDirs: string[],
        settings: ICOBOLSettings,
        fileSearchDirectory: string[],
        invalidSearchDirectory: string[]
    ) {
        VSCOBOLUtils.processCopybookAndWorkspaceDirs(copybookDirs, settings, fileSearchDirectory, invalidSearchDirectory, wsFolders);
    }

    private static processCopybookAndWorkspaceDirs(
        copybookDirs: string[],
        settings: ICOBOLSettings,
        fileSearchDirectory: string[],
        invalidSearchDirectory: string[],
        wsFolders?: readonly vscode.WorkspaceFolder[]
    ) {
        const validateAndAddDir = (dir: string, allowNetworkLogging = false) => {
            try {
                const isDirect = COBOLFileUtils.isDirectPath(dir);
                const isNetwork = COBOLFileUtils.isNetworkPath(dir);

                if (workspace.isTrusted === false && (isDirect || isNetwork)) {
                    invalidSearchDirectory.push(dir);
                    return;
                }

                if (settings.disable_unc_copybooks_directories && isNetwork) {
                    VSLogger.logMessage(`Copybook directory ${dir} is marked invalid (UNC path).`);
                    invalidSearchDirectory.push(dir);
                    return;
                }

                if (isDirect) {
                    if (allowNetworkLogging && wsFolders && !VSCOBOLFileUtils.isPathInWorkspace(dir, settings) && isNetwork) {
                        VSLogger.logMessage(`Directory ${dir} for performance should be part of the workspace.`);
                    }

                    const startTime = VSExternalFeatures.performance_now();
                    if (COBOLFileUtils.isDirectory(dir)) {
                        const elapsed = VSExternalFeatures.performance_now() - startTime;
                        if (elapsed <= 2000) {
                            fileSearchDirectory.push(dir);
                        } else {
                            VSLogger.logMessage(`Slow copybook directory dropped: ${dir} (${elapsed.toFixed(2)}ms)`);
                            invalidSearchDirectory.push(dir);
                        }
                    } else {
                        invalidSearchDirectory.push(dir);
                    }
                }
            } catch (e) {
                // Ignore errors
            }
        };

        // Process top-level copybook directories
        for (const dir of copybookDirs) {
            validateAndAddDir(dir, true);
        }

        // Process workspace folders if provided
        if (wsFolders) {
            for (const folder of wsFolders) {
                const folderPath = folder.uri.fsPath;
                fileSearchDirectory.push(folderPath);

                for (const extDir of copybookDirs) {
                    if (!COBOLFileUtils.isDirectPath(extDir)) {
                        const fullPath = path.join(folderPath, extDir);
                        validateAndAddDir(fullPath);
                    }
                }
            }
        }
    }

    /** Helper: Log file paths for workspace, search, per-file, and invalid directories */
    private static logFilePaths(
        wsFolders: readonly vscode.WorkspaceFolder[] | undefined,
        fileSearchDirectory: string[],
        perfileCopybookDirs: string[],
        invalidSearchDirectory: string[]
    ) {
        if (wsFolders && wsFolders.length > 0) {
            VSLogger.logMessage("Workspace Folders:");
            wsFolders.forEach(folder => VSLogger.logMessage(`  => ${folder.name} @ ${folder.uri.fsPath}`));
        }

        if (fileSearchDirectory.length > 0) {
            VSLogger.logMessage("Combined Workspace and CopyBook Folders to search:");
            fileSearchDirectory.forEach(dir => VSLogger.logMessage(`  => ${dir}`));
        }

        if (perfileCopybookDirs.length > 0) {
            VSLogger.logMessage(`Per File CopyBook directories (${perfileCopybookDirs.length}):`);
            perfileCopybookDirs.forEach(dir => VSLogger.logMessage(`  => ${dir}`));
        }

        if (invalidSearchDirectory.length > 0) {
            VSLogger.logMessage(`Invalid CopyBook directories (${invalidSearchDirectory.length}):`);
            invalidSearchDirectory.forEach(dir => VSLogger.logMessage(`  => ${dir}`));
        }

        VSLogger.logMessage("");
    }

    public static setupUrlPathsSync(settings: ICOBOLSettings) {
        async() => {
            await VSCOBOLUtils.setupUrlPaths(settings);
        }
    }

    public static async setupUrlPaths(settings: ICOBOLSettings) {
        const invalidSearchDirectory: Set<string> = new Set(settings.invalid_copybookdirs);
        const URLSearchDirectory: Set<string> = new Set(VSExternalFeatures.getURLCopyBookSearchPath());
        const copybookDirs = settings.copybookdirs;

        const wsURLs = VSWorkspaceFolders.getFiltered("", settings);
        if (!wsURLs || wsURLs.length === 0) {
            return;
        }

        for (const folder of wsURLs) {
            // Add workspace folder URL
            URLSearchDirectory.add(folder.uri.toString());

            // Add extra directories under workspace folder
            for (const extDir of copybookDirs) {
                try {
                    if (!COBOLFileUtils.isDirectPath(extDir)) {
                        const dirUrl = `${folder.uri.toString()}/${extDir}`;
                        const stat = await vscode.workspace.fs.stat(vscode.Uri.parse(dirUrl));
                        if (stat.type & vscode.FileType.Directory) {
                            URLSearchDirectory.add(dirUrl);
                        } else {
                            invalidSearchDirectory.add(`URL as ${dirUrl}`);
                        }
                    }
                } catch (e) {
                    // Ignore errors
                }
            }
        }

        // Logging
        VSLogger.logMessage("  Workspace Folders (URLs):");
        wsURLs.forEach(folder => VSLogger.logMessage(`   => ${folder.name} @ ${folder.uri.fsPath}`));

        if (URLSearchDirectory.size > 0) {
            VSLogger.logMessage("  Combined Workspace and CopyBook Folders to search (URL):");
            URLSearchDirectory.forEach(dir => VSLogger.logMessage(`   => ${dir}`));
        }

        if (invalidSearchDirectory.size > 0) {
            VSLogger.logMessage(`  Invalid CopyBook directories (${invalidSearchDirectory.size}):`);
            invalidSearchDirectory.forEach(dir => VSLogger.logMessage(`   => ${dir}`));
        }

        VSLogger.logMessage("");
    }

    public static getDebugConfig(workspaceFolder: vscode.WorkspaceFolder, debugFile: string): vscode.DebugConfiguration {
        return {
            name: "Debug COBOL",
            type: "cobol",
            request: "launch",
            cwd: workspaceFolder.uri.fsPath,
            program: debugFile,
            stopOnEntry: true
        };
    }

    public static runOrDebug(fsPath: string, debug: boolean): void {
        if (!commandTerminal) {
            commandTerminal = vscode.window.createTerminal(commandTerminalName);
        }

        const fsDir = path.dirname(fsPath);
        let runner = "";
        let runnerArgs = "";

        const showTerminalAndRun = (cmd: string) => {
            commandTerminal!.show(true);
            commandTerminal!.sendText(cmd);
        };

        // Handle ACU COBOL files
        if (fsPath.endsWith("acu")) {
            runner = COBOLFileUtils.isWin32 ? "wrun32" : "runcbl";
            runnerArgs = debug ? "-d " : "";
            showTerminalAndRun(`${runner} ${runnerArgs}${fsPath}`);
            return;
        }

        // Handle .NET and other runtime files
        if (fsPath.endsWith("int") || fsPath.endsWith("gnt") || fsPath.endsWith("so") || fsPath.endsWith("dll")) {
            const mfExtension = vscode.extensions.getExtension(ExtensionDefaults.rocketCOBOLExtension);

            if (!mfExtension) {
                if (COBOLFileUtils.isWin32) {
                    runner = "run";
                    runnerArgs = debug ? "(+A) " : "";
                } else {
                    runner = debug ? "anim" : "cobrun";
                }
                showTerminalAndRun(`${runner} ${runnerArgs}${fsPath}`);
                return;
            }

            const workspacePath = VSCOBOLFileUtils.getBestWorkspaceFolder(fsDir);
            if (workspacePath) {
                vscode.debug.startDebugging(workspacePath, VSCOBOLUtils.getDebugConfig(workspacePath, fsPath));
            }
        }
    }
}

