import * as vscode from "vscode";
import { SharedSourceReferences } from "./cobolsourcescanner";
import { CodeActionProvider, CodeAction } from "vscode";
import { ICOBOLSettings } from "./iconfiguration";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { CobolLinterProviderSymbols } from "./externalfeatures";
import { TextLanguage, VSExtensionUtils } from "./vsextutis";
import { VSCOBOLUtils } from "./vscobolutils";
import { VSCOBOLFileUtils } from "./vsfileutils";
import path from "path";
import { VSCOBOLConfiguration, VSCOBOLEditorConfiguration } from "./vsconfiguration";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

export class CobolLinterActionFixer implements CodeActionProvider {

    // Helper to create code actions
    private createCodeAction(title: string, commandName: string, document: vscode.TextDocument, lineOrOffset: number, codeArg: string, diagnostic: vscode.Diagnostic): vscode.CodeAction {
        return {
            title,
            diagnostics: [diagnostic],
            command: {
                title,
                command: commandName,
                arguments: [document.uri, lineOrOffset, codeArg],
            },
            kind: vscode.CodeActionKind.QuickFix
        };
    }

    provideCodeActions(document: vscode.TextDocument, range: vscode.Range | vscode.Selection, context: vscode.CodeActionContext, token: vscode.CancellationToken): vscode.ProviderResult<(vscode.Command | vscode.CodeAction)[]> {
        const codeActions: CodeAction[] = [];

        for (const diagnostic of context.diagnostics) {
            if (!diagnostic.code) continue;

            const codeMsg = diagnostic.code.toString();
            const startOfLine = document.offsetAt(new vscode.Position(diagnostic.range.start.line, 0));

            if (codeMsg.startsWith(CobolLinterProviderSymbols.NotReferencedMarker_internal)) {
                const insertCode = codeMsg.replace(CobolLinterProviderSymbols.NotReferencedMarker_internal, CobolLinterProviderSymbols.NotReferencedMarker_external);
                codeActions.push(this.createCodeAction(
                    `Add COBOL lint ignore comment for '${diagnostic.message}'`,
                    "cobolplugin.insertIgnoreCommentLine",
                    document,
                    startOfLine,
                    insertCode,
                    diagnostic
                ));
            } else if (codeMsg.startsWith(CobolLinterProviderSymbols.CopyBookNotFound)) {
                const insertCode = codeMsg.replace(CobolLinterProviderSymbols.CopyBookNotFound, "").trim();
                codeActions.push(this.createCodeAction(
                    `Find Copybook ${insertCode}`,
                    "cobolplugin.findCopyBookDirectory",
                    document,
                    startOfLine,
                    insertCode,
                    diagnostic
                ));
            } else if (codeMsg.startsWith(CobolLinterProviderSymbols.PortMessage)) {
                let portChangeLine = codeMsg.replace(CobolLinterProviderSymbols.PortMessage, "").trim();
                portChangeLine = portChangeLine.startsWith("$") ? "      " + portChangeLine : "       " + portChangeLine;

                if (portChangeLine.trim().length > 0) {
                    codeActions.push(this.createCodeAction(
                        `Change line to ${portChangeLine.trim()}`,
                        "cobolplugin.portCodeCommandLine",
                        document,
                        diagnostic.range.start.line,
                        portChangeLine,
                        diagnostic
                    ));
                }
            }
        }

        return codeActions;
    }

    public async insertIgnoreCommentLine(docUri: vscode.Uri, offset: number, code: string): Promise<void> {
        await vscode.window.showTextDocument(docUri);
        const editor = vscode.window.activeTextEditor;
        if (editor && code) {
            const pos = editor.document.positionAt(offset);
            await editor.edit(edit => edit.insert(pos, "      *> cobol-lint " + code + "\n"));
        }
    }

    public async portCodeCommandLine(docUri: vscode.Uri, lineNumber: number, code: string): Promise<void> {
        await vscode.window.showTextDocument(docUri);
        const editor = vscode.window.activeTextEditor;
        if (editor && code) {
            const line = editor.document.lineAt(lineNumber);
            await editor.edit(edit => edit.replace(line.range, code));
        }
    }

    private knownExternalCopybooks = new Map<string, string>([
        ["cblproto.cpy", "$COBCPY"],
        ["cbltypes.cpy", "$COBCPY"],
        ["windows.cpy", "$COBCPY"],
        ["mfunit.cpy", "$COBCPY"],
        ["mfunit_prototypes.cpy", "$COBCPY"]
    ]);

    private uniqByFilter<T>(array: T[]) {
        return array.filter((value, index) => array.indexOf(value) === index);
    }

    public async findCopyBookDirectory(settings: ICOBOLSettings, docUri: vscode.Uri, linenum: number, copybook: string) {
        const fileSearchDirs = settings.config_copybookdirs;
        let update = false;
        let update4found = false;

        if (!fileSearchDirs.includes("$COBCPY") && process.env) {
            const COB = process.env["COBCPY"];
            if (COB && COB.length > 0) {
                fileSearchDirs.push("$COBCPY");
                update = true;
            } else {
                const knownDir = this.knownExternalCopybooks.get(copybook);
                if (knownDir) {
                    fileSearchDirs.push(knownDir);
                    update = true;
                }
            }
        }

        let files = await vscode.workspace.findFiles(`**/${copybook}`);
        if (files.length === 0) {
            const globPattern = VSCOBOLUtils.getCopyBookGlobPatternForPartialName(settings, copybook);
            files = await vscode.workspace.findFiles(globPattern);
        }

        for (const uri of files) {
            const wsFile = VSCOBOLFileUtils.getShortWorkspaceFilename(uri.scheme, uri.fsPath, settings);
            if (wsFile) {
                let dirName = path.dirname(wsFile);
                if (copybook.includes("/")) {
                    const up = copybook.split("/").length - 1;
                    for (let i = 0; i < up; i++) dirName = path.dirname(dirName);
                }
                if (!fileSearchDirs.includes(dirName)) {
                    fileSearchDirs.push(dirName);
                    update = update4found = true;
                }
            }
        }

        if (update) {
            const editorConfig = VSCOBOLEditorConfiguration.getEditorConfig();
            await editorConfig.update("copybookdirs", this.uniqByFilter(fileSearchDirs));
        }

        if (!update4found) {
            await vscode.window.showInformationMessage(`Unable to locate ${copybook} in workspace`);
        } else {
            await vscode.commands.executeCommand("workbench.action.reloadWindow");
        }
    }
}

export class CobolLinterProvider {
    private collection: vscode.DiagnosticCollection;
    private current: ICOBOLSourceScanner | undefined;
    private currentVersion?: number;
    private sourceRefs?: SharedSourceReferences;

    constructor(collection: vscode.DiagnosticCollection) {
        this.collection = collection;
    }

    // Helper to add diagnostics
    private addDiagnostic(diagRefs: Map<string, vscode.Diagnostic[]>, filename: string, range: vscode.Range, message: string, code: string | number | undefined, severity: vscode.DiagnosticSeverity) {
        const diagnostic = new vscode.Diagnostic(range, message, severity);
        diagnostic.code = code;
        diagnostic.tags = [vscode.DiagnosticTag.Unnecessary];

        if (!diagRefs.has(filename)) diagRefs.set(filename, []);
        diagRefs.get(filename)?.push(diagnostic);
    }

    public async updateLinter(document: vscode.TextDocument): Promise<void> {
        const settings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        const linterSev = settings.linter_mark_as_information ? vscode.DiagnosticSeverity.Information : vscode.DiagnosticSeverity.Hint;

        if (!settings.linter || VSExtensionUtils.isSupportedLanguage(document) !== TextLanguage.COBOL) return;
        if (!this.setupCOBOLScannner(document, settings)) return;

        const diagRefs = new Map<string, vscode.Diagnostic[]>();
        this.collection.clear();

        if (!this.current) return;
        const qp = this.current;

        if (this.sourceRefs && !qp.sourceIsCopybook) {
            this.processScannedDocumentForUnusedSymbols(settings, qp, diagRefs, qp.configHandler.linter_unused_paragraphs, qp.configHandler.linter_unused_sections, linterSev);
            if (qp.configHandler.linter_house_standards_rules) {
                this.processParsedDocumentForStandards(qp, diagRefs, linterSev);
            }
        }

        this.processScannerWarnings(qp, settings, diagRefs, linterSev);

        for (const [file, diags] of diagRefs) {
            this.collection.set(vscode.Uri.file(file), diags);
        }
    }

    private processScannerWarnings(qp: ICOBOLSourceScanner, settings: ICOBOLSettings, diagRefs: Map<string, vscode.Diagnostic[]>, linterSev: vscode.DiagnosticSeverity) {
        const processWarning = (fileSymbol: any, prefix: string, message: string) => {
            if (fileSymbol.filename !== undefined && fileSymbol.linenum !== undefined) {
                const range = new vscode.Range(
                    new vscode.Position(fileSymbol.linenum, 0),
                    new vscode.Position(fileSymbol.linenum, 0)
                );
                this.addDiagnostic(diagRefs, fileSymbol.filename, range, message, prefix + (fileSymbol.messageOrMissingFile || fileSymbol.replaceLine), linterSev);
            }
        };

        if (!settings.linter_ignore_missing_copybook && qp.diagMissingFileWarnings.size !== 0) {
            for (const [msg, fileSymbol] of qp.diagMissingFileWarnings) {
                processWarning(fileSymbol, CobolLinterProviderSymbols.CopyBookNotFound, msg);
            }
        }

        for (const fileSymbol of qp.portWarnings) {
            processWarning(fileSymbol, CobolLinterProviderSymbols.PortMessage, fileSymbol.message);
        }

        for (const fileSymbol of qp.generalWarnings) {
            processWarning(fileSymbol, CobolLinterProviderSymbols.GeneralMessage, fileSymbol.messageOrMissingFile);
        }
    }
    private processParsedDocumentForStandards(qp: ICOBOLSourceScanner, diagRefs: Map<string, vscode.Diagnostic[]>, linterSev: vscode.DiagnosticSeverity) {
        if (!this.sourceRefs) return;

        const standards = qp.configHandler.linter_house_standards_rules;
        const standardsMap = new Map<string, string>();
        const ruleRegexMap = new Map<string, RegExp>();

        for (const standard of standards) {
            const [section, rule] = standard.split("=", 2);
            standardsMap.set(section.toLowerCase(), rule);
        }

        for (const [key, tokens] of qp.constantsOrVariables) {
            for (const variable of tokens) {
                const token = variable.token;
                if (!token?.inSection || token.tokenNameLower === "filler") continue;

                const rule = standardsMap.get(token.inSection.tokenNameLower);
                if (!rule) continue;

                let regex = ruleRegexMap.get(token.inSection.tokenNameLower);
                if (!regex) {
                    regex = this.makeRegex(rule);
                    if (!regex) continue;
                    ruleRegexMap.set(token.inSection.tokenNameLower, regex);
                }

                if (!regex.test(token.tokenName)) {
                    const range = new vscode.Range(new vscode.Position(token.startLine, token.startColumn),
                        new vscode.Position(token.startLine, token.startColumn + token.tokenName.length));
                    this.addDiagnostic(diagRefs, token.filename, range, `${key} breaks house standards rule for ${token.inSection.tokenNameLower} section`, undefined, linterSev);
                }
            }
        }
    }

    private processScannedDocumentForUnusedSymbols(settings: ICOBOLSettings, qp: ICOBOLSourceScanner, diagRefs: Map<string, vscode.Diagnostic[]>, processParas: boolean, processSections: boolean, linterSev: vscode.DiagnosticSeverity) {
        if (!this.sourceRefs) return;

        const sharedRefs = this.sourceRefs;

        const addUnusedDiagnostic = (token: any, key: string, kind: string) => {
            const range = new vscode.Range(new vscode.Position(token.startLine, token.startColumn),
                new vscode.Position(token.startLine, token.startColumn + token.tokenName.length));
            this.addDiagnostic(diagRefs, token.filename, range, `${key} ${kind} is not referenced`, CobolLinterProviderSymbols.NotReferencedMarker_internal + " " + key, linterSev);
        };

        if (processParas) {
            for (const [key, token] of qp.paragraphs) {
                if (sharedRefs.ignoreUnusedSymbol.has(key.toLowerCase())) continue;
                const refCount = qp.sourceReferences.getReferenceInformation4targetRefs(key.toLowerCase(), qp.sourceFileId, token.startLine, token.startColumn)[1];
                if (refCount === 0) addUnusedDiagnostic(token, key, "paragraph");
            }
        }

        if (processSections) {
            for (const [key, token] of qp.sections) {
                if (!token.inProcedureDivision || sharedRefs.ignoreUnusedSymbol.has(key.toLowerCase())) continue;

                const refCount = qp.sourceReferences.getReferenceInformation4targetRefs(key.toLowerCase(), qp.sourceFileId, token.startLine, token.startColumn)[1];
                if (refCount === 0) {
                    let ignore = false;
                    if (settings.linter_ignore_section_before_entry) {
                        const nextLine = qp.sourceHandler.getLine(token.startLine + 1, false);
                        if (nextLine?.toLowerCase().includes("entry")) ignore = true;
                    }
                    if (!ignore) addUnusedDiagnostic(token, key, "section");
                }
            }
        }
    }

    private setupCOBOLScannner(document: vscode.TextDocument, settings: ICOBOLSettings): boolean {
        if (this.current && this.current.filename !== document.fileName) this.current = undefined;

        if (!this.current || this.currentVersion !== document.version) {
            this.current = VSCOBOLSourceScanner.getCachedObject(document, settings);
            if (this.current) this.sourceRefs = this.current.sourceReferences;
            this.currentVersion = document.version;
            return true;
        }
        return false;
    }

    private makeRegex(partialRegEx: string): RegExp | undefined {
        try {
            return new RegExp("^" + partialRegEx + "$", "i");
        } catch {
            return undefined;
        }
    }
}