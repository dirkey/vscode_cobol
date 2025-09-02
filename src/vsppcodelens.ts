/* eslint-disable @typescript-eslint/no-explicit-any */
import * as vscode from "vscode";
import { COBOLToken } from "./cobolsourcescanner";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { ICOBOLSettings } from "./iconfiguration";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { ExtensionDefaults } from "./extensionDefaults";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

export class VSPPCodeLens implements vscode.CodeLensProvider {
    private _onDidChangeCodeLenses: vscode.EventEmitter<void> = new vscode.EventEmitter<void>();
    public readonly onDidChangeCodeLenses: vscode.Event<void> = this._onDidChangeCodeLenses.event;

    constructor() {
        vscode.workspace.onDidChangeConfiguration(() => {
            this._onDidChangeCodeLenses.fire();
        });
    }

    private scanTargetUse(
        settings: ICOBOLSettings,
        document: vscode.TextDocument,
        lens: vscode.CodeLens[],
        current: ICOBOLSourceScanner,
        target: string,
        targetToken: COBOLToken
    ) {
        if (targetToken.isFromScanCommentsForReferences || targetToken.ignoreInOutlineView) return;

        const refs = current.sourceReferences.targetReferences.get(target);
        if (!refs || !settings.enable_codelens_section_paragraph_references) return;

        const [_, refCount] = current.sourceReferences.getReferenceInformation4targetRefs(
            target,
            current.sourceFileId,
            targetToken.startLine,
            targetToken.startColumn
        );

        if (refCount === 0) return;

        const refCountMsg = refCount === 1 ? "1 reference" : `${refCount} references`;
        const range = new vscode.Range(
            new vscode.Position(targetToken.rangeStartLine, targetToken.rangeStartColumn),
            new vscode.Position(targetToken.rangeEndLine, targetToken.rangeEndColumn)
        );

        const cl = new vscode.CodeLens(range, {
            title: refCountMsg,
            command: "editor.action.findReferences",
            arguments: [document.uri, new vscode.Position(targetToken.rangeStartLine, targetToken.rangeStartColumn)]
        });

        lens.push(cl);
    }

    public provideCodeLenses(
        document: vscode.TextDocument,
        token: vscode.CancellationToken
    ): vscode.ProviderResult<vscode.CodeLens[]> {
        const lens: vscode.CodeLens[] = [];
        const settings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        const current: ICOBOLSourceScanner | undefined = VSCOBOLSourceScanner.getCachedObject(document, settings);

        if (!current) return lens;

        const sourceFileId = current.sourceFileId;

        // CodeLens for variables
        if (settings.enable_codelens_variable_references && current.constantsOrVariables) {
            for (const [varName, vars] of current.constantsOrVariables) {
                for (const currentVar of vars) {
                    const token = currentVar.token;
                    if (token.isFromScanCommentsForReferences || token.ignoreInOutlineView) continue;

                    const [_, refCount] = current.sourceReferences.getReferenceInformation4variables(
                        varName,
                        sourceFileId,
                        token.startLine,
                        token.startColumn
                    );

                    if (refCount === 0) continue;

                    const refMsg = vars.length === 1 ? `${refCount} reference${refCount > 1 ? "s" : ""}` : "View references";
                    const range = new vscode.Range(
                        new vscode.Position(token.rangeStartLine, token.rangeStartColumn),
                        new vscode.Position(token.rangeEndLine, token.rangeEndColumn)
                    );

                    const cl = new vscode.CodeLens(range, {
                        title: refMsg,
                        command: "editor.action.findReferences",
                        arguments: [document.uri, new vscode.Position(token.startLine, token.startColumn)]
                    });

                    lens.push(cl);
                }
            }
        }

        // CodeLens for sections & paragraphs
        if (settings.enable_codelens_section_paragraph_references && current.sourceReferences?.sharedParagraphs) {
            for (const [name, token] of current.sections) {
                this.scanTargetUse(settings, document, lens, current, name, token);
            }
            for (const [name, token] of current.paragraphs) {
                this.scanTargetUse(settings, document, lens, current, name, token);
            }
        }

        // CodeLens for copybook replacement
        if (settings.enable_codelens_copy_replacing) {
            for (const [, cbInfos] of current.copyBooksUsed) {
                for (const cbInfo of cbInfos) {
                    if (!cbInfo.scanComplete || !cbInfo.statementInformation || cbInfo.statementInformation.copyReplaceMap.size === 0) continue;

                    const stmt = cbInfo.statementInformation;
                    const line = document.lineAt(stmt.startLineNumber);
                    const range = new vscode.Range(
                        new vscode.Position(stmt.startLineNumber, 0),
                        new vscode.Position(stmt.startLineNumber, line.text.length)
                    );
                    const cl = new vscode.CodeLens(range);

                    let src = "";
                    let prevSrc = "";
                    let prevMaxLines = 10;

                    if (stmt.sourceHandler) {
                        for (let i = 0; i < stmt.sourceHandler.getLineCount(); i++) {
                            const updatedLine = stmt.sourceHandler.getUpdatedLine(i);
                            src += updatedLine + "\n";
                            if (prevMaxLines-- > 0) prevSrc += updatedLine + "\n";
                        }
                        if (prevMaxLines <= 0) prevSrc += "\n......";
                    }

                    if (src.length > 0) {
                        const arg = `*> Caution: This is an approximation\n*> Original file: ${stmt.fileName}\n${src}`;
                        cl.command = {
                            title: "View copybook replacement",
                            tooltip: prevSrc,
                            command: "cobolplugin.ppcodelenaction",
                            arguments: [arg]
                        };
                        this.resolveCodeLens(cl, token);
                        lens.push(cl);
                    }
                }
            }
        }

        return lens;
    }

    public resolveCodeLens(codeLens: vscode.CodeLens, token: vscode.CancellationToken): vscode.CodeLens {
        return codeLens;
    }

    public static actionCodeLens(arg: string): void {
        vscode.workspace.openTextDocument({ content: arg, language: "text" }).then(doc => {
            vscode.window.showTextDocument(doc).then(editor => {
                if (arg.startsWith("*>")) {
                    vscode.languages.setTextDocumentLanguage(editor.document, ExtensionDefaults.defaultCOBOLLanguage);
                }
            });
        });
    }
}
