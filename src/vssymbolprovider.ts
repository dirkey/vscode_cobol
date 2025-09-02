import * as vscode from "vscode";
import { COBOLTokenStyle } from "./cobolsourcescanner";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { VSLogger } from "./vslogger";
import { outlineFlag } from "./iconfiguration";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { SplitTokenizer } from "./splittoken";
import { VSExternalFeatures } from "./vsexternalfeatures";

function makeRange(startLine: number, startCol: number, endLine: number, endCol: number): vscode.Range {
    return new vscode.Range(new vscode.Position(startLine, startCol), new vscode.Position(endLine, endCol));
}

function makeLocation(uri: vscode.Uri, range: vscode.Range): vscode.Location {
    return new vscode.Location(uri, range);
}

function makeSymbol(
    name: string,
    kind: vscode.SymbolKind,
    container: string,
    uri: vscode.Uri,
    range: vscode.Range
): vscode.SymbolInformation {
    return new vscode.SymbolInformation(name, kind, container, makeLocation(uri, range));
}

export class MFDirectivesSymbolProvider implements vscode.DocumentSymbolProvider {
    public async provideDocumentSymbols(
        document: vscode.TextDocument,
        _canceltoken: vscode.CancellationToken
    ): Promise<vscode.SymbolInformation[]> {
        const symbols: vscode.SymbolInformation[] = [];
        const settings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        if (settings.outline === outlineFlag.Off) return symbols;

        const ownerUri = document.uri;
        const container = "";

        for (let i = 0; i < document.lineCount; i++) {
            const text = document.lineAt(i).text.trimEnd();
            if (text.startsWith("#")) continue;

            if (text.trim() === "@root") {
                symbols.push(makeSymbol("@root", vscode.SymbolKind.Constant, container, ownerUri, makeRange(i, 1, i, text.length)));

                try {
                    const uris = await vscode.workspace.findFiles("**/*/directives.mf");
                    for (const uri of uris) {
                        const relativePath = uri.fsPath.replace(ownerUri.fsPath, "").replace(/^\/+/, "");
                        symbols.push(makeSymbol(relativePath, vscode.SymbolKind.File, container, uri, makeRange(1, 1, 1, 1)));
                    }
                } catch (error) {
                    VSLogger.logException("Failed to find directives.mf files", error as Error);
                }
                continue;
            }

            if (text.startsWith("[") && text.endsWith("]")) {
                const item = text.slice(1, -1);
                let lastLine = i;
                let lastLineLength = text.length;

                for (i = i + 1; i < document.lineCount; i++) {
                    const nextText = document.lineAt(i).text.trimEnd();
                    if (!nextText.trim()) {
                        lastLine = i;
                        lastLineLength = nextText.length;
                        break;
                    }
                    if (nextText.startsWith("[") && nextText.endsWith("]")) {
                        lastLine = i - 1;
                        break;
                    }
                    lastLineLength = nextText.length;
                }

                symbols.push(makeSymbol(item, vscode.SymbolKind.Array, container, ownerUri, makeRange(i, 1, lastLine, lastLineLength)));
            }
        }
        return symbols;
    }
}

export class JCLDocumentSymbolProvider implements vscode.DocumentSymbolProvider {
    public async provideDocumentSymbols(
        document: vscode.TextDocument,
        _canceltoken: vscode.CancellationToken
    ): Promise<vscode.SymbolInformation[]> {
        const symbols: vscode.SymbolInformation[] = [];
        const settings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        if (settings.outline === outlineFlag.Off) return symbols;

        const ownerUri = document.uri;
        const lastLine = document.lineCount;
        const lastLineColumn = document.lineAt(lastLine - 1).text.length;
        let container = "";

        for (let i = 0; i < document.lineCount; i++) {
            const text = document.lineAt(i).text.trimEnd();
            if (text.startsWith("//*")) continue;

            if (text.startsWith("//")) {
                const tokens: string[] = [];
                SplitTokenizer.splitArgument(text.substring(2), tokens);
                const lineTokens = tokens.map(t => t.trim()).filter(Boolean);

                if (lineTokens.length > 1) {
                    const [firstToken, secondToken] = [lineTokens[0], lineTokens[1].toLowerCase()];

                    if (secondToken.includes("job")) {
                        symbols.push(makeSymbol(firstToken, vscode.SymbolKind.Field, container, ownerUri, makeRange(i, 0, lastLine, lastLineColumn)));
                        container = firstToken;
                    }

                    if (secondToken.includes("exec")) {
                        symbols.push(makeSymbol(firstToken, vscode.SymbolKind.Function, container, ownerUri, makeRange(i, 0, i, text.length)));
                    }
                }
            }
        }
        return symbols;
    }
}

export class CobolSymbolInformationProvider implements vscode.DocumentSymbolProvider {
    public async provideDocumentSymbols(
        document: vscode.TextDocument,
        _canceltoken: vscode.CancellationToken
    ): Promise<vscode.SymbolInformation[]> {
        const symbols: vscode.SymbolInformation[] = [];
        const config = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        const outlineLevel = config.outline;

        if (outlineLevel === outlineFlag.Off) return symbols;

        const sf = VSCOBOLSourceScanner.getCachedObject(document, config);
        if (!sf) return symbols;

        const ownerUri = document.uri;
        const includePara = outlineLevel !== outlineFlag.Partial && outlineLevel !== outlineFlag.Skeleton;
        const includeVars = outlineLevel !== outlineFlag.Skeleton;
        const includeSections = outlineLevel !== outlineFlag.Skeleton;

        for (const token of sf.tokensInOrder) {
            try {
                if (token.ignoreInOutlineView) continue;

                const range = makeRange(token.rangeStartLine, token.rangeStartColumn, token.rangeEndLine, token.rangeEndColumn);
                const container = token.parentToken?.description ?? "";

                const add = (kind: vscode.SymbolKind) =>
                    symbols.push(makeSymbol(token.description, kind, container, ownerUri, range));

                switch (token.tokenType) {
                    case COBOLTokenStyle.ClassId:
                    case COBOLTokenStyle.ProgramId: add(vscode.SymbolKind.Class); break;
                    case COBOLTokenStyle.CopyBook:
                    case COBOLTokenStyle.CopyBookInOrOf:
                    case COBOLTokenStyle.File:
                    case COBOLTokenStyle.Region: add(vscode.SymbolKind.File); break;
                    case COBOLTokenStyle.Declaratives:
                    case COBOLTokenStyle.Division:
                    case COBOLTokenStyle.MethodId:
                    case COBOLTokenStyle.Paragraph:
                        if (includePara) add(vscode.SymbolKind.Method);
                        break;
                    case COBOLTokenStyle.Section:
                        if (includeSections) add(vscode.SymbolKind.Method);
                        break;
                    case COBOLTokenStyle.Exec:
                    case COBOLTokenStyle.EntryPoint:
                    case COBOLTokenStyle.FunctionId: add(vscode.SymbolKind.Function); break;
                    case COBOLTokenStyle.EnumId: add(vscode.SymbolKind.Enum); break;
                    case COBOLTokenStyle.InterfaceId: add(vscode.SymbolKind.Interface); break;
                    case COBOLTokenStyle.ValueTypeId: add(vscode.SymbolKind.Struct); break;
                    case COBOLTokenStyle.IgnoreLS: add(vscode.SymbolKind.Null); break;
                    case COBOLTokenStyle.Variable:
                        if (!includeVars || token.tokenNameLower === "filler") break;
                        if (["fd", "sd", "rd", "select"].includes(token.extraInformation1)) {
                            add(vscode.SymbolKind.File);
                        } else if (token.extraInformation1.includes("-GROUP")) {
                            add(vscode.SymbolKind.Struct);
                        } else if (token.extraInformation1.includes("88")) {
                            add(vscode.SymbolKind.EnumMember);
                        } else if (token.extraInformation1.includes("-OCCURS")) {
                            add(vscode.SymbolKind.Array);
                        } else {
                            add(vscode.SymbolKind.Field);
                        }
                        break;
                    case COBOLTokenStyle.ConditionName:
                        if (includeVars) add(vscode.SymbolKind.TypeParameter);
                        break;
                    case COBOLTokenStyle.Union:
                        if (includeVars) add(vscode.SymbolKind.Struct);
                        break;
                    case COBOLTokenStyle.Constant:
                        if (includeVars) add(vscode.SymbolKind.Constant);
                        break;
                    case COBOLTokenStyle.Property: add(vscode.SymbolKind.Property); break;
                    case COBOLTokenStyle.Constructor: add(vscode.SymbolKind.Constructor); break;
                }
            } catch (e) {
                VSLogger.logException(`Failed ${e} on ${JSON.stringify(token)}`, e as Error);
            }
        }
        return symbols;
    }
}
