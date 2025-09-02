import * as vscode from "vscode";
import { VSLogger } from "./vslogger";
import { VSWorkspaceFolders } from "./vscobolfolders";
import { ICOBOLSettings } from "./iconfiguration";
import path from "path";

export class VSDiagCommands {

    private static dumpSymbol(sb: string[], depth: string, symbol: vscode.DocumentSymbol) {
        const symbolKind = vscode.SymbolKind[symbol.kind];
        const detail = symbol.detail?.length ? ` "${symbol.detail}"` : "";
        const tag = symbol.tags?.length ? ` (Tag=${symbol.tags.join(',')})` : "";

        const rrInfo = symbol.range.isSingleLine
            ? `${symbol.range.start.line}:${symbol.range.start.character}-${symbol.range.end.character}`
            : `${symbol.range.start.line}:${symbol.range.start.character} -> ${symbol.range.end.line}:${symbol.range.end.character}`;

        const srInfo = symbol.selectionRange.isSingleLine
            ? `${symbol.selectionRange.start.line}:${symbol.selectionRange.start.character}-${symbol.selectionRange.end.character}`
            : `${symbol.selectionRange.start.line}:${symbol.selectionRange.start.character} -> ${symbol.selectionRange.end.line}:${symbol.selectionRange.end.character}`;

        const rangeInfo = !symbol.range.isSingleLine && rrInfo !== srInfo ? `${rrInfo} / ${srInfo}` : rrInfo;

        sb.push(`${depth}  ${symbolKind} : "${symbol.name}"${detail}${tag} @ ${rangeInfo}`);

        const activeEditor = vscode.window.activeTextEditor;
        const text = activeEditor?.document.getText(symbol.range) ?? "";
        sb.push(`\`\`\`${activeEditor?.document.languageId}\n${text}\n\`\`\`\n\n`);

        for (const child of symbol.children) {
            VSDiagCommands.dumpSymbol(sb, depth + "#", child);
        }
    }

    public static async DumpAllSymbols(config: ICOBOLSettings) {
        const activeEditor = vscode.window.activeTextEditor;
        if (!activeEditor) return;

        const sb: string[] = ["Document Symbols found:"];
        const documentUri = activeEditor.document.uri;

        const symbolsArray = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>(
            "vscode.executeDocumentSymbolProvider",
            documentUri
        );

        if (!symbolsArray) return;

        for (const symbol of symbolsArray) {
            try {
                VSDiagCommands.dumpSymbol(sb, "#", symbol);
            } catch (e) {
                VSLogger.logException(`Symbol: ${symbol.name}`, e as Error);
            }
        }

        const fileName = "dump.md";
        const wsFolders = VSWorkspaceFolders.get(config);
        const filePath = wsFolders?.[0].uri.fsPath
            ? path.join(wsFolders[0].uri.fsPath, fileName)
            : path.join(process.cwd(), fileName);

        const untitledUri = vscode.Uri.file(filePath).with({ scheme: "untitled" });
        const document = await vscode.workspace.openTextDocument(untitledUri);
        const editor = await vscode.window.showTextDocument(document);

        await vscode.languages.setTextDocumentLanguage(document, "markdown");

        await editor.edit(editBuilder => {
            editBuilder.insert(new vscode.Position(0, 0), sb.join("\n"));
        });

        VSLogger.logMessage(sb.join("\n"));
    }
}