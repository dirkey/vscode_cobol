"use strict";

import * as vscode from "vscode";
import * as path from "path";
import { Range, TextDocument, Definition, Position, CancellationToken, Uri, Location, DefinitionProvider } from "vscode";
import { IVSCOBOLSettings, VSCOBOLConfiguration } from "./vsconfiguration";
import { ICOBOLSettings } from "./iconfiguration";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { VSCOBOLFileUtils } from "./vsfileutils";
import { IExternalFeatures } from "./externalfeatures";
import { VSLogger } from "./vslogger";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

export class COBOLCopyBookProvider implements DefinitionProvider {

    private features: IExternalFeatures;

    constructor(features: IExternalFeatures) {
        this.features = features;
    }

    public provideDefinition(
        document: TextDocument,
        position: Position,
        token: CancellationToken
    ): Promise<Definition> {
        return this.resolveDefinitions(document, position, token);
    }

    private async getURIForCopybook(copybook: string, settings: ICOBOLSettings): Promise<Uri | undefined> {
        try {
            const uriString = await VSCOBOLFileUtils.findCopyBookViaURL(copybook, settings, this.features);
            return uriString.length ? Uri.parse(uriString) : undefined;
        } catch (e) {
            VSLogger.logException("getURIForCopybook", e as Error);
            return undefined;
        }
    }

    private async getURIForCopybookInDirectory(copybook: string, inDirectory: string, settings: ICOBOLSettings): Promise<Uri | undefined> {
        try {
            const uriString = await VSCOBOLFileUtils.findCopyBookInDirectoryViaURL(copybook, inDirectory, settings, this.features);
            return uriString.length ? Uri.parse(uriString) : undefined;
        } catch (e) {
            VSLogger.logException("getURIForCopybookInDirectory", e as Error);
            return undefined;
        }
    }

    private async resolveDefinitions(document: TextDocument, pos: Position, ct: CancellationToken): Promise<Definition> {
        const config = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        const scanner: ICOBOLSourceScanner | undefined = VSCOBOLSourceScanner.getCachedObject(document, config);

        if (!scanner) return this.resolveDefinitionsFallback(true, document, pos, ct);

        for (const [, copyBooks] of scanner.copyBooksUsed) {
            for (const copyBook of copyBooks) {
                const st = copyBook.statementInformation;
                if (!st) continue;

                const stRange = new Range(
                    new Position(st.startLineNumber, st.startCol),
                    new Position(st.endLineNumber, st.endCol)
                );
                if (!stRange.contains(pos)) continue;

                if (document.uri.scheme === "file" && st.fileName.length) {
                    return new Location(Uri.file(st.fileName), new Range(new Position(0, 0), new Position(0, 0)));
                }

                if (st.isIn) {
                    const uri = await this.getURIForCopybookInDirectory(st.copyBook, st.literal2, config);
                    if (uri) return new Location(uri, new Range(new Position(0, 0), new Position(0, 0)));
                } else {
                    const uri = await this.getURIForCopybook(st.copyBook, config);
                    if (uri) return new Location(uri, new Range(new Position(0, 0), new Position(0, 0)));
                }
            }
        }

        return this.resolveDefinitionsFallback(false, document, pos, ct);
    }

    private async resolveDefinitionsFallback(everything: boolean, doc: TextDocument, pos: Position, ct: CancellationToken): Promise<Definition> {
        const config = VSCOBOLConfiguration.get_resource_settings(doc, VSExternalFeatures);
        const lineText = doc.lineAt(pos).text;
        let filename = this.extractCopyBookFilename(everything, config, lineText);

        // fallback to editor selection if no filename
        const editor = vscode.window.activeTextEditor;
        if (!filename && editor) {
            const selText = editor.document.getText(editor.selection).trim();
            if (selText.length) filename = selText;
        }
        if (!filename) return [];

        // determine inDirectory from "IN" or "OF"
        const textLower = lineText.toLowerCase();
        let inDirectory = "";
        const copyIndex = textLower.indexOf("copy");
        if (copyIndex !== -1) {
            let posIn = textLower.indexOf(" in ");
            if (posIn === -1) posIn = textLower.indexOf(" of ");
            if (posIn !== -1) {
                inDirectory = lineText.substr(posIn + 4).trim();
                inDirectory = this.cleanQuotesAndDot(inDirectory);
            }
        }

        const fullPath = COBOLCopyBookProvider.expandLogicalCopyBookOrEmpty(filename.trim(), inDirectory, config, doc.fileName, this.features);
        if (fullPath.length) return new Location(Uri.file(fullPath), new Range(new Position(0, 0), new Position(0, 0)));

        // handle scan-comment-copybook-token
        const commentPos = textLower.indexOf(config.scan_comment_copybook_token.toLowerCase());
        if (commentPos !== -1) {
            const fileRegex = /[-9a-zA-Z\/\ \.:_]+/g;
            const wordRange = doc.getWordRangeAtPosition(pos, fileRegex);
            const wordText = wordRange ? doc.getText(wordRange) : "";
            for (const possibleFilename of wordText.split(" ")) {
                const possiblePath = COBOLCopyBookProvider.expandLogicalCopyBookOrEmpty(possibleFilename.trim(), inDirectory, config, doc.fileName, this.features);
                if (possiblePath.length && this.features.isFile(possiblePath) && !this.features.isDirectory(possiblePath)) {
                    return new Location(Uri.file(possiblePath), new Range(new Position(0, 0), new Position(0, 0)));
                }
            }
        }

        return [];
    }

    private cleanQuotesAndDot(str: string): string {
        let result = str.trim();
        if (result.endsWith(".")) result = result.slice(0, -1);
        if ((result.startsWith("\"") && result.endsWith("\"")) || (result.startsWith("'") && result.endsWith("'"))) {
            result = result.slice(1, -1);
        }
        return result.trim();
    }

    private extractCopyBookFilename(everything: boolean, config: ICOBOLSettings, str: string): string | undefined {
        const strLower = str.toLowerCase();
        if (!everything) return undefined;

        const copyPos = strLower.indexOf("copy");
        if (copyPos !== -1) {
            let result = str.substr(copyPos + 4).trimStart();
            const spaceIndex = result.indexOf(" ");
            if (spaceIndex !== -1) result = result.substr(0, spaceIndex).trim();
            if (result.endsWith(".")) result = result.slice(0, -1).trim();
            result = this.cleanQuotesAndDot(result);
            return result;
        }

        // handle exec sql include
        if (strLower.includes("exec") && strLower.includes("sql")) {
            const includePos = strLower.indexOf("include");
            if (includePos !== -1) {
                let filename = str.substr(includePos + 7).trimStart();
                const endExec = filename.toLowerCase().indexOf("end-exec");
                if (endExec !== -1) filename = filename.substr(0, endExec).trim();
                return filename;
            }
        }

        return undefined;
    }

    public static expandLogicalCopyBookOrEmpty(
        filename: string,
        inDirectory: string,
        config: IVSCOBOLSettings,
        sourceFilename: string,
        features: IExternalFeatures
    ): string {
        const fileDirname = path.dirname(sourceFilename);

        for (const perCopyDir of config.perfile_copybookdirs) {
            const perFileDir = perCopyDir.replace("${fileDirname}", fileDirname);
            let firstFile = path.join(perFileDir, filename);
            if (features.isFile(firstFile)) return firstFile;

            if (!filename.includes(".")) {
                for (const ext of config.copybookexts) {
                    const fileWithExt = path.join(perFileDir, filename + "." + ext);
                    if (features.isFile(fileWithExt)) return fileWithExt;
                }
            }
        }

        if (!inDirectory.length) {
            const fullPath = VSCOBOLFileUtils.findCopyBook(filename, config, features);
            return fullPath.length ? path.normalize(fullPath) : "";
        }

        const fullPath = VSCOBOLFileUtils.findCopyBookInDirectory(filename, inDirectory, config, features);
        return fullPath.length ? path.normalize(fullPath) : "";
    }
}