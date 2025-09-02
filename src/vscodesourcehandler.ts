import * as vscode from "vscode";
import { workspace } from "vscode";
import { ESourceFormat, IExternalFeatures } from "./externalfeatures";
import { ISourceHandler, ICommentCallback, ISourceHandlerLite, CommentRange } from "./isourcehandler";
import { getCOBOLKeywordDictionary } from "./keywords/cobolKeywords";
import { SimpleStringBuilder } from "./stringutils";
import { colourCommentHandler } from "./vscolourcomments";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { VSCOBOLFileUtils } from "./vsfileutils";
import { ICOBOLSettings } from "./iconfiguration";

export class VSCodeSourceHandlerLite implements ISourceHandlerLite {
    private document: vscode.TextDocument;
    private tabSize: number;
    private lineCount: number;
    private languageId: string;
    private notedCommentRanges: CommentRange[];

    constructor(document: vscode.TextDocument) {
        this.document = document;
        this.lineCount = document.lineCount;
        this.languageId = document.languageId;
        this.notedCommentRanges = [];

        const editorConfig = workspace.getConfiguration("editor", document.uri);
        this.tabSize = editorConfig.get<number>("tabSize", 4);
    }

    getLineCount(): number {
        return this.lineCount;
    }

    getLanguageId(): string {
        return this.languageId;
    }

    getFilename(): string {
        return this.document.fileName;
    }

    getRawLine(lineNumber: number): string | undefined {
        if (lineNumber >= this.lineCount) return undefined;
        return this.document.lineAt(lineNumber).text;
    }

    getLineTabExpanded(lineNumber: number): string | undefined {
        const line = this.getRawLine(lineNumber);
        if (!line || line.indexOf("\t") === -1) return line;

        let col = 0;
        const buf = new SimpleStringBuilder();
        for (const c of line) {
            if (c === "\t") {
                do buf.Append(" "); while (++col % this.tabSize !== 0);
            } else {
                buf.Append(c);
                col++;
            }
        }
        return buf.ToString();
    }

    getNotedComments(): CommentRange[] {
        return this.notedCommentRanges;
    }

    getCommentAtLine(lineNumber: number): string {
        return ""; // Lite handler has no comments
    }
}

export class VSCodeSourceHandler implements ISourceHandler, ISourceHandlerLite {
    private document: vscode.TextDocument;
    private commentCount = 0;
    private dumpNumbersInAreaA = false;
    private dumpAreaBOnwards = false;
    private commentCallbacks: ICommentCallback[] = [];
    private lineCount: number;
    private documentVersionId: BigInt;
    private isSourceInWorkSpace: boolean;
    private shortWorkspaceFilename: string;
    private updatedSource = new Map<number, string>();
    private languageId: string;
    private format: ESourceFormat = ESourceFormat.unknown;
    private externalFeatures: IExternalFeatures = VSExternalFeatures;
    private notedCommentRanges: CommentRange[] = [];
    private commentsIndex = new Map<number, string>();
    private commentsIndexInline = new Map<number, boolean>();
    private tabSize: number;
    private config: ICOBOLSettings;

    private static readonly paraPrefixRegex1 = /^[0-9 ][0-9 ][0-9 ][0-9 ][0-9 ][0-9 ]/g;

    constructor(document: vscode.TextDocument) {
        this.document = document;
        this.lineCount = document.lineCount;
        this.documentVersionId = BigInt(document.version);
        this.languageId = document.languageId;

        this.config = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        this.shortWorkspaceFilename = VSCOBOLFileUtils.getShortWorkspaceFilename(document.uri.scheme, document.fileName, this.config) ?? "";
        this.isSourceInWorkSpace = this.shortWorkspaceFilename.length > 0;

        const editorConfig = workspace.getConfiguration("editor", document.uri);
        this.tabSize = editorConfig.get<number>("tabSize", 4);

        if (!vscode.workspace.isTrusted && !this.isSourceInWorkSpace) this.clear();
        if (this.isFileExcluded(this.config)) this.clear();

        this.addCommentCallback(colourCommentHandler);
    }

    private clear(): void {
        this.commentCallbacks = [];
        this.document = undefined as unknown as vscode.TextDocument;
        this.lineCount = 0;
    }

    private isFileExcluded(config: ICOBOLSettings): boolean {
        if (!this.document) return true;
        if (this.document.lineCount > config.scan_line_limit) {
            this.externalFeatures.logMessage(`Aborted scanning ${this.shortWorkspaceFilename} after line limit`);
            return true;
        }
        for (const fileEx of config.files_exclude) {
            if (vscode.languages.match({ pattern: fileEx }, this.document)) {
                this.externalFeatures.logMessage(`Aborted scanning ${this.shortWorkspaceFilename} (files_exclude)`);
                return true;
            }
        }
        return false;
    }

    getDocumentVersionId(): BigInt { return this.documentVersionId; }
    getUriAsString(): string { return this.document?.uri.toString() ?? ""; }
    getLineCount(): number { return this.lineCount; }
    getCommentCount(): number { return this.commentCount; }
    getLanguageId(): string { return this.languageId; }
    getFilename(): string { return this.document?.fileName ?? ""; }
    getShortWorkspaceFilename(): string { return this.shortWorkspaceFilename; }
    getIsSourceInWorkSpace(): boolean { return this.isSourceInWorkSpace; }
    getNotedComments(): CommentRange[] { return this.notedCommentRanges; }

    setDumpAreaA(flag: boolean): void { this.dumpNumbersInAreaA = flag; }
    setDumpAreaBOnwards(flag: boolean): void { this.dumpAreaBOnwards = flag; }
    setSourceFormat(format: ESourceFormat): void { this.format = format; }

    addCommentCallback(cb: ICommentCallback): void { this.commentCallbacks.push(cb); }
    resetCommentCount(): void { this.commentCount = 0; }

    setUpdatedLine(lineNumber: number, line: string): void { this.updatedSource.set(lineNumber, line); }
    getUpdatedLine(lineNumber: number): string | undefined {
        return this.updatedSource.get(lineNumber) ?? this.getLine(lineNumber, false);
    }

    private sendCommentCallback(line: string, lineNumber: number, startPos: number, format: ESourceFormat) {
        if (!this.commentsIndex.has(lineNumber)) {
            let isInline = false;
            let l = line.substring(startPos).trimStart();

            if (l.startsWith("*>") && format !== ESourceFormat.fixed) {
                isInline = true;
                l = l.substring(2).trim();
            } else if (line.length >= 7 && (line[6] === "*" || line[6] === "/")) {
                l = line[6] === "*" && line[7] === ">" ? line.substring(8, 72) : line.substring(7, 72);
            }

            if (l.length > 0) {
                this.commentsIndex.set(lineNumber, l);
                this.commentsIndexInline.set(lineNumber, isInline);
            }
        }

        for (const cb of this.commentCallbacks) {
            cb.processComment(this.config, this, line, this.getFilename(), lineNumber, startPos, format);
        }
    }

    getLine(lineNumber: number, raw: boolean): string | undefined {
        if (!this.document || lineNumber >= this.lineCount) return undefined;
        let line = this.document.lineAt(lineNumber).text;
        if (raw) return line;

        const startComment = line.indexOf("*>");
        if (startComment !== -1 && startComment !== 6) {
            this.commentCount++;
            this.sendCommentCallback(line, lineNumber, startComment, ESourceFormat.variable);
            line = line.substring(0, startComment);
        }

        if ((line.length > 1 && line[0] === "*") || (line.length >= 7 && (line[6] === "*" || line[6] === "/"))) {
            this.commentCount++;
            this.sendCommentCallback(line, lineNumber, line[0] === "*" ? 0 : 6, line[0] === "*" ? ESourceFormat.variable : ESourceFormat.fixed);
            return "";
        }

        if (this.format === ESourceFormat.terminal && (line.startsWith("\\D") || line.startsWith("|"))) {
            this.commentCount++;
            this.sendCommentCallback(line, lineNumber, 0, ESourceFormat.terminal);
            return "";
        }

        if (this.dumpNumbersInAreaA && (line.match(VSCodeSourceHandler.paraPrefixRegex1) || (line.length > 7 && line[6] === " " && !this.isValidKeyword(line.substring(0, 6).trim())))) {
            line = "      " + line.substring(6);
        }

        if (this.dumpAreaBOnwards && line.length >= 73) line = line.substring(0, 72);
        return line;
    }

    getLineTabExpanded(lineNumber: number): string | undefined {
        const line = this.getLine(lineNumber, true);
        if (!line || line.indexOf("\t") === -1) return line;

        let col = 0;
        const buf = new SimpleStringBuilder();
        for (const c of line) {
            if (c === "\t") do buf.Append(" "); while (++col % this.tabSize !== 0);
            else { buf.Append(c); col++; }
        }
        return buf.ToString();
    }

    isValidKeyword(keyword: string): boolean {
        return getCOBOLKeywordDictionary(this.languageId).has(keyword);
    }

    getCommentAtLine(lineNumber: number): string {
        if (this.commentsIndex.has(lineNumber)) return "" + this.commentsIndex.get(lineNumber);

        if (lineNumber > 1 && this.commentsIndex.has(lineNumber - 1) && !this.commentsIndexInline.get(lineNumber - 1)) {
            let maxLines = 5, offset = 2;
            let lines = "" + this.commentsIndex.get(lineNumber - 1) + "\n";

            while (maxLines-- > 0 && this.commentsIndex.has(lineNumber - offset) && !this.commentsIndexInline.get(lineNumber - offset)) {
                lines = "" + this.commentsIndex.get(lineNumber - offset) + "\n" + lines;
                offset++;
            }
            return lines;
        }

        return "";
    }

    getText(startLine: number, startColumn: number, endLine: number, endColumn: number): string {
        try {
            if (!this.document) return "";
            return this.document.getText(new vscode.Range(startLine, startColumn, endLine, endColumn));
        } catch {
            return "";
        }
    }
}
