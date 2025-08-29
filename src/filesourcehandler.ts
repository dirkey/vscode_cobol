import fs from "fs";
import path from "path";
import { pathToFileURL } from "url";

import { ISourceHandler, ICommentCallback, ISourceHandlerLite, commentRange } from "./isourcehandler";
import { ESourceFormat, IExternalFeatures } from "./externalfeatures";
import { getCOBOLKeywordDictionary } from "./keywords/cobolKeywords";
import { ExtensionDefaults } from "./extensionDefaults";
import { SimpleStringBuilder } from "./stringutils";
import { ICOBOLSettings } from "./iconfiguration";

export class FileSourceHandler implements ISourceHandler, ISourceHandlerLite {
    document: string;
    dumpNumbersInAreaA = false;
    dumpAreaBOnwards = false;
    lines: string[] = [];
    commentCount = 0;
    commentCallbacks: ICommentCallback[] = [];
    documentVersionId: bigint;
    isSourceInWorkspace = false;
    updatedSource = new Map<number, string>();
    shortFilename: string;
    languageId = ExtensionDefaults.defaultCOBOLLanguage;
    format: ESourceFormat = ESourceFormat.unknown;
    notedCommentRanges: commentRange[] = [];
    commentsIndex = new Map<number, string>();
    commentsIndexInline = new Map<number, boolean>();
    settings: ICOBOLSettings;

    constructor(settings: ICOBOLSettings, regEx: RegExp | undefined, document: string, features: IExternalFeatures) {
        this.settings = settings;
        this.document = document;
        this.shortFilename = this.findShortWorkspaceFilename(document, features, settings);

        const docstat = fs.statSync(document, { bigint: true });
        this.documentVersionId = docstat.mtimeMs;

        const startTime = features.performance_now();
        try {
            const linesRead = fs.readFileSync(document, "utf-8").split(/\r?\n/);
            this.lines = regEx
                ? linesRead.map(l => (regEx.test(l) ? l : ""))
                : linesRead;

            features.logTimedMessage(features.performance_now() - startTime, ` - Loading File ${document}`);
        } catch (e) {
            features.logException(`File load failed! (${document})`, e as Error);
        }
    }

    getDocumentVersionId(): bigint {
        return this.documentVersionId;
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
                l = l.trim();
            }

            if (l) {
                this.commentsIndex.set(lineNumber, l);
                this.commentsIndexInline.set(lineNumber, isInline);
            }
        }

        this.commentCallbacks.forEach(cb => cb.processComment(this.settings, this, line, this.getFilename(), lineNumber, startPos, format));
    }

    getUriAsString(): string {
        return pathToFileURL(this.getFilename()).href;
    }

    getLineCount(): number {
        return this.lines.length;
    }

    getCommentCount(): number {
        return this.commentCount;
    }

    private static readonly paraPrefixRegex1 = /^[0-9 ][0-9 ][0-9 ][0-9 ][0-9 ][0-9 ]/;

    getLine(lineNumber: number, raw = false): string | undefined {
        if (lineNumber >= this.lines.length) return undefined;

        let line = this.lines[lineNumber];
        if (raw) return line;

        // Process comments
        const startComment = line.indexOf("*>");
        if (startComment !== -1 && startComment !== 6) {
            this.commentCount++;
            this.sendCommentCallback(line, lineNumber, startComment, ESourceFormat.variable);
            line = line.substring(0, startComment);
        }

        // Drop fixed format comments
        if (line.length > 0 && (line[0] === "*" || (line.length >= 7 && (line[6] === "*" || line[6] === "/")))) {
            this.commentCount++;
            this.sendCommentCallback(line, lineNumber, line[0] === "*" ? 0 : 6, line[0] === "*" ? ESourceFormat.variable : ESourceFormat.fixed);
            return "";
        }

        // Terminal format debug lines
        if (this.format === ESourceFormat.terminal && (line.startsWith("\\D") || line.startsWith("|"))) {
            this.commentCount++;
            this.sendCommentCallback(line, lineNumber, 0, ESourceFormat.terminal);
            return "";
        }

        // Area A/B adjustments
        if (this.dumpNumbersInAreaA && line.match(FileSourceHandler.paraPrefixRegex1)) {
            line = "      " + line.substring(6);
        } else if (this.dumpAreaBOnwards && line.length >= 73) {
            line = line.substring(0, 72);
        }

        return line;
    }

    getLineTabExpanded(lineNumber: number): string | undefined {
        const unexpanded = this.getLine(lineNumber, true);
        if (!unexpanded) return undefined;

        const tabSize = 4;
        const buf = new SimpleStringBuilder();
        let col = 0;

        for (const c of unexpanded) {
            if (c === "\t") {
                do buf.Append(" "); while (++col % tabSize !== 0);
            } else {
                buf.Append(c);
            }
            col++;
        }
        return buf.ToString();
    }

    setDumpAreaA(flag: boolean) { this.dumpNumbersInAreaA = flag; }
    setDumpAreaBOnwards(flag: boolean) { this.dumpAreaBOnwards = flag; }
    isValidKeyword(keyword: string) { return getCOBOLKeywordDictionary(this.languageId).has(keyword); }
    getFilename() { return this.document; }
    addCommentCallback(cb: ICommentCallback) { this.commentCallbacks.push(cb); }
    resetCommentCount() { this.commentCount = 0; }
    getIsSourceInWorkSpace() { return this.isSourceInWorkspace; }
    getShortWorkspaceFilename() { return this.shortFilename; }
    getUpdatedLine(lineNumber: number) { return this.updatedSource.get(lineNumber) ?? this.getLine(lineNumber, false); }
    setUpdatedLine(lineNumber: number, line: string) { this.updatedSource.set(lineNumber, line); }

    private findShortWorkspaceFilename(ddir: string, features: IExternalFeatures, config: ICOBOLSettings): string {
        const ws = features.getWorkspaceFolders(config);
        if (!ws?.length) return "";

        const fullPath = path.normalize(ddir);
        let best = "";
        for (const folderPath of ws) {
            if (fullPath.startsWith(folderPath)) {
                const candidate = fullPath.substring(folderPath.length + 1);
                if (!best || candidate.length < best.length) best = candidate;
            }
        }
        return best;
    }

    getLanguageId() { return this.languageId; }
    setSourceFormat(format: ESourceFormat) { this.format = format; }
    getNotedComments() { return this.notedCommentRanges; }

    getCommentAtLine(lineNumber: number): string {
        if (this.commentsIndex.has(lineNumber)) return this.commentsIndex.get(lineNumber)!;

        let lines = "";
        if (lineNumber > 1 && this.commentsIndex.has(lineNumber - 1) && this.commentsIndexInline.get(lineNumber - 1) === false) {
            let maxLines = 5;
            let offset = 2;
            lines = this.commentsIndex.get(lineNumber - 1)! + "\n";
            while (maxLines-- > 0 && this.commentsIndex.has(lineNumber - offset) && this.commentsIndexInline.get(lineNumber - offset) === false) {
                lines = this.commentsIndex.get(lineNumber - offset)! + "\n" + lines;
                offset++;
            }
        }

        return lines;
    }

    getText(startLine: number, startColumn: number, endLine: number, endColumn: number): string {
        let result = this.getLine(startLine, true)?.substring(startColumn) ?? "";
        for (let ln = startLine + 1; ln <= endLine; ln++) {
            const line = this.getLine(ln, true) ?? "";
            result += "\n" + (ln === endLine ? line.substring(0, endColumn) : line);
        }
        return result;
    }
}
