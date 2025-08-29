import { ESourceFormat } from "./externalfeatures";
import { fileformatStrategy, ICOBOLSettings } from "./iconfiguration";
import { ISourceHandlerLite } from "./isourcehandler";
import globToRegExp = require("glob-to-regexp");
import { getCOBOLKeywordDictionary } from "./keywords/cobolKeywords";

const INLINE_SOURCE_FORMAT: string[] = ["sourceformat", ">>source format"];

class LineAnalyzer {
    public validFixedLine = false;
    public hasKeywordAtColumn0 = false;
    public formatDetected?: ESourceFormat;

    constructor(
        private line: string,
        private keywords: Set<string>,
        private isTerminal: boolean
    ) {
        this.analyze();
    }

    private analyze() {
        const trimmedLine = this.line.trimEnd();
        if (!trimmedLine) return;

        const lowerLine = trimmedLine.toLowerCase();

        // Fixed line detection
        this.validFixedLine = lowerLine.length >= 7 && ["*", "D", "/", " ", "-"].includes(lowerLine[6]);

        // Terminal format detection
        if (this.isTerminal && (lowerLine.startsWith("*") || lowerLine.startsWith("|") || lowerLine.startsWith("\\D"))) {
            this.formatDetected = ESourceFormat.terminal;
            return;
        }

        // Inline source format detection
        for (const token of INLINE_SOURCE_FORMAT) {
            const idx = lowerLine.indexOf(token);
            if (idx !== -1) {
                const remainder = lowerLine.substring(idx + token.length + 1);
                if (remainder.includes("fixed")) this.formatDetected = ESourceFormat.fixed;
                else if (remainder.includes("variable")) this.formatDetected = ESourceFormat.variable;
                else if (remainder.includes("free")) this.formatDetected = ESourceFormat.free;
                return;
            }
        }

        // Free/variable heuristics
        if (lowerLine.includes("*>")) {
            this.formatDetected = ESourceFormat.variable;
            return;
        }

        if (!this.validFixedLine && lowerLine.length > 80) {
            this.formatDetected = ESourceFormat.variable;
            return;
        }

        // Keyword detection
        let firstWord = lowerLine.split(" ")[0].replace(/\.$/, "");
        if (firstWord && this.keywords.has(firstWord)) this.hasKeywordAtColumn0 = true;
    }
}

export class SourceFormat {
    private static getFileFormat(doc: ISourceHandlerLite, config: ICOBOLSettings): ESourceFormat | undefined {
        for (const filter of config.editor_margin_files) {
            try {
                const re = globToRegExp(filter.pattern, { flags: "i" });
                if (re.test(doc.getFilename())) return ESourceFormat[filter.sourceformat];
            } catch {}
        }
        return undefined;
    }

    public static get(doc: ISourceHandlerLite, config: ICOBOLSettings): ESourceFormat {
        // Strategy overrides
        switch (config.fileformat_strategy) {
            case fileformatStrategy.AlwaysFixed: return ESourceFormat.fixed;
            case fileformatStrategy.AlwaysVariable: return ESourceFormat.variable;
            case fileformatStrategy.AlwaysFree: return ESourceFormat.free;
            case fileformatStrategy.AlwaysTerminal: return ESourceFormat.terminal;
        }

        // File-based overrides
        if (config.check_file_format_before_file_scan) {
            const fileFormat = this.getFileFormat(doc, config);
            if (fileFormat) return fileFormat;
        }

        const langId = doc.getLanguageId().toLowerCase();
        const isTerminal = langId === "acucobol";
        const keywords = new Set(getCOBOLKeywordDictionary(langId).keys());

        let defFormat: ESourceFormat = ESourceFormat.unknown;
        let validFixedLines = 0;
        let invalidFixedLines = 0;
        let skippedLines = 0;
        let linesGT80 = 0;
        let keywordAtColumn0 = 0;

        const maxLines = Math.min(doc.getLineCount(), config.pre_scan_line_limit);

        for (let i = 0; i < maxLines; i++) {
            const lineText = doc.getLineTabExpanded(i);
            if (!lineText || !lineText.trim()) {
                skippedLines++;
                continue;
            }

            const analyzer = new LineAnalyzer(lineText, keywords, isTerminal);

            if (analyzer.formatDetected) return analyzer.formatDetected;

            analyzer.validFixedLine ? validFixedLines++ : invalidFixedLines++;
            if (analyzer.hasKeywordAtColumn0) keywordAtColumn0++;

            if (!analyzer.validFixedLine && lineText.length > 80) linesGT80++;
        }

        // Heuristics
        if (keywordAtColumn0 >= 2 && invalidFixedLines >= 2) {
            defFormat = isTerminal ? ESourceFormat.terminal : ESourceFormat.free;
        }

        if (defFormat === ESourceFormat.unknown) {
            defFormat = isTerminal ? ESourceFormat.terminal : ESourceFormat.variable;
        }

        if (invalidFixedLines === 0 && linesGT80 === 0 && (validFixedLines + skippedLines === maxLines)) {
            return ESourceFormat.fixed;
        }

        // Late file-based override
        if (!config.check_file_format_before_file_scan) {
            const fileFormat = this.getFileFormat(doc, config);
            return fileFormat ?? defFormat;
        }

        return defFormat;
    }
}