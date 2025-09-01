import vscode, { ProviderResult } from "vscode";
import { ICOBOLSettings, hoverApi } from "./iconfiguration";
import { CallTarget, KnownAPIs } from "./keywords/cobolCallTargets";
import { VSCOBOLUtils } from "./vscobolutils";
import { COBOLToken, COBOLVariable, SQLDeclare } from "./cobolsourcescanner";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

const nhexRegEx = /[nN][xX]["'][0-9A-Fa-f]*["']/;
const hexRegEx = /[xX]["'][0-9A-Fa-f]*["']/;
const wordRegEx = /[#0-9a-zA-Z][a-zA-Z0-9-_$]*/;
const variableRegEx = /[$#0-9a-zA-Z_][a-zA-Z0-9-_]*/;

export class VSHoverProvider {

    private static wrapCommentAndCode(comment: string, code: string, language = "COBOL"): string {
        const cleanCode = code.endsWith("\n") ? code.slice(0, -1) : code;
        const hasComment = comment.trim().length > 0;

        const codeBlock = `\`\`\`${language}
${cleanCode}
\`\`\`
`;

        if (!hasComment) return codeBlock;

        return `\`\`\`text
${comment}
\`\`\`

${codeBlock}
`;
    }

    private static createHoverFromToken(token: COBOLToken, positionLine: number, language = "COBOL"): vscode.Hover | undefined {
        if (!token || token.startLine === positionLine) return undefined;

        const line = token.sourceHandler.getLine(token.startLine, false);
        if (!line) return undefined;

        const hoverMessage = VSHoverProvider.wrapCommentAndCode(
            token.sourceHandler.getCommentAtLine(token.startLine),
            line.trimEnd(),
            language
        );

        return hoverMessage.length > 0 ? new vscode.Hover(new vscode.MarkdownString(hoverMessage)) : undefined;
    }

    public static provideHover(
        settings: ICOBOLSettings,
        document: vscode.TextDocument,
        position: vscode.Position
    ): ProviderResult<vscode.Hover> {

        // --- Known APIs
        if (settings.hover_show_known_api !== hoverApi.Off) {
            const txt = document.getText(document.getWordRangeAtPosition(position, wordRegEx));
            const target: CallTarget | undefined = KnownAPIs.getCallTarget(document.languageId, txt);
            if (target) {
                const example = settings.hover_show_known_api === hoverApi.Long
                    ? `\n\n---\n\n~~~\n${target.example.join("\r\n")}\n~~~\n`
                    : "";
                return new vscode.Hover(`**${target.api}** - ${target.description}\n\n[\u2192 ${target.apiGroup}](${target.url})${example}`);
            }
        }

        // --- Encoded literals
        if (settings.hover_show_encoded_literals) {
            const nxtxt = document.getText(document.getWordRangeAtPosition(position, nhexRegEx));
            if (/^nx['"]/.test(nxtxt.toLowerCase())) {
                const ascii = VSCOBOLUtils.nxhex2a(nxtxt);
                if (ascii) return new vscode.Hover(`UTF16=${ascii}`);
            }

            const txt = document.getText(document.getWordRangeAtPosition(position, hexRegEx));
            if (/^x['"]/.test(txt.toLowerCase())) {
                const ascii = VSCOBOLUtils.hex2a(txt);
                if (ascii) return new vscode.Hover(`ASCII=${ascii}`);
            }
        }

        // --- Word lookup
        const wordRange = document.getWordRangeAtPosition(position, variableRegEx);
        const word = document.getText(wordRange);
        if (!word) return undefined;

        const sf: ICOBOLSourceScanner | undefined = VSCOBOLSourceScanner.getCachedObject(document, settings);
        if (!sf) return undefined;

        const inProcedureDivision =
            sf.sourceReferences.state.procedureDivision?.startLine !== undefined &&
            position.line > sf.sourceReferences.state.procedureDivision.startLine;

        // --- Variable definition hover
        if (settings.hover_show_variable_definition &&
            sf.sourceReferences.state.procedureDivision?.startLine !== undefined &&
            position.line >= sf.sourceReferences.state.procedureDivision.startLine) {

            const variables: COBOLVariable[] | undefined = sf.constantsOrVariables.get(word.toLowerCase());
            if (variables?.length) {
                const hoverMessage = variables.map(v => {
                    const token = v.token;
                    if (!token) return "";
                    const line = token.sourceHandler.getLine(token.startLine, false);
                    if (!line) return "";
                    const text = variables.length === 1 ? line.trim() : line.trimEnd();
                    return VSHoverProvider.wrapCommentAndCode(token.sourceHandler.getCommentAtLine(token.startLine), text);
                }).filter(Boolean).join("\n\n----\n\n");

                if (hoverMessage) return new vscode.Hover(new vscode.MarkdownString(hoverMessage));
            }
        }

        // --- Sections
        if (inProcedureDivision) {
            const lowerWord = word.toLowerCase();
            const selection = sf.sections.get(lowerWord);
            if (selection) {
                const sectionHover = VSHoverProvider.createHoverFromToken(selection, position.line);
                if (sectionHover) return sectionHover;
            }
            const paragraph = sf.paragraphs.get(lowerWord);
            if (paragraph) {
                const paragraphHover = VSHoverProvider.createHoverFromToken(paragraph, position.line);
                if (paragraphHover) return paragraphHover;
            }
        }

        // --- SQL DECLARE
        if (settings.enable_exec_sql_cursors) {
            const sqlToken: SQLDeclare | undefined = sf.execSQLDeclare.get(word.toLowerCase());
            if (sqlToken?.token && sqlToken.token.startLine !== position.line) {
                const t = sqlToken.token;
                const sc = t.rangeStartLine === t.rangeEndLine ? t.rangeStartColumn : 0;
                const lines = t.sourceHandler.getText(t.rangeStartLine, sc, t.rangeEndLine, t.rangeEndColumn);
                if (lines) {
                    const hoverMessage = VSHoverProvider.wrapCommentAndCode(
                        t.sourceHandler.getCommentAtLine(t.startLine),
                        lines.trimEnd()
                    );
                    return hoverMessage ? new vscode.Hover(new vscode.MarkdownString(hoverMessage)) : undefined;
                }
            }
        }

        return undefined;
    }
}