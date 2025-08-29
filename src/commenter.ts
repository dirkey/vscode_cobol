"use strict";

import { Position, Range, TextDocument, TextEditorEdit, window } from "vscode";
import { ESourceFormat } from "./externalfeatures";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { ICOBOLSettings } from "./iconfiguration";

export class commentUtils {
    static readonly var_free_insert_at_comment_column = true;

    public static processCommentLine(settings: ICOBOLSettings): void {
        const editor = window.activeTextEditor;
        if (!editor) return;

        const doc = editor.document;
        const format = VSCOBOLSourceScanner.getCachedObject(doc, settings)?.sourceFormat ?? ESourceFormat.variable;

        editor.edit(edit => {
            for (const sel of editor.selections) {
                for (let line = sel.start.line; line <= sel.end.line; line++) {
                    commentUtils.toggleLine(edit, doc, line, format);
                }
            }
        });
    }

    private static toggleLine(edit: TextEditorEdit, doc: TextDocument, lineNum: number, format: ESourceFormat) {
        const text = doc.lineAt(lineNum).text;

        if (format === ESourceFormat.fixed) {
            // Toggle '*' at column 6
            const replaceRange = new Range(new Position(lineNum, 6), new Position(lineNum, 7));
            edit.replace(replaceRange, text[6] === "*" || text[6] === "/" ? " " : "*");
            return;
        }

        const token = format === ESourceFormat.terminal ? "|" : "*> ";
        const existingIndex = text.indexOf(token);
        if (existingIndex >= 0) {
            const deleteRange = new Range(new Position(lineNum, existingIndex), new Position(lineNum, existingIndex + token.length));
            edit.delete(deleteRange);
            return;
        }

        const firstNonSpace = text.search(/\S|$/);
        let insertPos = 0;

        if (token === "|") {
            insertPos = Math.max(0, firstNonSpace);
        } else if (
            firstNonSpace === 6 ||
            (this.var_free_insert_at_comment_column && firstNonSpace > 6 && text.slice(6, 9) === "   ")
        ) {
            insertPos = 6;
        } else if (firstNonSpace === 5 && text[6] === "*") {
            insertPos = 7;
        } else {
            insertPos = Math.max(0, firstNonSpace - 1);
        }

        edit.insert(new Position(lineNum, insertPos), token);
    }
}