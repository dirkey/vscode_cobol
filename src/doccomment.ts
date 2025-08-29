"use strict";

import { CancellationToken, FormattingOptions, TextDocument, TextEdit, Position } from "vscode";

export class COBOLDocumentationCommentHandler {

    /**
     * Automatically continues Micro Focus COBOL documentation comments (*>>)
     * when the user presses Enter.
     */
    public provideOnTypeFormattingEdits(
        document: TextDocument,
        position: Position,
        ch: string,
        options: FormattingOptions,
        token: CancellationToken
    ): TextEdit[] | undefined {
        // Only handle Enter key
        if (ch !== "\n") return;

        const previousLine = document.lineAt(position.line - 1).text.trim();

        // Only continue if previous line is a Micro Focus doc comment
        if (previousLine.startsWith("*>>")) {
            return [TextEdit.insert(position, "*>> ")];
        }

        return [];
    }
}