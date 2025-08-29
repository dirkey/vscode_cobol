"use strict";

import { Selection, TextEditorRevealType, window } from "vscode";

export class COBOLProgramCommands {

    private static readonly anyNextPatterns = [
        /.*\s*division/i,
        /entry\s*"/i,
        /.*\s*section\./i,
        /eject/i,
        /program-id\./i,
        /class-id[.|\s]/i,
        /method-id[.|\s]/i
    ];

    public static moveToProcedureDivision(): void {
        this.moveToLine(this.findMatch(/procedure\s*division/i), "PROCEDURE DIVISION");
    }

    public static moveToDataDivision(): void {
        let line = this.findMatch(/data\s*division/i) || this.findMatch(/working-storage\s*section/i);
        this.moveToLine(line, "DATA DIVISION or WORKING-STORAGE SECTION");
    }

    public static moveToWorkingStorage(): void {
        this.moveToLine(this.findMatch(/working-storage\s*section/i), "WORKING-STORAGE SECTION");
    }

    public static moveForward(): void {
        this.moveAny(1);
    }

    public static moveBackward(): void {
        this.moveAny(-1);
    }

    private static moveToLine(line: number, errorMsg: string): void {
        if (line > 0) {
            this.goToLine(line);
        } else {
            window.setStatusBarMessage(`ERROR: '${errorMsg}' not found.`, 4000);
        }
    }

    private static moveAny(direction: number): void {
        const line = this.findAnyNext(direction);
        if (line > 0) this.goToLine(line);
    }

    private static findMatch(regex: RegExp): number {
        const editor = window.activeTextEditor;
        if (!editor) return 0;

        for (let line = 0; line < editor.document.lineCount; line++) {
            if (editor.document.lineAt(line).text.match(regex)) {
                return line;
            }
        }
        return 0;
    }

    private static findAnyMatch(patterns: RegExp[], direction: number): number {
        const editor = window.activeTextEditor;
        if (!editor) return 0;

        const start = editor.selection.active.line + direction;
        const end = direction === 1 ? editor.document.lineCount : -1;

        for (let line = start; line !== end; line += direction) {
            const text = editor.document.lineAt(line).text;
            if (patterns.some(p => text.match(p))) return line;
        }
        return 0;
    }

    private static findAnyNext(direction: number): number {
        return this.findAnyMatch(this.anyNextPatterns, direction);
    }

    private static goToLine(line: number): void {
        const editor = window.activeTextEditor;
        if (!editor) return;

        const revealType = line === editor.selection.active.line
            ? TextEditorRevealType.InCenterIfOutsideViewport
            : TextEditorRevealType.InCenter;

        const selection = new Selection(line, 0, line, 0);
        editor.selection = selection;
        editor.revealRange(selection, revealType);
    }
}