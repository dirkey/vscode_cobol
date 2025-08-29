import * as vscode from "vscode";
import { CancellationToken, FormattingOptions, TextDocument, TextEdit, Position } from "vscode";
import { VSCOBOLUtils, FoldAction } from "./vscobolutils";
import { ICOBOLSettings } from "./iconfiguration";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

export class COBOLCaseFormatter {

    private convertLine(settings: ICOBOLSettings, line: string, current: ICOBOLSourceScanner, foldConstantToUpper: boolean, langid: string): string {
        const defaultStyle = settings.intellisense_style;
        let text = line;
        const actions: FoldAction[] = [
            FoldAction.Keywords,
            FoldAction.ConstantsOrVariables,
            FoldAction.PerformTargets
        ];

        for (const action of actions) {
            text = VSCOBOLUtils.foldTokenLine(text, current, action, foldConstantToUpper, langid, settings, defaultStyle);
        }

        return text;
    }

    public provideOnTypeFormattingEdits(
        document: TextDocument,
        position: Position,
        ch: string,
        options: FormattingOptions,
        token: CancellationToken
    ): TextEdit[] | undefined {

        if (ch !== "\n") return;

        const settings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        if (!settings.format_on_return) return;

        const current = VSCOBOLSourceScanner.getCachedObject(document, settings);
        if (!current) return;

        const lineIndex = position.line - 1;
        const line = document.lineAt(lineIndex)?.text;
        if (!line) return;

        const newText = this.convertLine(settings, line, current, settings.format_constants_to_uppercase, document.languageId);
        if (newText === line) return [];

        const range = new vscode.Range(
            new vscode.Position(lineIndex, 0),
            new vscode.Position(lineIndex, line.length)
        );

        return [TextEdit.replace(range, newText)];
    }
}