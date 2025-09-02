import * as vscode from "vscode";
import { COBOLToken, COBOLVariable } from "./cobolsourcescanner";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { ICOBOLSettings } from "./iconfiguration";
import { VSLogger } from "./vslogger";
import { getCOBOLKeywordDictionary } from "./keywords/cobolKeywords";
import { VSCOBOLSourceScannerTools } from "./vssourcescannerutils";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

export class COBOLSourceDefinition implements vscode.DefinitionProvider {
    // 🔒 Regex are fully immutable and typed
    private static readonly sectionRegEx = /[$0-9a-zA-Z][a-zA-Z0-9-_]*/;
    private static readonly variableRegEx = /[$#0-9a-zA-Z_][a-zA-Z0-9-_]*/;
    private static readonly classRegEx = /[$0-9a-zA-Z][a-zA-Z0-9-_]*/;
    private static readonly methodRegEx = /[$0-9a-zA-Z][a-zA-Z0-9-_]*/;

    public provideDefinition(
        document: vscode.TextDocument,
        position: vscode.Position,
        _token: vscode.CancellationToken
    ): vscode.ProviderResult<vscode.Definition> {
        const locations: vscode.Location[] = [];
        const config: ICOBOLSettings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        const lineText: string = document.lineAt(position.line).text;

        const qcp: ICOBOLSourceScanner | undefined = VSCOBOLSourceScanner.getCachedObject(document, config);

        if (/.*(perform|thru|go\s*to|until|varying).*$/i.test(lineText) && qcp) {
            const loc = this.getSectionOrParaLocation(document, qcp, position);
            if (loc) return [loc];
        }

        if (/.*(new\s*|type).*$/i.test(lineText) && qcp) {
            const loc = this.getClassTarget(document, qcp, position);
            if (loc) return [loc];
        }

        if (/.*(invoke\s*|::).*$/i.test(lineText) && qcp) {
            const loc = this.getMethodTarget(document, qcp, position);
            if (loc) return [loc];
        }

        if (this.getVariableInCurrentDocument(locations, document, position, config)) {
            return locations;
        }

        if (qcp && VSCOBOLSourceScannerTools.isPositionInEXEC(qcp, position)) {
            const loc = this.getSQLCursor(document, qcp, position);
            if (loc) return [loc];
        }

        return locations;
    }

    private isValidToken(token: unknown): token is COBOLToken {
        return (
            typeof token === "object" &&
            token !== null &&
            "rangeStartLine" in (token as COBOLToken) &&
            "rangeStartColumn" in (token as COBOLToken) &&
            "filenameAsURI" in (token as COBOLToken)
        );
    }

    private createLocation(token: COBOLToken): vscode.Location {
        const start = new vscode.Position(token.rangeStartLine, token.rangeStartColumn);
        const end = new vscode.Position(token.rangeEndLine, token.rangeEndColumn);
        const range = new vscode.Range(start, end);
        const uri = vscode.Uri.parse(token.filenameAsURI);
        return new vscode.Location(uri, range);
    }

    private getWord(document: vscode.TextDocument, position: vscode.Position, regex: RegExp): string | undefined {
        const range = document.getWordRangeAtPosition(position, regex);
        return range ? document.getText(range) : undefined;
    }

    private getSectionOrParaLocation(
        document: vscode.TextDocument,
        sf: ICOBOLSourceScanner,
        position: vscode.Position
    ): vscode.Location | undefined {
        const word = this.getWord(document, position, COBOLSourceDefinition.sectionRegEx)?.toLowerCase();
        if (!word) return undefined;

        try {
            const token = sf.sections.get(word) ?? sf.paragraphs.get(word);
            return this.isValidToken(token) ? this.createLocation(token) : undefined;
        } catch (e) {
            VSLogger.logMessage((e as Error).message);
            return undefined;
        }
    }

    private getSQLCursor(
        document: vscode.TextDocument,
        sf: ICOBOLSourceScanner,
        position: vscode.Position
    ): vscode.Location | undefined {
        const word = this.getWord(document, position, COBOLSourceDefinition.sectionRegEx)?.toLowerCase();
        if (!word) return undefined;

        try {
            const decl = sf.execSQLDeclare.get(word);
            return decl && this.isValidToken(decl.token) ? this.createLocation(decl.token) : undefined;
        } catch (e) {
            VSLogger.logMessage((e as Error).message);
            return undefined;
        }
    }

    private getVariableInCurrentDocument(
        locations: vscode.Location[],
        document: vscode.TextDocument,
        position: vscode.Position,
        settings: ICOBOLSettings
    ): boolean {
        const word = this.getWord(document, position, COBOLSourceDefinition.variableRegEx);
        if (!word) return false;

        const tokenLower: string = word.toLowerCase();
        if (getCOBOLKeywordDictionary(document.languageId).has(tokenLower)) return false;

        const sf: ICOBOLSourceScanner | undefined = VSCOBOLSourceScanner.getCachedObject(document, settings);
        if (!sf) return false;

        const variables: COBOLVariable[] | undefined = sf.constantsOrVariables.get(tokenLower);
        if (!variables?.length) return false;

        for (const variable of variables) {
            const token = variable.token;
            if (token.tokenNameLower !== "filler" && this.isValidToken(token)) {
                locations.push(this.createLocation(token));
            }
        }

        return locations.length > 0;
    }

    private getGenericTarget(
        regex: RegExp,
        tokenMap: Map<string, COBOLToken>,
        document: vscode.TextDocument,
        position: vscode.Position
    ): vscode.Location | undefined {
        const word = this.getWord(document, position, regex)?.toLowerCase();
        const token = word ? tokenMap.get(word) : undefined;
        return this.isValidToken(token) ? this.createLocation(token) : undefined;
    }

    private getClassTarget(
        document: vscode.TextDocument,
        sf: ICOBOLSourceScanner,
        position: vscode.Position
    ): vscode.Location | undefined {
        return this.getGenericTarget(COBOLSourceDefinition.classRegEx, sf.classes, document, position);
    }

    private getMethodTarget(
        document: vscode.TextDocument,
        sf: ICOBOLSourceScanner,
        position: vscode.Position
    ): vscode.Location | undefined {
        return this.getGenericTarget(COBOLSourceDefinition.methodRegEx, sf.methods, document, position);
    }
}
