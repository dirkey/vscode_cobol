import * as vscode from "vscode";
import { ICOBOLSettings } from "./iconfiguration";

export enum TextLanguage {
    Unknown,
    COBOL,
    JCL
}

export class VSExtensionUtils {
    private static readonly knownSchemes = [
        "file",
        "ftp",
        "git",
        "member",
        "sftp",
        "ssh",
        "streamfile",
        "untitled",
        "vscode-vfs",
        "zip",
    ];

    private static readonly bmsColumnNumber = /^([0-9]{6}).*$/g;

    public static isSupportedLanguage(document: vscode.TextDocument): TextLanguage {
        switch (document.languageId.toLowerCase()) {
            case "cobolit":
            case "bitlang-cobol":
            case "cobol":
            case "acucobol":
            case "rmcobol":
            case "ilecobol":
                return TextLanguage.COBOL;
            case "jcl":
                return TextLanguage.JCL;
            default:
                return TextLanguage.Unknown;
        }
    }

    public static isKnownScheme(scheme: string): boolean {
        return this.knownSchemes.includes(scheme);
    }

    public static isKnownCOBOLLanguageId(config: ICOBOLSettings, langId: string): boolean {
        return config.valid_cobol_language_ids.includes(langId);
    }

    public static isKnownPLILanguageId(_: ICOBOLSettings, langId: string): boolean {
        return langId.toLowerCase() === "pli";
    }

    public static getAllCobolSelectors(
        config: ICOBOLSettings,
        forIntelliSense: boolean
    ): vscode.DocumentSelector {
        const ids = forIntelliSense
            ? config.valid_cobol_language_ids_for_intellisense
            : config.valid_cobol_language_ids;

        return ids.flatMap(langId =>
            this.knownSchemes.map(scheme => ({ scheme, language: langId }))
        );
    }

    public static getAllMFProvidersSelectors(_: ICOBOLSettings): vscode.DocumentSelector {
        return this.knownSchemes.map(scheme => ({ scheme, language: "directivesmf" }));
    }

    public static getAllCobolSelector(langId: string): vscode.DocumentSelector {
        return this.knownSchemes.map(scheme => ({ scheme, language: langId }));
    }

    public static getAllJCLSelectors(_: ICOBOLSettings): vscode.DocumentSelector {
        return this.knownSchemes.map(scheme => ({ scheme, language: "JCL" }));
    }

    public static flipPlaintext(doc: vscode.TextDocument): void {
        if (!doc) return;

        const { languageId, lineCount, uri } = doc;

        // Detect AcuBench COBOL
        if (languageId.toLowerCase() === "cobol" && lineCount >= 3) {
            const firstLine = doc.lineAt(0).text;
            const secondLine = doc.lineAt(1).text;
            if (firstLine.includes("*{Bench}") || secondLine.includes("*{Bench}")) {
                vscode.languages.setTextDocumentLanguage(doc, "ACUCOBOL");
                return;
            }
        }

        // Detect COBOL list files
        if (["plaintext", "tsql"].includes(languageId) && lineCount >= 3) {
            const firstLine = doc.lineAt(0).text;
            const secondLine = doc.lineAt(1).text;

            if (firstLine.charCodeAt(0) === 12 && secondLine.startsWith("* Micro Focus COBOL ")) {
                vscode.languages.setTextDocumentLanguage(doc, "COBOL_MF_LISTFILE");
                return;
            }

            if (firstLine.startsWith("Pro*COBOL: Release")) {
                vscode.languages.setTextDocumentLanguage(doc, "COBOL_PCOB_LISTFILE");
                return;
            }

            if (firstLine.includes("ACUCOBOL-GT ") && firstLine.includes("Page:")) {
                vscode.languages.setTextDocumentLanguage(doc, "COBOL_ACU_LISTFILE");
                return;
            }
        }

        // Detect BMS map files
        if (uri.fsPath.endsWith(".map") && !languageId.startsWith("bms")) {
            const maxLines = Math.min(lineCount, 10);
            for (let i = 0; i < maxLines; i++) {
                const line = doc.lineAt(i).text;
                if (line.includes("DFHMSD")) {
                    const targetLang = line.match(this.bmsColumnNumber) ? "bmsmap" : "bms";
                    vscode.languages.setTextDocumentLanguage(doc, targetLang);
                    return;
                }
            }
        }
    }
}