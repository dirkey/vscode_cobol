import { FileType, Position, Range, Uri, workspace } from "vscode";
import { COBOLFileUtils } from "./fileutils";
import { ICOBOLSettings } from "./iconfiguration";
import { COBOLToken } from "./cobolsourcescanner";
import { ICOBOLSourceScanner } from "./icobolsourcescanner";

export class VSCOBOLSourceScannerTools {
    public static async howManyCopyBooksInDirectory(
        directory: string,
        settings: ICOBOLSettings
    ): Promise<number> {
        const folder = Uri.file(directory);
        const entries = await workspace.fs.readDirectory(folder);

        let copyBookCount = 0;

        for (const [entry, fileType] of entries) {
            if (
                (fileType & FileType.File) !== 0 &&
                COBOLFileUtils.isValidCopybookExtension(entry, settings)
            ) {
                copyBookCount++;
            }
        }

        return copyBookCount;
    }

    public static ignoreDirectory(partialName: string): boolean {
        // ignore hidden directories (e.g., .git, .vscode, etc.)
        return partialName.startsWith(".");
    }

    public static getExecToken(
        sf: ICOBOLSourceScanner,
        position: Position
    ): COBOLToken | undefined {
        for (const token of sf.execTokensInOrder) {
            const range = new Range(
                new Position(token.rangeStartLine, token.rangeStartColumn),
                new Position(token.rangeEndLine, token.rangeEndColumn)
            );

            if (range.contains(position)) {
                return token;
            }
        }

        return undefined;
    }

    public static isPositionInEXEC(
        sf: ICOBOLSourceScanner,
        position: Position
    ): boolean {
        return this.getExecToken(sf, position) !== undefined;
    }
}
