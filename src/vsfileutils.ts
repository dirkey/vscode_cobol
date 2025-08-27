import path from "path";
import fs from "fs";

import { VSWorkspaceFolders } from "./vscobolfolders";
import { Range, TextEditor, Uri, window, workspace, WorkspaceFolder } from "vscode";

import { ICOBOLSettings } from "./iconfiguration";
import { IExternalFeatures } from "./externalfeatures";
import { IVSCOBOLSettings } from "./vsconfiguration";

export class VSCOBOLFileUtils {
    public static isPathInWorkspace(ddir: string, config: ICOBOLSettings): boolean {
        const ws = VSWorkspaceFolders.get(config);
        if (!workspace || !ws) return false;

        const fullPath = Uri.file(ddir).fsPath;
        return ws.some(folder => folder.uri.fsPath === fullPath);
    }

    public static getFullWorkspaceFilename(
        features: IExternalFeatures,
        sdir: string,
        sdirMs: BigInt,
        config: ICOBOLSettings
    ): string | undefined {
        const ws = VSWorkspaceFolders.get(config);
        if (!workspace || !ws) return undefined;

        for (const folder of ws) {
            if (folder.uri.scheme !== "file") continue;

            const possibleFile = path.join(folder.uri.fsPath, sdir);
            if (features.isFile(possibleFile)) {
                const modTime = features.getFileModTimeStamp(possibleFile);
                if (sdirMs === modTime) {
                    return possibleFile;
                }
            }
        }
        return undefined;
    }

    public static getShortWorkspaceFilename(
        schema: string,
        ddir: string,
        config: ICOBOLSettings
    ): string | undefined {
        if (schema === "untitled") return undefined;

        const ws = VSWorkspaceFolders.get(config);
        if (!workspace || !ws) return undefined;

        const fullPath = Uri.file(ddir).fsPath;
        let bestShortName: string | undefined;

        for (const folder of ws) {
            if (folder.uri.scheme !== schema) continue;

            const folderPath = folder.uri.fsPath;
            if (fullPath.startsWith(folderPath)) {
                const relativePath = fullPath.substring(folderPath.length + 1);
                if (!bestShortName || relativePath.length < bestShortName.length) {
                    bestShortName = relativePath;
                }
            }
        }
        return bestShortName;
    }

    /**
     * Generic implementation for copybook finding
     */
    private static async findCopyBookGeneric(
        filename: string,
        dirs: string[],
        extensions: string[],
        features: IExternalFeatures,
        inDirectory = "",
        useAsync = false,
        isURL = false
    ): Promise<string> {
        if (!filename) return "";

        const hasDot = filename.includes(".");
        for (const baseDir of dirs) {
            const dir = inDirectory ? path.posix.join(baseDir, inDirectory) : baseDir;

            // Build candidate paths
            const candidates: string[] = [];
            if (isURL) {
                candidates.push(`${dir}/${filename}`);
                if (!hasDot) {
                    for (const ext of extensions) {
                        candidates.push(`${dir}/${filename}.${ext}`);
                    }
                }
            } else {
                candidates.push(path.join(dir, filename));
                if (!hasDot) {
                    for (const ext of extensions) {
                        candidates.push(path.join(dir, `${filename}.${ext}`));
                    }
                }
            }

            // Test candidates
            for (const candidate of candidates) {
                if (useAsync) {
                    if (await features.isFileASync(candidate)) return candidate;
                } else {
                    if (features.isFile(candidate)) return candidate;
                }
            }
        }
        return "";
    }

    // Local (sync) versions
    public static findCopyBook(filename: string, config: IVSCOBOLSettings, features: IExternalFeatures): string {
        return (this.findCopyBookGeneric(
            filename,
            config.file_search_directory,
            config.copybookexts,
            features,
            "",
            false,
            false
        ) as unknown) as string;
    }

    public static findCopyBookInDirectory(
        filename: string,
        inDirectory: string,
        config: IVSCOBOLSettings,
        features: IExternalFeatures
    ): string {
        return (this.findCopyBookGeneric(
            filename,
            config.file_search_directory,
            config.copybookexts,
            features,
            inDirectory,
            false,
            false
        ) as unknown) as string;
    }

    // Remote (async) versions
    public static async findCopyBookViaURL(
        filename: string,
        config: ICOBOLSettings,
        features: IExternalFeatures
    ): Promise<string> {
        return this.findCopyBookGeneric(
            filename,
            features.getURLCopyBookSearchPath(),
            config.copybookexts,
            features,
            "",
            true,
            true
        );
    }

    public static async findCopyBookInDirectoryViaURL(
        filename: string,
        inDirectory: string,
        config: ICOBOLSettings,
        features: IExternalFeatures
    ): Promise<string> {
        return this.findCopyBookGeneric(
            filename,
            features.getURLCopyBookSearchPath(),
            config.copybookexts,
            features,
            inDirectory,
            true,
            true
        );
    }

    public static extractSelectionToCopybook(
        activeTextEditor: TextEditor,
        features: IExternalFeatures
    ): void {
        const sel = activeTextEditor.selection;
        const ran = new Range(sel.start, sel.end);
        const text = activeTextEditor.document.getText(ran);
        const dir = path.dirname(activeTextEditor.document.fileName);

        window.showInputBox({
            prompt: "Copybook name?",
            validateInput: (copybookFilename: string): string | undefined => {
                const invalid =
                    !copybookFilename ||
                    /\s|\./.test(copybookFilename) ||
                    features.isFile(path.join(dir, `${copybookFilename}.cpy`));
                return invalid ? "Invalid copybook" : undefined;
            }
        }).then(copybookFilename => {
            if (!copybookFilename) return;

            const filename = path.join(dir, `${copybookFilename}.cpy`);
            fs.writeFileSync(filename, text);
            activeTextEditor.edit(edit => {
                edit.replace(ran, `           copy "${copybookFilename}.cpy".`);
            });
        });
    }

    public static getBestWorkspaceFolder(workspaceDirectory: string): WorkspaceFolder | undefined {
        const workspaces = workspace.workspaceFolders;
        if (!workspaces) return undefined;

        // Pick the deepest matching workspace
        return workspaces
            .filter(ws => workspaceDirectory.startsWith(ws.uri.fsPath))
            .sort((a, b) => b.uri.fsPath.length - a.uri.fsPath.length)[0];
    }
}