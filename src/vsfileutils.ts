import path from "path";
import fs from "fs";
import { VSWorkspaceFolders } from "./vscobolfolders";
import { Range, TextEditor, Uri, window, workspace, WorkspaceFolder } from "vscode";
import { ICOBOLSettings } from "./iconfiguration";
import { IExternalFeatures } from "./externalfeatures";

export class VSCOBOLFileUtils {
    /** Checks if a directory is part of the workspace */
    public static isPathInWorkspace(ddir: string, config: ICOBOLSettings): boolean {
        const ws = VSWorkspaceFolders.get(config);
        if (!workspace || !ws) return false;

        const fullPath = Uri.file(ddir).fsPath;
        return ws.some(folder => folder.uri.fsPath === fullPath);
    }

    /** Returns full path for workspace file if its modification timestamp matches */
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
            if (features.isFile(possibleFile) && sdirMs === features.getFileModTimeStamp(possibleFile)) {
                return possibleFile;
            }
        }
        return undefined;
    }

    /** Returns shortest relative workspace path for a given file */
    public static getShortWorkspaceFilename(schema: string, ddir: string, config: ICOBOLSettings): string | undefined {
        if (schema === "untitled") return undefined;

        const ws = VSWorkspaceFolders.get(config);
        if (!workspace || !ws) return undefined;

        const fullPath = Uri.file(ddir).fsPath;
        let bestShortName: string | undefined;

        for (const folder of ws) {
            if (folder.uri.scheme !== schema) continue;

            const folderPath = folder.uri.fsPath;
            if (fullPath.startsWith(folderPath)) {
                const relativePath = fullPath.slice(folderPath.length + 1);
                if (!bestShortName || relativePath.length < bestShortName.length) {
                    bestShortName = relativePath;
                }
            }
        }
        return bestShortName;
    }

    /** Generic copybook search (sync or async, local or remote) */
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

            const candidates = isURL
                ? [filename, ...(!hasDot ? extensions.map(ext => `${filename}.${ext}`) : [])].map(f => `${dir}/${f}`)
                : [filename, ...(!hasDot ? extensions.map(ext => `${filename}.${ext}`) : [])].map(f => path.join(dir, f));

            for (const candidate of candidates) {
                if (useAsync ? await features.isFileASync(candidate) : features.isFile(candidate)) {
                    return candidate;
                }
            }
        }
        return "";
    }

    /** Local (sync) search for copybooks */
    public static findCopyBook(filename: string, config: ICOBOLSettings, features: IExternalFeatures): string {
        return this.findCopyBookGeneric(filename, config.file_search_directory, config.copybookexts, features) as unknown as string;
    }

    public static findCopyBookInDirectory(filename: string, inDirectory: string, config: ICOBOLSettings, features: IExternalFeatures): string {
        return this.findCopyBookGeneric(filename, config.file_search_directory, config.copybookexts, features, inDirectory) as unknown as string;
    }

    /** Remote (async) search via URL */
    public static async findCopyBookViaURL(filename: string, config: ICOBOLSettings, features: IExternalFeatures): Promise<string> {
        return this.findCopyBookGeneric(filename, features.getURLCopyBookSearchPath(), config.copybookexts, features, "", true, true);
    }

    public static async findCopyBookInDirectoryViaURL(filename: string, inDirectory: string, config: ICOBOLSettings, features: IExternalFeatures): Promise<string> {
        return this.findCopyBookGeneric(filename, features.getURLCopyBookSearchPath(), config.copybookexts, features, inDirectory, true, true);
    }

    /** Extract selected text to a new copybook file */
    public static extractSelectionToCopybook(activeTextEditor: TextEditor, features: IExternalFeatures): void {
        const sel = activeTextEditor.selection;
        const ran = new Range(sel.start, sel.end);
        const text = activeTextEditor.document.getText(ran);
        const dir = path.dirname(activeTextEditor.document.fileName);

        window.showInputBox({
            prompt: "Copybook name?",
            validateInput: (name: string) => {
                const invalid = !name || /\s|\./.test(name) || features.isFile(path.join(dir, `${name}.cpy`));
                return invalid ? "Invalid copybook" : undefined;
            }
        }).then(copybookFilename => {
            if (!copybookFilename) return;

            const filename = path.join(dir, `${copybookFilename}.cpy`);
            fs.writeFileSync(filename, text);
            activeTextEditor.edit(edit => edit.replace(ran, `           copy "${copybookFilename}.cpy".`));
        });
    }

    /** Returns the deepest matching workspace folder for a given directory */
    public static getBestWorkspaceFolder(workspaceDirectory: string): WorkspaceFolder | undefined {
        const workspaces = workspace.workspaceFolders;
        if (!workspaces) return undefined;

        return workspaces
            .filter(ws => workspaceDirectory.startsWith(ws.uri.fsPath))
            .sort((a, b) => b.uri.fsPath.length - a.uri.fsPath.length)[0];
    }
}