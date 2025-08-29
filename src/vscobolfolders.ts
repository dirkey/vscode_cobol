import { workspace, WorkspaceFolder } from "vscode";
import { ICOBOLSettings } from "./iconfiguration";

export class VSWorkspaceFolders {

    /** Get workspace folders filtered by "file" scheme with applied order */
    public static get(settings: ICOBOLSettings): ReadonlyArray<WorkspaceFolder> | undefined {
        return this.getFiltered("file", settings);
    }

    /** Get workspace folders filtered by schema with applied order */
    public static getFiltered(requiredSchema: string, settings: ICOBOLSettings): ReadonlyArray<WorkspaceFolder> | undefined {
        const wsFolders = workspace.workspaceFolders;
        if (!wsFolders) return undefined;

        // Filter folders by schema
        const filteredFolders = wsFolders.filter(f => !requiredSchema || f.uri.scheme === requiredSchema);
        const folderMap = new Map(filteredFolders.map(f => [f.name, f]));

        const ordered: WorkspaceFolder[] = [];

        // Add explicitly ordered folders first
        for (const name of settings.workspacefolders_order) {
            const folder = folderMap.get(name);
            if (folder) {
                ordered.push(folder);
                folderMap.delete(name);
            }
        }

        // Add remaining folders
        ordered.push(...folderMap.values());

        return ordered;
    }
}