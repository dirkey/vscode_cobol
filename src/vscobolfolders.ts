import { workspace, WorkspaceFolder } from "vscode";
import { ICOBOLSettings } from "./iconfiguration";

export class VSWorkspaceFolders {
    /**
     * Get workspace folders filtered by "file" scheme with applied order.
     */
    public static get(settings: ICOBOLSettings): ReadonlyArray<WorkspaceFolder> | undefined {
        return this.getFiltered("file", settings);
    }

    /**
     * Get workspace folders filtered by schema with applied order.
     */
    public static getFiltered(requiredSchema: string, settings: ICOBOLSettings): ReadonlyArray<WorkspaceFolder> | undefined {
        const ws = workspace.workspaceFolders;
        if (!ws) return undefined;

        // Build a map of workspace folders, optionally filtering by schema
        const folderMap = new Map<string, WorkspaceFolder>(
            ws
                .filter(folder => !requiredSchema || folder.uri.scheme === requiredSchema)
                .map(folder => [folder.name, folder])
        );

        const orderedFolders: WorkspaceFolder[] = [];  
        const { workspacefolders_order: foldersOrder } = settings;
        // Add explicitly ordered folders first
        for (const name of foldersOrder) {
            const folder = folderMap.get(name);
            if (folder) {
                orderedFolders.push(folder);
                folderMap.delete(name);
            }
        }

        // Add remaining folders
        orderedFolders.push(...folderMap.values());

        return orderedFolders;
    }
}