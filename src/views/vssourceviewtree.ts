import * as vscode from "vscode";

import { SourceOrFolderTreeItem } from "../sourceItem";
import { workspace } from "vscode";
import { ICOBOLSettings } from "../iconfiguration";
import { VSWorkspaceFolders } from "../vscobolfolders";
import { VSLogger } from "../vslogger";
import { VSCOBOLSourceScannerTools } from "../vssourcescannerutils";
import { VSCOBOLUtils } from "../vscobolutils";

let sourceTreeView: SourceViewTree | undefined;
let sourceTreeWatcher: vscode.FileSystemWatcher | undefined;

export class VSSourceTreeViewHandler {
    static async setupSourceViewTree(config: ICOBOLSettings, reinit: boolean): Promise<void> {
        if ((config.sourceview === false || reinit) && sourceTreeView) {
            sourceTreeWatcher?.dispose();
            sourceTreeView = undefined;
        }

        if (config.sourceview && !sourceTreeView) {
            sourceTreeView = new SourceViewTree(config);
            await sourceTreeView.init(config);

            sourceTreeWatcher = workspace.createFileSystemWatcher("**/*");
            sourceTreeWatcher.onDidCreate(uri => sourceTreeView?.checkFile(uri));
            sourceTreeWatcher.onDidDelete(uri => sourceTreeView?.clearFile(uri));

            vscode.window.registerTreeDataProvider("flat-source-view", sourceTreeView);
        }
    }

    static actionSourceViewItemFunction(si: SourceOrFolderTreeItem, debug: boolean): void {
        const fsPath = si?.uri?.path ?? "";
        VSCOBOLUtils.runOrDebug(fsPath, debug);
    }
}

export class SourceViewTree implements vscode.TreeDataProvider<SourceOrFolderTreeItem> {
    private readonly topLevelItems = new Map<string, SourceOrFolderTreeItem>();
    private readonly settings: ICOBOLSettings;

    private depth = 0;
    private readonly maxDepth = 5;

    private readonly _onDidChangeTreeData = new vscode.EventEmitter<SourceOrFolderTreeItem | undefined>();
    readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

    constructor(config: ICOBOLSettings) {
        this.settings = config;

        // Define all categories in one place
        const categories: [keyof ICOBOLSettings, string, boolean?][] = [
            ["program_extensions", "COBOL", true],
            ["copybookexts", "Copybooks", true],
            ["sourceview_include_jcl_files", "JCL"],
            ["sourceview_include_hlasm_files", "HLASM"],
            ["sourceview_include_pli_files", "PL/I"],
            ["sourceview_include_doc_files", "Documents"],
            ["sourceview_include_script_files", "Scripts"],
            ["sourceview_include_object_files", "Objects"],
            ["sourceview_include_test_files", "Tests"],
        ];

        for (const [flag, label, always] of categories) {
            if (always || (config[flag] as boolean)) {
                const item = new SourceOrFolderTreeItem(false, label);
                this.topLevelItems.set(label, item);
            }
        }
    }

    public async init(settings: ICOBOLSettings): Promise<void> {
        const folders = [
            ...(VSWorkspaceFolders.get(settings) || []),
            ...(VSWorkspaceFolders.getFiltered("", settings) || []),
        ];
        for (const folder of folders) {
            await this.addFolder(folder.uri);
        }
    }

    private async addFolder(uri: vscode.Uri): Promise<void> {
        if (this.depth >= this.maxDepth) return;

        this.depth++;
        const entries = await workspace.fs.readDirectory(uri);

        for (const [name, type] of entries) {
            const subUri = vscode.Uri.joinPath(uri, name);
            if (type === vscode.FileType.File) {
                const ext = name.split(".").pop();
                if (ext) this.addExtension(ext, subUri);
            } else if (type === vscode.FileType.Directory && !VSCOBOLSourceScannerTools.ignoreDirectory(name)) {
                await this.addFolder(subUri);
            }
        }

        this.depth--;
        this.refreshItems();
    }

    private getCommand(fileUri: vscode.Uri, ext: string): vscode.Command | undefined {
        if (["acu", "int", "gnt", "so", "dll"].includes(ext)) return undefined;
        return {
            arguments: [fileUri, { selection: new vscode.Range(0, 0, 0, 0) }],
            command: "vscode.open",
            title: "Open",
        };
    }

    private newSourceItem(contextValue: string, label: string, file: vscode.Uri, ext: string): SourceOrFolderTreeItem {
        const item = new SourceOrFolderTreeItem(true, label, file, 0);
        item.command = this.getCommand(file, ext);
        item.contextValue = contextValue;
        item.tooltip = file.fsPath;
        return item;
    }

    private refreshItems(): void {
        for (const item of this.topLevelItems.values()) {
            item.collapsibleState = item.children.size
                ? vscode.TreeItemCollapsibleState.Collapsed
                : vscode.TreeItemCollapsibleState.None;
        }
        this.refreshAll();
    }

    private refreshAll(): void {
        for (const item of this.topLevelItems.values()) {
            this._onDidChangeTreeData.fire(item);
        }
    }

    public async checkFile(uri: vscode.Uri): Promise<void> {
        try {
            const ext = uri.fsPath.split(".").pop();
            if (ext) {
                this.addExtension(ext, uri);
                this.refreshItems();
            }
        } catch (e) {
            VSLogger.logException("checkFile", e as Error);
        }
    }

    public clearFile(uri: vscode.Uri): void {
        const f = uri.fsPath;
        for (const item of this.topLevelItems.values()) {
            item.children.delete(f);
        }
        this.refreshItems();
    }

    private addExtension(ext: string, file: vscode.Uri): void {
        const base = file.fsPath.split(/[\\/]/).pop() ?? "";
        const extLower = ext.toLowerCase();

        // Handle special cases (tests, jcl, etc.)
        const isTest = base.startsWith("Test") || base.startsWith("MFU");
        const categories: Record<string, string[]> = {
            jcl: ["jcl", "job", "cntl", "prc", "proc"],
            hlasm: ["hlasm", "asm", "s", "asmpgm", "mac", "mlc", "asmmac"],
            pli: ["pl1", "pli", "plinc", "pc", "pci", "pcx", "inc"],
            document: ["lst", "md", "txt", "html"],
            scripts: ["sh", "bat"],
            objects: ["int", "gnt", "so", "dll", "acu"],
        };

        for (const [category, extensions] of Object.entries(categories)) {
            if (extensions.includes(extLower)) {
                const parent = this.topLevelItems.get(category.toUpperCase()) ?? this.topLevelItems.get(category);
                if (parent && !parent.children.has(file.fsPath)) {
                    parent.children.set(file.fsPath, this.newSourceItem(category, base, file, extLower));
                }
                return;
            }
        }

        // COBOL/Copybooks
        if (this.settings.program_extensions.includes(extLower)) {
            const cobol = this.topLevelItems.get("COBOL");
            cobol?.children.set(file.fsPath, this.newSourceItem("cobol", base, file, extLower));
        }
        if (this.settings.copybookexts.includes(extLower)) {
            const copybooks = this.topLevelItems.get("Copybooks");
            copybooks?.children.set(file.fsPath, this.newSourceItem("copybook", base, file, extLower));
        }

        // Tests
        if (isTest) {
            const tests = this.topLevelItems.get("Tests");
            if (tests && !tests.children.has(file.fsPath)) {
                tests.children.set(file.fsPath, this.newSourceItem("test", base, file, extLower));
            }
        }
    }

    public getChildren(element?: SourceOrFolderTreeItem): Thenable<SourceOrFolderTreeItem[]> {
        if (!element) {
            return Promise.resolve([...this.topLevelItems.values()]);
        }
        return Promise.resolve([...element.children.values()]);
    }

    getTreeItem(element: SourceOrFolderTreeItem): vscode.TreeItem {
        return element;
    }
}
