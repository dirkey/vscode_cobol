import * as vscode from "vscode";

export class VSHelpAndFeedback implements vscode.TreeDataProvider<FeedbackItem> {
    private readonly items: FeedbackItem[] = [
        { label: "Get Started", icon: "star", url: "https://github.com/spgennard/vscode_cobol" },
        { label: "Read Documentation", icon: "book", url: "https://github.com/spgennard/vscode_cobol#readme" },
        { label: "Review Issues", icon: "issues", url: "https://github.com/spgennard/vscode_cobol/issues" },
        { label: "Report Issue", icon: "comment", url: "https://github.com/spgennard/vscode_cobol/issues/new/choose" },
        { label: "Join the 'Rocket Software' Community", icon: "organization", url: "https://community.rocketsoftware.com/forums/forum-home?CommunityKey=fc99efc3-3189-48e3-860f-01928efec020" },
        { label: "Review extension", icon: "book", url: "https://marketplace.visualstudio.com/items?itemName=bitlang.cobol&ssr=false#review-details" },
        { label: "'Rocket Software' On-Demand Courses", icon: "play-circle", url: "https://www.rocketsoftware.com/learn-cobol" },
        { label: "Introduction to OO Programming", icon: "file-pdf", url: "https://docs-be.rocketsoftware.com/bundle/enterprisedeveloper_dg5_100_pdf/raw/resource/enus/enterprise_developer_intro_to_oo_programming_for_cobol_developers_vvc70.pdf" }
    ];

    private _onDidChangeTreeData = new vscode.EventEmitter<FeedbackItem | undefined | null | void>();
    readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

    constructor() {
        const treeView = vscode.window.createTreeView("help-and-feedback-view", { treeDataProvider: this });
        treeView.onDidChangeSelection(e => {
            const item = e.selection[0];
            if (item?.url) vscode.commands.executeCommand("vscode.open", item.url);
        });
    }

    getTreeItem(element: FeedbackItem): vscode.TreeItem {
        return {
            label: element.label,
            collapsibleState: vscode.TreeItemCollapsibleState.None,
            iconPath: new vscode.ThemeIcon(element.icon)
        };
    }

    getChildren(element?: FeedbackItem): vscode.ProviderResult<FeedbackItem[]> {
        return element ? [] : this.items;
    }
}

interface FeedbackItem {
    label: string;
    icon: string;
    url: string;
}