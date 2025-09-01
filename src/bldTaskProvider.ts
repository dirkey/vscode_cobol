import path, { dirname } from "path";
import * as vscode from "vscode";
import { VSWorkspaceFolders } from "./vscobolfolders";
import { COBOLFileUtils } from "./fileutils";
import { ICOBOLSettings } from "./iconfiguration";
import { VSCOBOLConfiguration } from "./vsconfiguration";

interface BldScriptDefinition extends vscode.TaskDefinition {
    arguments: string;
}

export class BldScriptTaskProvider implements vscode.TaskProvider {
    private bldScriptPromise?: Thenable<vscode.Task[]>;

    static readonly scriptPlatform = COBOLFileUtils.isWin32 ? "Windows Batch" : "Unix Script";
    static readonly BldScriptType = "COBOLBuildScript";
    static readonly BldSource = `${BldScriptTaskProvider.scriptPlatform} File`;
    private static readonly scriptName = COBOLFileUtils.isWin32 ? "bld.bat" : "./bld.sh";
    public static readonly scriptPrefix = COBOLFileUtils.isWin32 ? "cmd.exe /c " : "";

    public provideTasks(): Thenable<vscode.Task[]> | undefined {
        if (!vscode.workspace.isTrusted) return;

        const settings = VSCOBOLConfiguration.get_workspace_settings();
        const scriptFile = this.getFileFromWorkspace(settings);
        if (!scriptFile) return;

        if (!this.bldScriptPromise) {
            this.bldScriptPromise = createBldScriptTasks(scriptFile);
        }
        return this.bldScriptPromise;
    }

    public resolveTask(task: vscode.Task): vscode.Task | undefined {
        if (!vscode.workspace.isTrusted) return;

        const settings = VSCOBOLConfiguration.get_workspace_settings();
        const scriptFile = this.getFileFromWorkspace(settings);
        if (!scriptFile) return;

        const definition = task.definition as BldScriptDefinition;
        if (definition.arguments === undefined) return;

        const she = new vscode.ShellExecution(
            `${BldScriptTaskProvider.scriptPrefix}${scriptFile} ${definition.arguments}`,
            BldScriptTaskProvider.getSHEOptions(scriptFile)
        );

        const resolvedTask = new vscode.Task(
            definition,
            vscode.TaskScope.Workspace,
            BldScriptTaskProvider.BldSource,
            BldScriptTaskProvider.BldScriptType,
            she,
            BldScriptTaskProvider.getProblemMatchers()
        );

        const dname = path.dirname(scriptFile);
        const fname = `${scriptFile.substring(dname.length + 1)} (in ${dname})`;
        resolvedTask.detail = `Execute ${fname}`;

        return resolvedTask;
    }

    private getFileFromWorkspace(settings: ICOBOLSettings): string | undefined {
        const folders = VSWorkspaceFolders.get(settings);
        if (!folders) return undefined;

        for (const folder of folders) {
            if (folder.uri.scheme !== "file") continue;

            const fullPath = path.join(folder.uri.fsPath, BldScriptTaskProvider.scriptName);
            try {
                if (COBOLFileUtils.isFile(fullPath)) return fullPath;
            } catch {
                continue;
            }
        }

        return undefined;
    }

    public static getSHEOptions(scriptPath: string): vscode.ShellExecutionOptions {
        return { cwd: dirname(scriptPath) };
    }

    public static getProblemMatchers(): string[] {
        const matchers: string[] = [];
        if (process.env.ACUCOBOL) matchers.push("$acucobol-ccbl");
        if (process.env.COБDIR) matchers.push("$mfcobol-errformat3");
        if (process.env.COBOLITDIR) matchers.push("$cobolit-cobc");
        return matchers;
    }
}

async function createBldScriptTasks(scriptFile: string): Promise<vscode.Task[]> {
    const she = new vscode.ShellExecution(
        `${BldScriptTaskProvider.scriptPrefix}${scriptFile}`,
        BldScriptTaskProvider.getSHEOptions(scriptFile)
    );

    const taskDef: BldScriptDefinition = {
        type: BldScriptTaskProvider.BldScriptType,
        arguments: ""
    };

    const task = new vscode.Task(
        taskDef,
        vscode.TaskScope.Workspace,
        BldScriptTaskProvider.BldSource,
        BldScriptTaskProvider.BldScriptType,
        she,
        BldScriptTaskProvider.getProblemMatchers()
    );

    const dname = path.dirname(scriptFile);
    const fname = scriptFile.substring(dname.length + 1);
    task.detail = `Execute ${fname}`;
    task.group = vscode.TaskGroup.Build;

    return [task];
}