import * as vscode from "vscode";
import { ICOBOLSettings, intellisenseStyle } from "./iconfiguration";
import { VSExtensionUtils } from "./vsextutis";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { COBOLProgramCommands } from "./cobolprogram";
import { TabUtils } from "./tabstopper";
import { VSLogger } from "./vslogger";
import { AlignStyle, VSCOBOLUtils, FoldAction } from "./vscobolutils";
import { commands, ExtensionContext } from "vscode";
import { VSPPCodeLens } from "./vsppcodelens";
import { ExtensionDefaults } from "./extensionDefaults";
import { COBOLSourceScanner } from "./cobolsourcescanner";
import path from "path";
import fs from "fs";
import { VSWorkspaceFolders } from "./vscobolfolders";
import { VSDiagCommands } from "./vsdiagcommands";
import { CopyBookDragDropProvider } from "./vscopybookdragdroprovider";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { COBOLHierarchyProvider } from "./vscallhierarchyprovider";
import { newFile_dot_callgraph, view_dot_callgraph } from "./vsdotmarkdown";


export async function createCobolFile(
    title: string,
    doclang: string,
    config: ICOBOLSettings,
    template?: string
): Promise<void> {
    let fpath = "";
    let fdir = "";

    const ws = VSWorkspaceFolders.get(config);
    fdir = ws ? ws[0].uri.fsPath : process.cwd();

    const fileName = await vscode.window.showInputBox({
        title,
        prompt: `In directory : ${fdir}`,
        value: "untitled",
        validateInput: (text: string): string | undefined => {
            if (!text || !COBOLSourceScanner.isValidLiteral(text)) {
                return "Invalid program name";
            }

            fpath = path.join(fdir, `${text}.cbl`);
            if (fs.existsSync(fpath)) {
                return `File already exists (${fpath})`;
            }
            return undefined;
        }
    });

    if (!fileName) {
        return; // user canceled
    }

    fpath = path.join(fdir, `${fileName}.cbl`);
    const furl = vscode.Uri.file(fpath).with({ scheme: "untitled" });

    const document = await vscode.workspace.openTextDocument(furl);
    const editor = await vscode.window.showTextDocument(document);

    if (editor) {
        if (template) {
            // Load template from configuration
            const lines = vscode.workspace.getConfiguration().get<string[]>(template, []);
            if (lines.length > 0) {
                const snippet = new vscode.SnippetString(lines.join("\n"));
                await editor.insertSnippet(snippet, new vscode.Range(0, 0, lines.length + 1, 0));
            }
        }
        await vscode.languages.setTextDocumentLanguage(document, doclang);
    }
}

function isLSPActive(
    document: vscode.TextDocument,
    configSection: string,
    settingKey: string
): boolean {
    const settings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);

    const mfeditorConfig = VSWorkspaceFolders.get(settings)
        ? vscode.workspace.getConfiguration(configSection)
        : vscode.workspace.getConfiguration(configSection, document);

    return mfeditorConfig.get<boolean>(settingKey, true);
}

export function isMicroFocusCOBOL_LSPActive(document: vscode.TextDocument): boolean {
    return isLSPActive(document, "rocketCOBOL", "languageServerAutostart");
}

export function isMicroFocusPLI_LSPActive(document: vscode.TextDocument): boolean {
    return isLSPActive(document, "rocketPLI", "languageServer.autostart");
}

async function updateConfigSetting(
    settings: ICOBOLSettings,
    section: string,
    key: string,
    value: boolean
): Promise<void> {
    const config = vscode.workspace.getConfiguration(section);
    const target =
        VSWorkspaceFolders.get(settings) === undefined
            ? vscode.ConfigurationTarget.Global
            : vscode.ConfigurationTarget.Workspace;

    await config.update(key, value, target);
}

export async function setMicroFocusSuppressFileAssociationsPrompt(
    settings: ICOBOLSettings,
    onOrOff: boolean
): Promise<void> {
    if (settings.enable_rocket_cobol_lsp_when_active === false) {
        return;
    }

    await updateConfigSetting(settings, "rocketCOBOL", "suppressFileAssociationsPrompt", onOrOff);
}

export async function toggleMicroFocusLSP(
    settings: ICOBOLSettings,
    document: vscode.TextDocument,
    onOrOff: boolean
): Promise<void> {
    if (settings.enable_rocket_cobol_lsp_when_active === false) {
        return;
    }

    if (isMicroFocusCOBOL_LSPActive(document) !== onOrOff) {
        await updateConfigSetting(settings, "rocketCOBOL", "languageServerAutostart", onOrOff);
    }

    if (isMicroFocusPLI_LSPActive(document) !== onOrOff) {
        await updateConfigSetting(settings, "rocketPLI", "languageServer.autostart", onOrOff);
    }
}

const blessed_extensions: string[] = [
    "HCLTechnologies.hclappscancodesweep",          // code scanner
    ExtensionDefaults.rocketCOBOLExtension,         // Rocket COBOL extension
    ExtensionDefaults.rocketEnterpriseExtenstion,   // Rocket enterprise extension
    "micro-focus-amc.",                             // old micro focus extension's
    "bitlang.",                                     // mine
    "vscode.",                                      // vscode internal extensions
    "ms-vscode.",                                   //
    "ms-python.",                                   //
    "ms-vscode-remote.",
    "redhat.",                                      // redhat
    "rocketsoftware."                               // rockset software
];

const known_problem_extensions: [string, string, boolean][] = [
    ["bitlang.cobol already provides autocomplete and highlight for COBOL source code", "BroadcomMFD.cobol-language-support", true],
    ["A control flow extension that is not compatible with this dialect of COBOL", "BroadcomMFD.ccf", true],             // control flow extension
    ["COBOL debugger for different dialect of COBOL", "COBOLworx.cbl-gdb", true],
    ["Inline completion provider causes problems with this extension", "bloop.bloop-write", false],
    ["Language provider of COBOL/PLI that is not supported with extension", "heirloomcomputinginc", true],
    ["GnuCOBOL based utility extension, does support RocketCOBOL, ACU", "jsaila.coboler", false],
    ["GnuCOBOL based utility extension, does support RocketCOBOL, ACU", "zokugun.cobol-folding", false],
    ["COBOL Formatter that does not support RocketCOBOL","kopo-formatter", false]
];

function getExtensionInformation(
    grab_info_for_ext: vscode.Extension<any>,
    reasons: string[]
): string {
    const pkg = grab_info_for_ext.packageJSON;
    if (!pkg || pkg.publisher === "bitlang") {
        return "";
    }

    let message = "";

    if (pkg.id) {
        message += `\nThe extension ${pkg.name} from ${pkg.publisher} has conflicting functionality\n`;
        message += " Solution      : Disable or uninstall this extension, e.g. use command:\n";
        message += `                 code --uninstall-extension ${pkg.id}\n`;
    }

    if (reasons.length > 0) {
        const reasonLabel = reasons.length === 1 ? "Reason" : "Reasons";
        reasons.forEach((reason, i) => {
            message += `${i === 0 ? ` ${reasonLabel}       :` : "               :"} ${reason}\n`;
        });
    }

    if (pkg.id) {
        message += ` Id            : ${pkg.id}\n`;
        if (pkg.description) {
            message += ` Description   : ${pkg.description}\n`;
        }
        if (pkg.version) {
            message += ` Version       : ${pkg.version}\n`;
        }
        if (pkg.repository?.url) {
            message += ` Repository    : ${pkg.repository.url}\n`;
        }
        if (pkg.bugs?.url) {
            message += ` Bug Reporting : ${pkg.bugs.url}\n`;
        }
        if (pkg.bugs?.email) {
            message += ` Bug Email     : ${pkg.bugs.email}\n`;
        }
        message += "\n";
    }

    return message;
}


/**
 * Manages detection and handling of extension conflicts.
 */
export class ExtensionConflictManager {
    private settings: ICOBOLSettings;
    private context: vscode.ExtensionContext;

    constructor(settings: ICOBOLSettings, context: vscode.ExtensionContext) {
        this.settings = settings;
        this.context = context;
    }

    /**
     * Checks all installed extensions for conflicts and returns detailed information.
     */
    private checkExtensions(): { message: string; debuggerConflict: boolean; fatalConflict: boolean } {
        let dupExtensionMessage = "";
        let conflictingDebuggerFound = false;
        let fatalEditorConflict = false;

        for (const ext of vscode.extensions.all) {
            const reason: string[] = [];
            if (!ext?.packageJSON?.id || ext.packageJSON.id === ExtensionDefaults.thisExtensionName) {
                continue;
            }

            const idLower = ext.packageJSON.id.toLowerCase();

            // Skip blessed extensions
            if (this.isBlessedExtension(idLower)) continue;

            // Check known problem extensions
            for (const [type, knownId, editorConflict] of known_problem_extensions) {
                if (this.matchesExtensionId(idLower, knownId)) {
                    reason.push(`contributes '${type}'`);
                    if (type.includes("debugger")) conflictingDebuggerFound = true;
                    fatalEditorConflict = editorConflict;
                }
            }

            // Check categories, grammars, languages, debugger support
            this.analyzeExtension(ext, reason, () => { conflictingDebuggerFound = true; fatalEditorConflict = true; });

            if (reason.length > 0) {
                dupExtensionMessage += getExtensionInformation(ext, reason);
            }
        }

        return { message: dupExtensionMessage, debuggerConflict: conflictingDebuggerFound, fatalConflict: fatalEditorConflict };
    }

    /**
     * Public method to check for extension conflicts and take appropriate actions.
     */
    public checkForConflicts(): boolean {
        const { message, debuggerConflict, fatalConflict } = this.checkExtensions();

        if (!message) return false;

        VSLogger.logMessage(message);

        if (fatalConflict) {
            this.switchOpenAndFutureDocsToPlaintext();
        }

        vscode.window.showInformationMessage(
            `${ExtensionDefaults.thisExtensionName} detected duplicate or conflicting functionality`,
            { modal: true }
        ).then(() => VSLogger.logChannelSetPreserveFocus(false));

        if (debuggerConflict) {
            const msg = "This extension is now inactive until conflicts are resolved";
            VSLogger.logMessage(`\n${msg}\nRestart VS Code once resolved, or disable ${ExtensionDefaults.thisExtensionName}.`);
            this.logDebuggerAdvice();
            throw new Error(msg);
        }

        return false;
    }

    /**
     * Returns true if any conflict exists
     */
    public hasConflict(): boolean {
        const { message } = this.checkExtensions();
        return message.length > 0;
    }

    // ---- Helper Methods ---- //

    private isBlessedExtension(extIdLower: string): boolean {
        for (const blessed of blessed_extensions) {
            if (blessed.includes(".")) {
                if (blessed.toLowerCase() === extIdLower) return true;
            } else {
                if (extIdLower.startsWith(blessed.toLowerCase())) return true;
            }
        }
        return false;
    }

    private matchesExtensionId(idLower: string, knownId: string): boolean {
        if (knownId.includes(".")) return knownId.toLowerCase() === idLower;
        return idLower.startsWith(knownId.toLowerCase()) || idLower.endsWith(knownId.toLowerCase());
    }

    private analyzeExtension(ext: vscode.Extension<any>, reason: string[], markConflict: () => void) {
        let isDebugger = false;

        // Check categories
        const categories = ext.packageJSON.categories;
        if (categories) {
            for (const cat of categories) {
                if (`${cat}`.toUpperCase() === "DEBUGGERS") isDebugger = true;
            }
        }

        const contributes = ext.packageJSON.contributes;
        if (!contributes) return;

        // Grammars
        if (contributes.grammars) {
            for (const g of contributes.grammars) {
                const lang = g?.language?.toUpperCase();
                if (lang === ExtensionDefaults.defaultCOBOLLanguage) {
                    reason.push("contributes conflicting grammar (COBOL)");
                    markConflict();
                }
                if (lang === ExtensionDefaults.defaultPLIanguage) {
                    reason.push("contributes conflicting grammar (PLI)");
                    markConflict();
                }
            }
        }

        // Languages
        if (contributes.languages) {
            for (const l of contributes.languages) {
                const langId = l?.id?.toUpperCase();
                if (langId === ExtensionDefaults.defaultCOBOLLanguage) {
                    reason.push("contributes language id (COBOL)");
                    markConflict();
                }
                if (langId === ExtensionDefaults.defaultPLIanguage) {
                    reason.push("contributes language id (PLI)");
                    markConflict();
                }
            }
        }

        // Debugger / breakpoints
        if (isDebugger) {
            const debuggers = contributes.debuggers ?? [];
            for (const dbg of debuggers) {
                const langs = dbg.languages ?? [];
                if (langs.includes(ExtensionDefaults.defaultCOBOLLanguage)) {
                    reason.push(`extension includes a debugger for a different COBOL vendor -> ${dbg.label} of type ${dbg.type}`);
                    markConflict();
                }
            }

            const breakpoints = contributes.breakpoints ?? [];
            for (const bp of breakpoints) {
                if (bp?.language === ExtensionDefaults.defaultCOBOLLanguage) {
                    reason.push("extension includes debug breakpoint support for a different COBOL vendor");
                    markConflict();
                }
            }
        }
    }

    private switchOpenAndFutureDocsToPlaintext() {
        const switchToPlaintext = (doc: vscode.TextDocument) => {
            if (VSExtensionUtils.isKnownCOBOLLanguageId(this.settings, doc.languageId) ||
                VSExtensionUtils.isKnownPLILanguageId(this.settings, doc.languageId)) {
                VSLogger.logMessage(`Document ${doc.fileName} changed to plaintext to avoid errors`);
                vscode.languages.setTextDocumentLanguage(doc, "plaintext");
            }
        };

        vscode.window.visibleTextEditors.forEach(editor => switchToPlaintext(editor.document));

        const openDocHandler = vscode.workspace.onDidOpenTextDocument(switchToPlaintext);
        this.context.subscriptions.push(openDocHandler);
    }

    private logDebuggerAdvice() {
        const mfExt = vscode.extensions.getExtension(ExtensionDefaults.rocketCOBOLExtension);
        if (mfExt) {
            VSLogger.logMessage("You already have a 'Rocket COBOL' compatible debugger installed; above extension(s) may be unnecessary.");
        } else {
            VSLogger.logMessage(`Install a 'Rocket COBOL' compatible debugger using:\ncode --install-extension ${ExtensionDefaults.rocketCOBOLExtension}`);
        }
    }
}

/**
 * Top-level helper function to check for conflicts, can replace your old global method.
 */
export function checkForExtensionConflicts(settings: ICOBOLSettings, context: vscode.ExtensionContext): boolean {
    const manager = new ExtensionConflictManager(settings, context);
    return manager.checkForConflicts();
}

export function activateCommonCommands(context: vscode.ExtensionContext) {
    context.subscriptions.push(commands.registerCommand("cobolplugin.change_lang_to_acu", function () {
        const act = vscode.window.activeTextEditor;
        if (act === null || act === undefined) {
            return;
        }

        const settings = VSCOBOLConfiguration.get_resource_settings(act.document, VSExternalFeatures);
        vscode.languages.setTextDocumentLanguage(act.document, "ACUCOBOL");
        VSCOBOLUtils.enforceFileExtensions(settings, act, VSExternalFeatures, true, "ACUCOBOL");
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.change_lang_to_rmcobol", function () {
        const act = vscode.window.activeTextEditor;
        if (act === null || act === undefined) {
            return;
        }

        const settings = VSCOBOLConfiguration.get_resource_settings(act.document, VSExternalFeatures);
        vscode.languages.setTextDocumentLanguage(act.document, "RMCOBOL");
        VSCOBOLUtils.enforceFileExtensions(settings, act, VSExternalFeatures, true, "RMCOBOL");
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.change_lang_to_ilecobol", function () {
        const act = vscode.window.activeTextEditor;
        if (act === null || act === undefined) {
            return;
        }

        const settings = VSCOBOLConfiguration.get_resource_settings(act.document, VSExternalFeatures);
        vscode.languages.setTextDocumentLanguage(act.document, "ILECOBOL");
        VSCOBOLUtils.enforceFileExtensions(settings, act, VSExternalFeatures, true, "ILECOBOL");
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.change_lang_to_cobol", async function () {
        const act = vscode.window.activeTextEditor;
        if (act === null || act === undefined) {
            return;
        }

        // ensure all documents with the same id are change to the current ext id
        await VSCOBOLUtils.changeDocumentId(act.document.languageId, ExtensionDefaults.defaultCOBOLLanguage);

        const settings = VSCOBOLConfiguration.get_resource_settings(act.document, VSExternalFeatures);
        const mfExt = vscode.extensions.getExtension(ExtensionDefaults.rocketCOBOLExtension);
        if (mfExt) {
            await toggleMicroFocusLSP(settings, act.document, false);
        }

        VSCOBOLUtils.enforceFileExtensions(settings, act, VSExternalFeatures, true, ExtensionDefaults.defaultCOBOLLanguage);
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.change_lang_to_mfcobol", async function () {
        const act = vscode.window.activeTextEditor;
        if (act === null || act === undefined) {
            return;
        }

        const settings = VSCOBOLConfiguration.get_resource_settings(act.document, VSExternalFeatures);

        // ensure all documents with the same id are change to the 'Rocket COBOL lang id'
        await VSCOBOLUtils.changeDocumentId(act.document.languageId, ExtensionDefaults.microFocusCOBOLLanguageId);
        VSCOBOLUtils.enforceFileExtensions(settings, act, VSExternalFeatures, true, ExtensionDefaults.microFocusCOBOLLanguageId);

        await toggleMicroFocusLSP(settings, act.document, true);

        // invoke 'Micro Focus LSP Control'
        if (settings.enable_rocket_cobol_lsp_lang_server_control) {
            vscode.commands.executeCommand("mfcobol.languageServer.controls");
        }
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.move2pd", function () {
        COBOLProgramCommands.moveToProcedureDivision();
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.move2dd", function () {
        COBOLProgramCommands.moveToDataDivision();
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.move2ws", function () {
        COBOLProgramCommands.moveToWorkingStorage();
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.move2anyforward", function () {
        COBOLProgramCommands.moveForward();
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.move2anybackwards", function () {
        COBOLProgramCommands.moveBackward();
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.tab", async function () {
        await TabUtils.processTabKey(true);
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.revtab", async function () {
        await TabUtils.processTabKey(false);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.removeAllComments", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;

            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);
            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.removeComments(vscode.window.activeTextEditor);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.removeIdentificationArea", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.removeIdentificationArea(vscode.window.activeTextEditor);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.removeColumnNumbers", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.removeColumnNumbers(vscode.window.activeTextEditor);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makeKeywordsLowercase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.Keywords, langid, intellisenseStyle.LowerCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makeKeywordsUppercase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.Keywords, langid, intellisenseStyle.UpperCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makeKeywordsCamelCase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.Keywords, langid, intellisenseStyle.CamelCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makeFieldsLowercase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.ConstantsOrVariables, langid, intellisenseStyle.LowerCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makeFieldsUppercase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.ConstantsOrVariables, langid, intellisenseStyle.UpperCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makeFieldsCamelCase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.ConstantsOrVariables, langid, intellisenseStyle.CamelCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makePerformTargetsLowerCase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.PerformTargets, langid, intellisenseStyle.LowerCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makePerformTargetsUpperCase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.PerformTargets, langid, intellisenseStyle.UpperCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.makePerformTargetsCamelCase", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {
                VSCOBOLUtils.foldToken(VSExternalFeatures, settings, vscode.window.activeTextEditor, FoldAction.PerformTargets, langid, intellisenseStyle.CamelCase);
            }
        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.showCOBOLChannel", () => {
        VSLogger.logChannelSetPreserveFocus(true);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.resequenceColumnNumbers", () => {
        if (vscode.window.activeTextEditor) {
            const langid = vscode.window.activeTextEditor.document.languageId;
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (VSExtensionUtils.isKnownCOBOLLanguageId(settings, langid)) {

                vscode.window.showInputBox({
                    prompt: "Enter start line number and increment",
                    validateInput: (text: string): string | undefined => {
                        if (!text || text.indexOf(" ") === -1) {
                            return "You must enter two spaced delimited numbers (start increment)";
                        } else {
                            return undefined;
                        }
                    }
                }).then(value => {
                    // leave early
                    if (value === undefined) {
                        return;
                    }
                    const values: string[] = value.split(" ");
                    const startValue: number = Number.parseInt(values[0], 10);
                    const incrementValue: number = Number.parseInt(values[1], 10);
                    if (startValue >= 0 && incrementValue >= 1) {
                        VSCOBOLUtils.resequenceColumnNumbers(vscode.window.activeTextEditor, startValue, incrementValue);
                    } else {
                        vscode.window.showErrorMessage("Sorry invalid re-sequence given");
                    }
                });
            }
        }
    }));

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    context.subscriptions.push(commands.registerCommand("cobolplugin.ppcodelenaction", (args: string) => {
        VSPPCodeLens.actionCodeLens(args);
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.indentToCursor", () => {
        VSCOBOLUtils.indentToCursor();
    }));

    context.subscriptions.push(commands.registerCommand("cobolplugin.leftAdjustLine", () => {
        VSCOBOLUtils.leftAdjustLine();
    }));

    context.subscriptions.push(vscode.commands.registerTextEditorCommand("cobolplugin.transposeSelection", (textEditor, edit) => {
        VSCOBOLUtils.transposeSelection(textEditor, edit);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.alignStorageFirst", () => {
        VSCOBOLUtils.alignStorage(AlignStyle.First);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.alignStorageLeft", () => {
        VSCOBOLUtils.alignStorage(AlignStyle.Left);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.alignStorageCenter", () => {
        VSCOBOLUtils.alignStorage(AlignStyle.Center);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.alignStorageRight", () => {
        VSCOBOLUtils.alignStorage(AlignStyle.Right);
    }));
    vscode.commands.executeCommand("setContext", "cobolplugin.enableStorageAlign", true);

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.padTo72", () => {
        VSCOBOLUtils.padTo72();
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.enforceFileExtensions", () => {
        if (vscode.window.activeTextEditor) {
            const dialects = ["COBOL", "ACUCOBOL", "RMCOBOL", "ILECOBOL", "COBOLIT"];
            const mfExt = vscode.extensions.getExtension(ExtensionDefaults.rocketCOBOLExtension);
            const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

            if (mfExt !== undefined) {
                dialects.push(ExtensionDefaults.microFocusCOBOLLanguageId);
            }


            vscode.window.showQuickPick(dialects, { placeHolder: "Which Dialect do you prefer?" }).then(function (dialect) {
                if (vscode.window.activeTextEditor && dialect) {
                    VSCOBOLUtils.enforceFileExtensions(settings, vscode.window.activeTextEditor, VSExternalFeatures, true, dialect);
                }
            });

        }
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.selectionToCOBOLHEX", () => {
        VSCOBOLUtils.selectionToHEX(true);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.selectionToHEX", () => {
        VSCOBOLUtils.selectionToHEX(false);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.selectionHEXToASCII", () => {
        VSCOBOLUtils.selectionHEXToASCII();
    }));


    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.selectionToCOBOLNXHEX", () => {
        VSCOBOLUtils.selectionToNXHEX(true);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.selectionToNXHEX", () => {
        VSCOBOLUtils.selectionToNXHEX(false);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.newFile_BlankFile", async function () {
        if (vscode.window.activeTextEditor === undefined) {
            return;
        }
        const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);

        createCobolFile("Empty COBOL file", "COBOL", settings);
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.newFile_MicroFocus", async function () {
        if (vscode.window.activeTextEditor === undefined) {
            return;
        }
        const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);
        createCobolFile("COBOL program name?", "COBOL", settings, "coboleditor.template_rocket_cobol");
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.newFile_MicroFocus_mfunit", async function () {
        if (vscode.window.activeTextEditor === undefined) {
            return;
        }
        const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);
        createCobolFile("COBOL Unit Test program name?", "COBOL", settings, "coboleditor.template_rocket_cobol_mfunit");
    }));

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.newFile_ACUCOBOL", async function () {
        if (vscode.window.activeTextEditor === undefined) {
            return;
        }
        const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);
        createCobolFile("ACUCOBOL program name?", "ACUCOBOL", settings, "coboleditor.template_acucobol");
    }));


    const _settings = VSCOBOLConfiguration.get_workspace_settings();

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.dumpAllSymbols", async function () {
        await VSDiagCommands.DumpAllSymbols(_settings);
    }));

    const langIds = _settings.valid_cobol_language_ids;
    const mfExt = vscode.extensions.getExtension(ExtensionDefaults.rocketCOBOLExtension);
    if (mfExt) {
        langIds.push(ExtensionDefaults.microFocusCOBOLLanguageId);
    }

    for (const langid of langIds) {
        if (langid !== ExtensionDefaults.microFocusCOBOLLanguageId) {
            context.subscriptions.push(getLangStatusItem("Output Window", "cobolplugin.showCOBOLChannel", "Show", _settings, langid + "_1", langid));
        }

        switch (langid) {
            case "ACUCOBOL":
                context.subscriptions.push(getLangStatusItem("Switch to COBOL", "cobolplugin.change_lang_to_cobol", "Change", _settings, langid + "_2", langid));
                break;
            case "BITLANG-COBOL":
            case "COBOL":
                {
                    context.subscriptions.push(getLangStatusItem("Switch to ACUCOBOL", "cobolplugin.change_lang_to_acu", "Change", _settings, langid + "_3", langid));

                    if (mfExt !== undefined) {
                        context.subscriptions.push(getLangStatusItem("Switch to 'Rocket COBOL'", "cobolplugin.change_lang_to_mfcobol", "Change", _settings, langid + "_6", langid));
                    }
                }
                break;
            case "RMCOBOL":
                context.subscriptions.push(getLangStatusItem("Switch to ACUCOBOL", "cobolplugin.change_lang_to_acu", "Change", _settings, langid + "_2", langid));
                context.subscriptions.push(getLangStatusItem("Switch to COBOL", "cobolplugin.change_lang_to_cobol", "Change", _settings, langid + "_5", langid));
                break;
            case "ILECOBOL":
                context.subscriptions.push(getLangStatusItem("Switch to ILECOBOL", "cobolplugin.change_lang_to_ilecobol", "Change", _settings, langid + "_2", langid));
                context.subscriptions.push(getLangStatusItem("Switch to COBOL", "cobolplugin.change_lang_to_cobol", "Change", _settings, langid + "_5", langid));
                break;
            case ExtensionDefaults.microFocusCOBOLLanguageId:
                context.subscriptions.push(getLangStatusItem("Switch to 'BitLang COBOL'", "cobolplugin.change_lang_to_cobol", "Change", _settings, langid + "_6", langid));
                break;
        }

        context.subscriptions.push(vscode.languages.registerDocumentDropEditProvider(VSExtensionUtils.getAllCobolSelector(langid), new CopyBookDragDropProvider()));
    
    }

    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.dot_callgraph", async function () {
        if (vscode.window.activeTextEditor === undefined) {
            return;
        }
        const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);
        await newFile_dot_callgraph(settings);
    }));


    context.subscriptions.push(vscode.commands.registerCommand("cobolplugin.view_dot_callgraph", async function () {
        if (vscode.window.activeTextEditor === undefined) {
            return;
        }
        const settings = VSCOBOLConfiguration.get_resource_settings(vscode.window.activeTextEditor.document, VSExternalFeatures);
        await view_dot_callgraph(context,settings);
    }));

    if (_settings.enable_program_information) {        
        install_call_hierarchy(_settings, context)
    }
}


let installed_call_hierarchy:boolean = false;

export function install_call_hierarchy(_settings:ICOBOLSettings,  context: ExtensionContext) {
    // already installed
    if (installed_call_hierarchy) {
        return;
    }

    const langIds = _settings.valid_cobol_language_ids;
    for (const langid of langIds) {
        context.subscriptions.push(vscode.languages.registerCallHierarchyProvider(VSExtensionUtils.getAllCobolSelector(langid), new COBOLHierarchyProvider()));
    }
    installed_call_hierarchy=true;
}

function getLangStatusItem(text: string, command: string, title: string, settings: ICOBOLSettings, id: string, langid: string): vscode.LanguageStatusItem {
    const langStatusItem = vscode.languages.createLanguageStatusItem(id, VSExtensionUtils.getAllCobolSelector(langid));
    langStatusItem.text = text;
    langStatusItem.command = {
        command: command,
        title: title
    };
    return langStatusItem;
}

