/* eslint-disable @typescript-eslint/naming-convention */
import path from "path";
import { extensions, Uri, WorkspaceFolder } from "vscode";
import { fork, ForkOptions } from "child_process";

import { VSWorkspaceFolders } from "./vscobolfolders";
import {
    COBSCANNER_ADDFILE,
    COBSCANNER_KNOWNCOPYBOOK,
    COBSCANNER_SENDCLASS,
    COBSCANNER_SENDENUM,
    COBSCANNER_SENDEP,
    COBSCANNER_SENDINTERFACE,
    COBSCANNER_SENDPRGID,
    COBSCANNER_STATUS,
    ScanData,
    ScanDataHelper
} from "./cobscannerdata";
import { progressStatusBarItem } from "./extension";
import { VSLogger } from "./vslogger";
import { ICOBOLSettings } from "./iconfiguration";
import { COBOLWorkspaceSymbolCacheHelper, TypeCategory } from "./cobolworkspacecache";
import { VSCOBOLUtils } from "./vscobolutils";
import { InMemoryGlobalCacheHelper, InMemoryGlobalSymbolCache } from "./globalcachehelper";
import { COBOLWorkspaceFile } from "./cobolglobalcache";
import { VSCOBOLFileUtils } from "./vsfileutils";
import { ExtensionDefaults } from "./extensionDefaults";

class FileScanStats {
    directoriesScanned = 0;
    maxDirectoryDepth = 0;
    fileCount = 0;
    showMessage = false;
    directoriesScannedMap: Map<string, Uri> = new Map<string, Uri>();
}

export class VSCobScanner {
    public static readonly scannerBinDir = VSCobScanner.getCobScannerDirectory();

    private static getCobScannerDirectory(): string {
        const thisExtension = extensions.getExtension(ExtensionDefaults.thisExtensionName);
        return thisExtension ? path.join(thisExtension.extensionPath, "dist") : "";
    }

    private static async forkScanner(
        settings: ICOBOLSettings,
        sf: ScanData,
        reason: string,
        updateNow: boolean,
        useThreaded: boolean,
        threadCount: number
    ): Promise<void> {
        const scannerPath = path.join(VSCobScanner.scannerBinDir, "cobscanner.js");

        const options: ForkOptions = {
            stdio: [0, 1, 2, "ipc"],
            cwd: VSCobScanner.scannerBinDir,
            env: {
                SCANDATA: ScanDataHelper.getScanData(sf),
                SCANDATA_TCOUNT: `${threadCount}`
            }
        };
        const scannerStyle = useThreaded ? "useenv_threaded" : "useenv";
        const child = fork(scannerPath, [scannerStyle], options);

        if (!child) return;

        const timeoutTimer = setTimeout(() => {
            try { child.kill(); } catch (err) {
                VSLogger.logException(`Timeout, ${reason}`, err as Error);
            }
        }, settings.cache_metadata_inactivity_timeout);

        child.on("error", err => VSLogger.logException(`Fork caused ${reason}`, err));
        child.on("exit", code => {
            clearTimeout(timeoutTimer);
            if (code !== 0 && sf.cache_metadata_verbose_messages) {
                VSLogger.logMessage(`External scan completed (${child.pid}) [Exit Code=${code}]`);
            } else {
                progressStatusBarItem.hide();
            }
            if (updateNow) VSCOBOLUtils.saveGlobalCacheToWorkspace(settings);
        });

        let percent = 0;
        child.on("message", (msg: unknown) => {
            timeoutTimer.refresh();
            if (typeof msg !== "string") return;

            InMemoryGlobalSymbolCache.isDirty = true;

            if (msg.startsWith("@@")) {
                VSCobScanner.handleScannerMessage(settings, msg);
                if (msg.startsWith(COBSCANNER_STATUS)) {
                    const args = msg.split(" ");
                    progressStatusBarItem.show();
                    percent += Number.parseInt(args[1], 10);
                    progressStatusBarItem.text = `Processing metadata: ${percent}%`;
                }
            } else {
                VSLogger.logMessage(msg);
            }
        });

        await VSCobScanner.readStreamLines(child.stdout, VSLogger.logMessage);
        await VSCobScanner.readStreamLines(child.stderr, line => VSLogger.logMessage(` [${line}]`));
    }

    private static async readStreamLines(stream: NodeJS.ReadableStream | null, lineHandler: (line: string) => void) {
        if (!stream) return;
        for await (const chunk of stream) {
            const text = chunk.toString();
            text.split("\n").forEach(line => {
                const trimmed = line.trim();
                if (trimmed.length) lineHandler(trimmed);
            });
        }
    }

    private static handleScannerMessage(settings: ICOBOLSettings, message: string) {
        const addClass = (tokenFilename: string, tokenName: string, tokenLine: number, type: TypeCategory) =>
            COBOLWorkspaceSymbolCacheHelper.addClass(tokenFilename, tokenName, tokenLine, type);

        const args = message.split(",");

        switch (true) {
            case message.startsWith(COBSCANNER_SENDEP):
                COBOLWorkspaceSymbolCacheHelper.addEntryPoint(args[3], args[1], Number.parseInt(args[2], 10));
                break;
            case message.startsWith(COBSCANNER_SENDPRGID):
                const linePRG = Number.parseInt(args[2], 10);
                COBOLWorkspaceSymbolCacheHelper.removeAllProgramEntryPoints(args[3]);
                COBOLWorkspaceSymbolCacheHelper.removeAllTypes(args[3]);
                COBOLWorkspaceSymbolCacheHelper.addCalableSymbol(args[3], args[1], linePRG);
                break;
            case message.startsWith(COBSCANNER_SENDCLASS):
                addClass(args[3], args[1], Number.parseInt(args[2], 10), TypeCategory.ClassId);
                break;
            case message.startsWith(COBSCANNER_SENDINTERFACE):
                addClass(args[3], args[1], Number.parseInt(args[2], 10), TypeCategory.InterfaceId);
                break;
            case message.startsWith(COBSCANNER_SENDENUM):
                addClass(args[3], args[1], Number.parseInt(args[2], 10), TypeCategory.EnumId);
                break;
            case message.startsWith(COBSCANNER_ADDFILE):
                VSCobScanner.addFileToCache(settings, args[2], BigInt(args[1]));
                break;
            case message.startsWith(COBSCANNER_KNOWNCOPYBOOK):
                COBOLWorkspaceSymbolCacheHelper.addReferencedCopybook(args[1], args[2]);
                break;
        }
    }

    private static addFileToCache(settings: ICOBOLSettings, fullFilename: string, ms: bigint) {
        const fsUri = Uri.file(fullFilename);
        const shortFilename = VSCOBOLFileUtils.getShortWorkspaceFilename(fsUri.scheme, fullFilename, settings);
        if (shortFilename) {
            const cws = new COBOLWorkspaceFile(ms, shortFilename);
            COBOLWorkspaceSymbolCacheHelper.removeAllProgramEntryPoints(shortFilename);
            InMemoryGlobalCacheHelper.addFilename(fullFilename, cws);
        } else {
            VSLogger.logMessage(`Unable to getShortWorkspaceFilename for ${fullFilename}`);
        }
    }

    private static getScanData(settings: ICOBOLSettings, ws: readonly WorkspaceFolder[], stats: FileScanStats, files: string[]): ScanData {
        const sf = new ScanData();
        sf.scannerBinDir = VSCobScanner.scannerBinDir;
        sf.directoriesScanned = stats.directoriesScanned;
        sf.maxDirectoryDepth = stats.maxDirectoryDepth;
        sf.fileCount = stats.fileCount;

        sf.parse_copybooks_for_references = settings.parse_copybooks_for_references;
        sf.Files = files;
        sf.cache_metadata_verbose_messages = settings.cache_metadata_verbose_messages;
        sf.md_symbols = settings.metadata_symbols;
        sf.md_entrypoints = settings.metadata_entrypoints;
        sf.md_metadata_files = settings.metadata_files;
        sf.md_metadata_knowncopybooks = settings.metadata_knowncopybooks;

        for (const f of ws) {
            if (f.uri.scheme === "file") sf.workspaceFolders.push(f.uri.fsPath);
        }

        return sf;
    }

    public static async processAllFilesInWorkspaceOutOfProcess(
        settings: ICOBOLSettings,
        viaCommand: boolean,
        useThreaded: boolean,
        threadCount: number
    ): Promise<void> {
        const msgViaCommand = `(${viaCommand ? "on demand" : "startup"})`;
        const ws = VSWorkspaceFolders.get(settings);
        const stats = new FileScanStats();
        const files: string[] = [];

        if (!ws) {
            VSLogger.logMessage(`No workspace folders available ${msgViaCommand}`);
            return;
        }

        if (!viaCommand) {
            VSLogger.logChannelHide();
        } else {
            VSLogger.logChannelSetPreserveFocus(!viaCommand);
        }

        VSLogger.logMessage("");
        VSLogger.logMessage(`Starting to process metadata from workspace folders ${msgViaCommand}`);

        await VSCOBOLUtils.populateDefaultCallableSymbols(settings, false);
        InMemoryGlobalSymbolCache.defaultCallableSymbols.forEach((b) => files.push(b));
        VSCOBOLUtils.saveGlobalCacheToWorkspace(settings, false);

        const sf = this.getScanData(settings, ws, stats, files);
        await VSCobScanner.forkScanner(settings, sf, msgViaCommand, true, useThreaded, threadCount);
        VSCOBOLUtils.saveGlobalCacheToWorkspace(settings, true);
    }
}