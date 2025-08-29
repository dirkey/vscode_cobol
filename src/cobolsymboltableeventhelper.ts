import { InMemoryGlobalCacheHelper } from "./globalcachehelper";
import { COBOLToken, COBOLTokenStyle } from "./cobolsourcescanner";
import { ICOBOLSettings } from "./iconfiguration";
import { COBOLSymbol, COBOLSymbolTable } from "./cobolglobalcache";
import { COBOLWorkspaceSymbolCacheHelper } from "./cobolworkspacecache";
import {
    COBSCANNER_ADDFILE,
    COBSCANNER_KNOWNCOPYBOOK,
    COBSCANNER_SENDCLASS,
    COBSCANNER_SENDENUM,
    COBSCANNER_SENDEP,
    COBSCANNER_SENDINTERFACE,
    COBSCANNER_SENDPRGID
} from "./cobscannerdata";
import { ICOBOLSourceScanner, ICOBOLSourceScannerEventer, ICOBOLSourceScannerEvents } from "./icobolsourcescanner";

export class COBOLSymbolTableEventHelper implements ICOBOLSourceScannerEvents {
    private st?: COBOLSymbolTable;
    private readonly parse_copybooks_for_references: boolean;
    private readonly sender: ICOBOLSourceScannerEventer;

    constructor(config: ICOBOLSettings, sender: ICOBOLSourceScannerEventer) {
        this.sender = sender;
        this.parse_copybooks_for_references = config.parse_copybooks_for_references;
    }

    public start(qp: ICOBOLSourceScanner): void {
        this.st = new COBOLSymbolTable();
        this.st.fileName = qp.filename;
        this.st.lastModifiedTime = qp.lastModifiedTime;

        if (this.st.fileName && this.st.lastModifiedTime) {
            InMemoryGlobalCacheHelper.addFilename(this.st.fileName, qp.workspaceFile);
            this.sender.sendMessage(`${COBSCANNER_ADDFILE},${this.st.lastModifiedTime},${this.st.fileName}`);
        }

        COBOLWorkspaceSymbolCacheHelper.removeAllPrograms(this.st.fileName);
        COBOLWorkspaceSymbolCacheHelper.removeAllProgramEntryPoints(this.st.fileName);
        COBOLWorkspaceSymbolCacheHelper.removeAllTypes(this.st.fileName);
    }

    public processToken(token: COBOLToken): void {
        if (!this.st || token.ignoreInOutlineView) return;

        // Handle variable and label symbols if copybooks are not parsed
        if (!this.parse_copybooks_for_references) {
            switch (token.tokenType) {
                case COBOLTokenStyle.Union:
                case COBOLTokenStyle.Constant:
                case COBOLTokenStyle.ConditionName:
                case COBOLTokenStyle.Variable:
                    this.st.variableSymbols.set(token.tokenNameLower, new COBOLSymbol(token.tokenName, token.startLine));
                    break;
                case COBOLTokenStyle.Paragraph:
                case COBOLTokenStyle.Section:
                    this.st.labelSymbols.set(token.tokenNameLower, new COBOLSymbol(token.tokenName, token.startLine));
                    break;
            }
        }

        // Handle messaging for known copybooks and program/entry/interface/class tokens
        const messageMap: Partial<Record<COBOLTokenStyle, string>> = {
            [COBOLTokenStyle.CopyBook]: COBSCANNER_KNOWNCOPYBOOK,
            [COBOLTokenStyle.CopyBookInOrOf]: COBSCANNER_KNOWNCOPYBOOK,
            [COBOLTokenStyle.ImplicitProgramId]: COBSCANNER_SENDPRGID,
            [COBOLTokenStyle.ProgramId]: COBSCANNER_SENDPRGID,
            [COBOLTokenStyle.EntryPoint]: COBSCANNER_SENDEP,
            [COBOLTokenStyle.InterfaceId]: COBSCANNER_SENDINTERFACE,
            [COBOLTokenStyle.EnumId]: COBSCANNER_SENDENUM,
            [COBOLTokenStyle.ClassId]: COBSCANNER_SENDCLASS
        };

        const msgType = messageMap[token.tokenType];
        if (msgType && this.sender) {
            if (token.tokenType === COBOLTokenStyle.ImplicitProgramId) {
                COBOLWorkspaceSymbolCacheHelper.addCalableSymbol(this.st.fileName, token.tokenNameLower, token.startLine);
            }
            this.sender.sendMessage(`${msgType},${token.tokenName},${token.startLine},${this.st.fileName}`);
        }
    }

    public finish(): void {
        // No-op
    }
}