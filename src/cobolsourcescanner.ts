/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { ISourceHandler, ICommentCallback, ISourceHandlerLite } from "./isourcehandler";
import { cobolProcedureKeywordDictionary, cobolStorageKeywordDictionary, getCOBOLKeywordDictionary } from "./keywords/cobolKeywords";

import { FileSourceHandler } from "./filesourcehandler";
import { COBOLFileAndColumnSymbol, COBOLFileSymbol, COBOLWorkspaceFile } from "./cobolglobalcache";

import { ICOBOLSettings } from "./iconfiguration";
import { CobolLinterProviderSymbols, ESourceFormat, IExternalFeatures } from "./externalfeatures";

import * as path from "path";
import { SourceFormat } from "./sourceformat";
import { ExtensionDefaults } from "./extensionDefaults";
import { SplitTokenizer } from "./splittoken";
import { SourcePorter, PortResult } from "./vsdirectivesconv";
import { ICOBOLSourceScanner, ICOBOLSourceScannerEvents } from "./icobolsourcescanner";

export enum COBOLTokenStyle {
    CopyBook = "Copybook",
    CopyBookInOrOf = "CopybookInOrOf",
    File = "File",
    ProgramId = "Program-Id",
    ImplicitProgramId = "ImplicitProgramId-Id",
    FunctionId = "Function-Id",
    EndFunctionId = "EndFunctionId",
    Constructor = "Constructor",
    MethodId = "Method-Id",
    Property = "Property",
    ClassId = "Class-Id",
    InterfaceId = "Interface-Id",
    ValueTypeId = "Valuetype-Id",
    EnumId = "Enum-id",
    Section = "Section",
    Paragraph = "Paragraph",
    Division = "Division",
    EntryPoint = "Entry",
    Variable = "Variable",
    ConditionName = "ConditionName",
    RenameLevel = "RenameLevel",
    Constant = "Constant",
    Union = "Union",
    EndDelimiter = "EndDelimiter",
    Exec = "Exec",
    EndExec = "EndExec",
    Declaratives = "Declaratives",
    EndDeclaratives = "EndDeclaratives",
    Region = "Region",
    SQLCursor = "SQLCursor",
    Unknown = "Unknown",
    IgnoreLS = "IgnoreLS",
    Null = "Null"
}

export class SourceScannerUtils {
    /**
     * Converts a string to camel case, removing dashes and underscores.
     * Example: "hello-world_test" -> "HelloWorldTest"
     */
    public static camelize(text: string): string {
        if (!text) {
            return "";
        }
        return text
            .split(/[-_]+/)
            .filter(Boolean)
            .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
            .join("");
    }

    /**
     * Converts a string to lower camel case.
     * Example: "hello-world_test" -> "helloWorldTest"
     */
    public static lowerCamelize(text: string): string {
        const camel = SourceScannerUtils.camelize(text);
        return camel.length > 0 ? camel.charAt(0).toLowerCase() + camel.slice(1) : "";
    }

    /**
     * Checks if a string is camel case.
     */
    public static isCamelCase(text: string): boolean {
        return /^[A-Z][a-zA-Z0-9]*$/.test(text);
    }
}

export class COBOLToken {
    public ignoreInOutlineView = false;
    public readonly filenameAsURI: string;
    public readonly filename: string;
    public readonly tokenType: COBOLTokenStyle;

    public readonly startLine: number;
    public startColumn: number;
    public endLine: number;
    public endColumn: number;

    public rangeStartLine: number;
    public rangeStartColumn: number;
    public rangeEndLine: number;
    public rangeEndColumn: number;

    public readonly tokenName: string;
    public readonly tokenNameLower: string;

    public description: string;
    public readonly parentToken?: COBOLToken;
    public readonly inProcedureDivision: boolean;

    public readonly extraInformation1: string;
    public inSection?: COBOLToken;

    public readonly sourceHandler: ISourceHandler;
    public readonly isFromScanCommentsForReferences: boolean;
    public readonly isImplicitToken: boolean;

    constructor(
        sourceHandler: ISourceHandler,
        filenameAsURI: string,
        filename: string,
        tokenType: COBOLTokenStyle,
        startLine: number,
        startColumn: number,
        token: string,
        description: string,
        parentToken: COBOLToken | undefined,
        inProcedureDivision: boolean,
        extraInformation1: string,
        isFromScanCommentsForReferences: boolean,
        isImplicitToken: boolean
    ) {
        this.sourceHandler = sourceHandler;
        this.filenameAsURI = filenameAsURI;
        this.filename = filename;
        this.tokenType = tokenType;
        this.startLine = startLine;
        this.startColumn = Math.max(0, startColumn);
        this.tokenName = token.trim();
        this.tokenNameLower = this.tokenName.toLowerCase();
        this.endLine = this.startLine;
        this.endColumn = this.startColumn + this.tokenName.length;

        this.description = description;
        this.parentToken = parentToken;
        this.inProcedureDivision = inProcedureDivision;
        this.extraInformation1 = extraInformation1;
        this.inSection = undefined;

        this.isFromScanCommentsForReferences = isFromScanCommentsForReferences;
        this.isImplicitToken = isImplicitToken;

        this.rangeStartLine = this.startLine;
        this.rangeStartColumn = this.startColumn;
        this.rangeEndLine = this.endLine;
        this.rangeEndColumn = this.endColumn;
    }
}

export class COBOLVariable {
    public readonly ignoreInOutlineView: boolean;
    public readonly token: COBOLToken;
    public readonly tokenType: COBOLTokenStyle;

    constructor(token: COBOLToken) {
        this.token = token;
        this.tokenType = token.tokenType;
        this.ignoreInOutlineView = token.ignoreInOutlineView;
    }
}

export class SQLDeclare {
    public readonly token: COBOLToken;
    public readonly currentLine: number;
    public readonly sourceReferences: SourceReference[] = [];
    public readonly ignoreInOutlineView: boolean;

    constructor(token: COBOLToken, currentLine: number) {
        this.token = token;
        this.currentLine = currentLine;
        this.ignoreInOutlineView = token.ignoreInOutlineView;
    }
}

export class SourceReference_Via_Length {
    constructor(
        public readonly fileIdentifer: number,
        public readonly line: number,
        public readonly column: number,
        public readonly length: number,
        public tokenStyle: COBOLTokenStyle,
        public readonly isFromScanCommentsForReferences: boolean,
        public readonly name: string,
        public readonly reason: string
    ) {}

    get nameLower(): string {
        return this.name.toLowerCase();
    }
}

export class SourceReference {
    constructor(
        public readonly fileIdentifer: number,
        public readonly line: number,
        public readonly column: number,
        public readonly endLine: number,
        public readonly endColumn: number,
        public readonly tokenStyle: COBOLTokenStyle
    ) {}
}

class StreamToken {
    constructor(
        public currentToken: string,
        public currentTokenLower: string,
        public endsWithDot: boolean,
        public currentLine: number,
        public currentCol: number
    ) {}

    static readonly Blank = new StreamToken("", "", false, 0, 0);
}

class StreamTokens implements Iterable<StreamToken> {
    public currentToken = "";
    public currentTokenLower = "";
    public currentCol = 0;
    public currentLineNumber: number;
    public endsWithDot = false;

    public prevToken = "";
    public prevTokenLower = "";
    public prevTokenLineNumber = 0;
    public prevCurrentCol = 0;

    private tokenIndex = 0;
    private readonly stokens: StreamToken[] = [];

    public static readonly Blank = new StreamTokens("", 0);

    constructor(line: string, lineNumber: number, previousToken?: StreamTokens) {
        this.currentLineNumber = lineNumber;

        if (previousToken) {
            this.prevToken = previousToken.currentToken;
            this.prevTokenLineNumber = previousToken.currentLineNumber;
        }

        this.tokenize(line, lineNumber);
        this.setupNextToken();

        if (previousToken && previousToken.stokens.length > 0) {
            const last = previousToken.stokens.at(-1)!;
            this.prevToken = last.currentToken;
            this.prevTokenLower = last.currentTokenLower;
            this.prevCurrentCol = last.currentCol;
        }
    }

    private tokenize(line: string, lineNumber: number): void {
        const lineTokens: string[] = [];
        SplitTokenizer.splitArgument(line, lineTokens);

        let rollingColumn = 0;
        for (const token of lineTokens) {
            const lower = token.toLowerCase();
            rollingColumn = line.indexOf(token, rollingColumn);

            const endsWithDot = token.endsWith(".");
            this.stokens.push(
                new StreamToken(token, lower, endsWithDot, lineNumber, rollingColumn)
            );

            rollingColumn += token.length;
        }
    }

    public nextSTokenOrBlank(): StreamToken {
        return this.stokens[this.tokenIndex + 1] ?? StreamToken.Blank;
    }

    public nextSTokenIndex(offset: number): StreamToken {
        return this.stokens[this.tokenIndex + offset] ?? StreamToken.Blank;
    }

    public compoundItems(startCompound: string): string {
        if (this.endsWithDot || this.tokenIndex + 1 >= this.stokens.length) {
            return startCompound;
        }

        let comp = startCompound;
        let addNext = false;

        for (const stok of this.stokens.slice(this.tokenIndex + 1)) {
            const trimmed = COBOLSourceScanner.trimLiteral(stok.currentToken, false);

            if (stok.endsWithDot) {
                return `${comp} ${trimmed}`;
            }

            if (addNext) {
                comp += ` ${trimmed}`;
                addNext = false;
            } else if (stok.currentToken === "&") {
                comp += ` ${trimmed}`;
                addNext = true;
            } else {
                return comp;
            }
        }

        return comp;
    }

    private setupNextToken(): void {
        this.prevToken = this.currentToken;
        this.prevTokenLower = this.currentTokenLower;
        this.prevTokenLineNumber = this.currentLineNumber;
        this.prevCurrentCol = this.currentCol;

        const stok = this.stokens[this.tokenIndex];
        if (stok) {
            this.currentToken = stok.currentToken;
            this.currentTokenLower = stok.currentTokenLower;
            this.endsWithDot = stok.endsWithDot;
            this.currentCol = stok.currentCol;
            this.currentLineNumber = stok.currentLine;
        } else {
            this.currentToken = this.currentTokenLower = "";
            this.endsWithDot = false;
            this.currentCol = 0;
            this.currentLineNumber = 0;
        }
    }

    public moveToNextToken(): boolean {
        if (this.tokenIndex + 1 > this.stokens.length) {
            return true;
        }
        this.tokenIndex++;
        this.setupNextToken();
        return false;
    }

    public endToken(): void {
        this.tokenIndex = this.stokens.length;
    }

    public isTokenPresent(possibleToken: string): boolean {
        const lower = possibleToken.toLowerCase();
        return this.stokens.some(
            t => t.currentTokenLower === lower || t.currentTokenLower === `${lower}.`
        );
    }

    // 🔹 Generator-based iterator
    public *[Symbol.iterator](): Generator<StreamToken> {
        for (const token of this.stokens) {
            yield token;
        }
    }
}

export class ReplaceToken {
    public readonly pattern: RegExp;

    constructor(rawToken: string, tokenState: IReplaceState) {
        const escaped = ReplaceToken.escapeRegExp(rawToken);
        this.pattern = tokenState.isPseudoTextDelimiter
            ? new RegExp(escaped, "g")
            : new RegExp(`\\b${escaped}\\b`, "g");
    }

    private static escapeRegExp(text: string): string {
        // Escapes regex metacharacters so they are treated literally
        return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
}

export interface IReplaceState {
    isPseudoTextDelimiter: boolean;
}

export class replaceState implements IReplaceState {
    public isPseudoTextDelimiter = false;
}

export class copybookState implements IReplaceState {
    public sourceHandler: ISourceHandler | undefined = undefined;
    public copyBook = "";
    public trimmedCopyBook = "";
    public isIn = false;
    public isOf = false;
    public isReplacing = false;
    public isReplacingBy = false;
    public startLineNumber = 0;
    public endLineNumber = 0;
    public endCol = 0;
    public startCol = 0;
    public line = "";
    public copyVerb = "";
    public literal2 = "";
    public library_name = "";
    public replaceLeft = "";
    public copyReplaceMap = new Map<string, ReplaceToken>();
    public isTrailing = false;
    public isLeading = false;
    public fileName = "";
    public fileNameMod: BigInt = BigInt(0);
    public isPseudoTextDelimiter = false;
    public saved01Group: COBOLToken | undefined;
    public copybookDepths: copybookState[] = [];
    constructor(current01Group: COBOLToken | undefined) {
        this.saved01Group = current01Group;
    }
}

export class COBOLCopybookToken {
    public readonly token?: COBOLToken;
    public scanComplete: boolean;
    public statementInformation?: copybookState;

    public static readonly Null = new COBOLCopybookToken(undefined, false, undefined);

    constructor(token?: COBOLToken, scanComplete = false, statementInformation?: copybookState) {
        this.token = token;
        this.scanComplete = scanComplete;
        this.statementInformation = statementInformation;
    }

    public hasCopybookChanged(features: IExternalFeatures, configHandler: ICOBOLSettings): boolean {
        const info = this.statementInformation;
        if (!info) {
            return false;
        }

        const fileName = info.fileName;
        if (!features.isFile(fileName)) {
            return true;
        }
        if (features.getFileModTimeStamp(fileName) !== info.fileNameMod) {
            return true;
        }

        if (this.token) {
            const expandedFileName = features.expandLogicalCopyBookToFilenameOrEmpty(
                this.token.tokenName,
                this.token.extraInformation1,
                this.token.sourceHandler,
                configHandler
            );
            if (expandedFileName !== fileName) {
                return true;
            }
        }

        return false;
    }
}

export class SharedSourceReferences {
    public filenames: string[] = [];
    public filenameURIs: string[] = [];

    public readonly targetReferences = new Map<string, SourceReference_Via_Length[]>();
    public readonly constantsOrVariablesReferences = new Map<string, SourceReference_Via_Length[]>();
    public readonly unknownReferences = new Map<string, SourceReference_Via_Length[]>();
    public readonly ignoreLSRanges: SourceReference[] = [];

    public readonly sharedConstantsOrVariables = new Map<string, COBOLVariable[]>();
    public readonly sharedSections = new Map<string, COBOLToken>();
    public readonly sharedParagraphs = new Map<string, COBOLToken>();
    public readonly copyBooksUsed = new Map<string, COBOLCopybookToken[]>();
    public readonly execSQLDeclare = new Map<string, SQLDeclare>();
    public state: ParseState;
    public tokensInOrder: COBOLToken[] = [];

    public readonly ignoreUnusedSymbol = new Map<string, string>();

    public topLevel: boolean;
    public startTime: number;

    constructor(configHandler: ICOBOLSettings, topLevel: boolean, startTime: number) {
        this.state = new ParseState(configHandler);
        this.topLevel = topLevel;
        this.startTime = startTime;
    }

    public reset(configHandler: ICOBOLSettings): void {
        this.filenames = [];
        this.targetReferences.clear();
        this.constantsOrVariablesReferences.clear();
        this.unknownReferences.clear();
        this.sharedConstantsOrVariables.clear();
        this.sharedSections.clear();
        this.sharedParagraphs.clear();
        this.copyBooksUsed.clear();
        this.execSQLDeclare.clear();
        this.state = new ParseState(configHandler);
        this.tokensInOrder = [];
        this.ignoreUnusedSymbol.clear();
        this.ignoreLSRanges.length = 0;
    }

    public getSourceFieldId(handFilename: string): number {
        return this.filenames.indexOf(handFilename);
    }

    private getReferenceInformation4(
        refMap: Map<string, SourceReference_Via_Length[]>,
        sourceFileId: number,
        variable: string,
        startLine: number,
        startColumn: number
    ): [number, number] {
        let defvars = refMap.get(variable) ?? refMap.get(variable.toLowerCase());
        if (!defvars) {
            return [0, 0];
        }
        let definedCount = 0;
        let referencedCount = 0;
        for (const defvar of defvars) {
            if (defvar.line === startLine && defvar.column === startColumn && defvar.fileIdentifer === sourceFileId) {
                definedCount++;
            } else {
                referencedCount++;
            }
        }
        return [definedCount, referencedCount];
    }

    public getReferenceInformation4variables(variable: string, sourceFileId: number, startLine: number, startColumn: number): [number, number] {
        return this.getReferenceInformation4(this.constantsOrVariablesReferences, sourceFileId, variable, startLine, startColumn);
    }

    public getReferenceInformation4targetRefs(variable: string, sourceFileId: number, startLine: number, startColumn: number): [number, number] {
        return this.getReferenceInformation4(this.targetReferences, sourceFileId, variable, startLine, startColumn);
    }
}

export enum UsingState {
    BY_VALUE,
    BY_REF,
    BY_CONTENT,
    BY_OUTPUT,
    RETURNING,
    UNKNOWN
}

export class COBOLParameter {
    constructor(
        public readonly using: UsingState,
        public readonly name: string
    ) {}
}

export class ParseState {
    currentToken?: COBOLToken;
    currentRegion?: COBOLToken;
    currentDivision?: COBOLToken;
    currentSection?: COBOLToken;
    currentSectionOutRefs = new Map<string, SourceReference_Via_Length[]>();
    currentParagraph?: COBOLToken;
    currentClass?: COBOLToken;
    currentMethod?: COBOLToken;
    currentFunctionId?: COBOLToken;
    current01Group?: COBOLToken;
    currentLevel?: COBOLToken;
    currentProgramTarget: CallTargetInformation;
    copyBooksUsed = new Map<string, COBOLCopybookToken[]>();
    procedureDivision?: COBOLToken;
    declaratives?: COBOLToken;
    captureDivisions = true;
    programs: COBOLToken[] = [];
    pickFields = false;
    pickUpUsing = false;
    skipToDot = false;
    endsWithDot = false;
    prevEndsWithDot = false;
    currentLineIsComment = false;
    skipNextToken = false;
    inValueClause = false;
    skipToEndLsIgnore = false;
    inProcedureDivision = false;
    inDeclaratives = false;
    inReplace = false;
    replace_state = new replaceState();
    inCopy = false;
    replaceLeft = "";
    replaceRight = "";
    captureReplaceLeft = true;
    ignoreInOutlineView = false;
    addReferencesDuringSkipToTag = false;
    addVariableDuringStipToTag = false;
    using: UsingState = UsingState.BY_REF;
    parameters: COBOLParameter[] = [];
    entryPointCount = 0;
    replaceMap = new Map<string, ReplaceToken>();
    enable_text_replacement: boolean;
    copybook_state: copybookState;
    inCopyStartColumn = 0;
    restorePrevState = false;

    constructor(configHandler: ICOBOLSettings) {
        this.enable_text_replacement = configHandler.enable_text_replacement;
        this.copybook_state = new copybookState(undefined);
        this.currentProgramTarget = new CallTargetInformation("", undefined, false, []);
    }
}

class PreParseState {
    numberTokensInHeader = 0;
    workingStorageRelatedTokens = 0;
    procedureDivisionRelatedTokens = 0;
    sectionsInToken = 0;
    divisionsInToken = 0;
    leaveEarly = false;
}

export class CallTargetInformation {
    public token?: COBOLToken;
    public originalToken: string;
    public isEntryPoint: boolean;
    public callParameters: COBOLParameter[];

    constructor(originalToken: string, token?: COBOLToken, isEntryPoint = false, callParameters: COBOLParameter[] = []) {
        this.originalToken = originalToken;
        this.token = token;
        this.isEntryPoint = isEntryPoint;
        this.callParameters = callParameters;
    }
}


export class EmptyCOBOLSourceScannerEventHandler implements ICOBOLSourceScannerEvents {
    static readonly Default = new EmptyCOBOLSourceScannerEventHandler();

    start(qp: ICOBOLSourceScanner): void {
        return;
    }

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    processToken(token: COBOLToken): void {
        return;
    }

    finish(): void {
        return;
    }
}

export class COBOLSourceScanner implements ICommentCallback, ICOBOLSourceScanner, ICOBOLSourceScanner {
    public id: string;
    public readonly sourceHandler: ISourceHandler;
    public filename: string;
    // eslint-disable-next-line @typescript-eslint/ban-types
    public lastModifiedTime: BigInt = BigInt(0);

    public tokensInOrder: COBOLToken[] = [];
    public execTokensInOrder: COBOLToken[] = [];
    public readonly execSQLDeclare: Map<string, SQLDeclare>;

    public readonly sections: Map<string, COBOLToken>;
    public readonly paragraphs: Map<string, COBOLToken>;
    public readonly constantsOrVariables: Map<string, COBOLVariable[]>;
    public readonly callTargets: Map<string, CallTargetInformation>;
    public readonly functionTargets: Map<string, CallTargetInformation>;
    public readonly classes: Map<string, COBOLToken>;
    public readonly methods: Map<string, COBOLToken>;
    public readonly copyBooksUsed: Map<string, COBOLCopybookToken[]>;

    public readonly diagMissingFileWarnings: Map<string, COBOLFileSymbol>;
    public readonly portWarnings: PortResult[];
    public readonly generalWarnings: COBOLFileSymbol[];

    public readonly commentReferences: COBOLFileAndColumnSymbol[];

    public readonly parse4References: boolean;
    public readonly sourceReferences: SharedSourceReferences;
    public readonly sourceFileId: number;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public cache4PerformTargets: any | undefined = undefined;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public cache4ConstantsOrVars: any | undefined = undefined;

    public ImplicitProgramId = "";
    public ProgramId = "";

    public sourceFormat: ESourceFormat = ESourceFormat.unknown;

    public sourceIsCopybook = false;

    public workspaceFile: COBOLWorkspaceFile;

    public readonly parse_copybooks_for_references: boolean;
    public readonly scan_comments_for_hints: boolean;
    public readonly isFromScanCommentsForReferences: boolean;
    public readonly scan_comment_for_ls_control: boolean;

    readonly copybookNestedInSection: boolean;

    readonly configHandler: ICOBOLSettings;

    parseHint_OnOpenFiles: string[] = [];
    parseHint_WorkingStorageFiles: string[] = [];
    parseHint_LocalStorageFiles: string[] = [];
    parseHint_ScreenSectionFiles: string[] = [];

    public eventHandler: ICOBOLSourceScannerEvents;
    public externalFeatures: IExternalFeatures;

    public scanAborted: boolean;

    private readonly languageId: string;
    private readonly usePortationSourceScanner: boolean;
    private currentExecToken: COBOLToken | undefined = undefined;
    private currentExec = "";
    private currentExecVerb = "";

    private readonly COBOLKeywordDictionary: Map<string, string>;

    private readonly sourcePorter: SourcePorter = new SourcePorter();

    private readonly activeRegions: COBOLToken[] = [];

    private readonly regions: COBOLToken[] = [];

    private implicitCount = 0;

    public static ScanUncached(sourceHandler: ISourceHandler,
        configHandler: ICOBOLSettings,
        parse_copybooks_for_references: boolean,
        eventHandler: ICOBOLSourceScannerEvents,
        externalFeatures: IExternalFeatures
    ): ICOBOLSourceScanner {

        const startTime = externalFeatures.performance_now();
        return new COBOLSourceScanner(
            startTime,
            sourceHandler,
            configHandler,
            new SharedSourceReferences(configHandler, true, startTime),
            parse_copybooks_for_references,
            eventHandler,
            externalFeatures,
            false
        );
    }


    private ScanUncachedInlineCopybook(
        sourceHandler: ISourceHandler,
        parentSource: ICOBOLSourceScanner,
        isFromScanCommentsForReferences: boolean
    ): boolean {
        const configHandler = parentSource.configHandler;
        const sharedSource = parentSource.sourceReferences;

        const parse_copybooks_for_references = parentSource.parse_copybooks_for_references;
        const eventHandler = parentSource.eventHandler;
        const externalFeatures = parentSource.externalFeatures;

        const state: ParseState = parentSource.sourceReferences.state;

        const prevIgnoreInOutlineView: boolean = state.ignoreInOutlineView;
        const prevEndsWithDot = state.endsWithDot;
        const prevPrevEndsWithDot = state.prevEndsWithDot;
        const prevCurrentDivision = state.currentDivision;
        const prevCurrentSection = state.currentSection;
        const prevProcedureDivision = state.procedureDivision;
        const prevPickFields = state.pickFields;
        const prevSkipToDot = state.skipToDot;

        state.current01Group = undefined;
        state.restorePrevState = false;
        state.restorePrevState = true;
        state.ignoreInOutlineView = true;

        new COBOLSourceScanner(
            sharedSource.startTime,
            sourceHandler,
            configHandler,
            sharedSource,
            parse_copybooks_for_references,
            eventHandler,
            externalFeatures,
            isFromScanCommentsForReferences);

        // unless the state has been replaces
        if (state.current01Group === undefined) {
            state.current01Group = state.copybook_state.saved01Group;
        }
        state.ignoreInOutlineView = prevIgnoreInOutlineView;
        state.endsWithDot = prevEndsWithDot;
        state.pickFields = prevPickFields;
        state.skipToDot = prevSkipToDot;
        if (state.restorePrevState) {
            state.currentDivision = prevCurrentDivision;
            state.currentSection = prevCurrentSection;
            state.procedureDivision = prevProcedureDivision;
            state.prevEndsWithDot = prevPrevEndsWithDot;
        }

        return true;
    }

    public constructor(
        startTime: number,
        sourceHandler: ISourceHandler, configHandler: ICOBOLSettings,
        sourceReferences: SharedSourceReferences = new SharedSourceReferences(configHandler, true, startTime),
        parse_copybooks_for_references: boolean,
        sourceEventHandler: ICOBOLSourceScannerEvents,
        externalFeatures: IExternalFeatures,
        isFromScanCommentsForReferences: boolean) {
        const filename = sourceHandler.getFilename();

        this.sourceHandler = sourceHandler;
        this.id = sourceHandler.getUriAsString();
        this.configHandler = configHandler;
        this.filename = path.normalize(filename);
        this.ImplicitProgramId = path.basename(filename, path.extname(filename));
        this.parse_copybooks_for_references = parse_copybooks_for_references;
        this.eventHandler = sourceEventHandler;
        this.externalFeatures = externalFeatures;
        this.scan_comments_for_hints = configHandler.scan_comments_for_hints;
        this.isFromScanCommentsForReferences = isFromScanCommentsForReferences;

        this.copybookNestedInSection = configHandler.copybooks_nested;
        this.scan_comment_for_ls_control = configHandler.scan_comment_for_ls_control;
        this.copyBooksUsed = new Map<string, COBOLCopybookToken[]>();
        this.sections = new Map<string, COBOLToken>();
        this.paragraphs = new Map<string, COBOLToken>();
        this.constantsOrVariables = new Map<string, COBOLVariable[]>();
        this.callTargets = new Map<string, CallTargetInformation>();
        this.functionTargets = new Map<string, CallTargetInformation>();
        this.classes = new Map<string, COBOLToken>();
        this.methods = new Map<string, COBOLToken>();
        this.diagMissingFileWarnings = new Map<string, COBOLFileSymbol>();
        this.portWarnings = [];
        this.generalWarnings = [];
        this.commentReferences = [];
        this.parse4References = sourceHandler !== null;
        this.cache4PerformTargets = undefined;
        this.cache4ConstantsOrVars = undefined;
        this.scanAborted = false;
        this.languageId = this.sourceHandler.getLanguageId();
        this.COBOLKeywordDictionary = getCOBOLKeywordDictionary(this.languageId);
        switch (this.languageId.toLocaleLowerCase()) {
            case "cobol":
            case "bitlang-cobol":
                this.usePortationSourceScanner = configHandler.linter_port_helper;
                break;
            default:
                this.usePortationSourceScanner = false;
                break;
        }

        let sourceLooksLikeCOBOL = false;
        let prevToken: StreamTokens = StreamTokens.Blank;

        const hasCOBOLExtension = path.extname(filename).length > 0 ? true : false;
        this.sourceReferences = sourceReferences;

        this.sourceFileId = sourceReferences.filenames.length;
        sourceReferences.filenames.push(sourceHandler.getFilename());
        sourceReferences.filenameURIs.push(sourceHandler.getUriAsString());

        this.constantsOrVariables = sourceReferences.sharedConstantsOrVariables;
        this.paragraphs = sourceReferences.sharedParagraphs;
        this.sections = sourceReferences.sharedSections;
        this.tokensInOrder = sourceReferences.tokensInOrder;
        this.copyBooksUsed = sourceReferences.copyBooksUsed;
        this.execSQLDeclare = sourceReferences.execSQLDeclare;

        // set the source handler for the comment parsing
        sourceHandler.addCommentCallback(this);

        const state: ParseState = this.sourceReferences.state;

        /* mark this has been processed (to help copy of self) */
        state.copyBooksUsed.set(this.filename, [COBOLCopybookToken.Null]);
        if (this.sourceReferences.topLevel) {
            this.lastModifiedTime = externalFeatures.getFileModTimeStamp(this.filename);
        }

        this.workspaceFile = new COBOLWorkspaceFile(this.lastModifiedTime, sourceHandler.getShortWorkspaceFilename());

        // setup the event handler
        if (this.sourceReferences.topLevel) {
            this.eventHandler = sourceEventHandler;
            this.eventHandler.start(this);
        }

        const maxLineLength = configHandler.editor_maxTokenizationLineLength;

        if (this.sourceReferences.topLevel) {
            /* if we have an extension, then don't do a relaxed parse to determine if it is COBOL or not */
            const lineLimit = configHandler.pre_scan_line_limit;
            const maxLinesInFile = sourceHandler.getLineCount();
            let maxLines = maxLinesInFile;
            if (maxLines > lineLimit) {
                maxLines = lineLimit;
            }

            let line: string | undefined = undefined;
            const preParseState: PreParseState = new PreParseState();

            for (let l = 0; l < maxLines + sourceHandler.getCommentCount(); l++) {
                if (l > maxLinesInFile) {
                    break;
                }

                try {
                    line = sourceHandler.getLine(l, false);
                    if (line === undefined) {
                        break; // eof
                    }

                    line = line.trimEnd();

                    // ignore large lines
                    if (line.length > maxLineLength) {
                        this.externalFeatures.logMessage(`Aborted scanning ${this.filename} max line length exceeded`);
                        this.clearScanData();
                        continue;
                    }

                    // don't parse a empty line
                    if (line.length > 0) {
                        if (prevToken.endsWithDot === false) {
                            prevToken = this.relaxedParseLineByLine(prevToken, line, l, preParseState);
                        }
                        else {
                            prevToken = this.relaxedParseLineByLine(StreamTokens.Blank, line, l, preParseState);
                        }
                    } else {
                        maxLines++;     // increase the max lines, as this line is
                    }

                    if (preParseState.leaveEarly) {
                        break;
                    }

                }
                catch (e) {
                    this.externalFeatures.logException("COBOLScannner - Parse error : " + e, e as Error);
                }
            }

            // Do we have some sections?
            if (preParseState.sectionsInToken === 0 && preParseState.divisionsInToken === 0) {
                /* if we have items that could be in a data division */

                if (preParseState.procedureDivisionRelatedTokens !== 0 && preParseState.procedureDivisionRelatedTokens > preParseState.workingStorageRelatedTokens) {
                    this.ImplicitProgramId = "";

                    const fakeDivision = this.newCOBOLToken(COBOLTokenStyle.Division, 0, "Procedure Division", 0, "Procedure", "Procedure Division (CopyBook)", state.currentDivision, "", false);
                    state.restorePrevState = true;
                    state.currentSection = undefined;
                    state.currentDivision = fakeDivision;
                    state.procedureDivision = fakeDivision;
                    state.pickFields = false;
                    state.inProcedureDivision = true;
                    sourceLooksLikeCOBOL = true;
                    fakeDivision.ignoreInOutlineView = true;
                    this.sourceIsCopybook = true;
                    state.endsWithDot = true;
                    state.prevEndsWithDot = true;
                }
                else if ((preParseState.workingStorageRelatedTokens !== 0 && preParseState.numberTokensInHeader !== 0)) {
                    const fakeDivision = this.newCOBOLToken(COBOLTokenStyle.Division, 0, "Data Division", 0, "Data", "Data Division (CopyBook)", state.currentDivision, "", false);

                    state.restorePrevState = true;

                    state.currentSection = undefined;
                    state.currentParagraph = undefined;

                    state.currentDivision = fakeDivision;
                    state.pickFields = true;
                    state.inProcedureDivision = false;
                    sourceLooksLikeCOBOL = true;
                    this.ImplicitProgramId = "";
                    fakeDivision.ignoreInOutlineView = true;
                    this.sourceIsCopybook = true;
                }
            }

            //any divs or left early?
            if (preParseState.divisionsInToken !== 0 || preParseState.leaveEarly) {
                sourceLooksLikeCOBOL = true;
            }

            // could it be COBOL (just by the comment area?)
            if (!sourceLooksLikeCOBOL && sourceHandler.getCommentCount() > 0) {
                sourceLooksLikeCOBOL = true;
            }

            /* if the source has an extension, then continue on.. */
            if (hasCOBOLExtension) {
                sourceLooksLikeCOBOL = true;
            }

            /* leave early */
            if (sourceLooksLikeCOBOL === false) {
                if (filename.length > 0) {
                    const linesinSource = sourceHandler.getLineCount();
                    if (linesinSource > maxLines) {
                        this.externalFeatures.logMessage(` Warning - Unable to determine if ${filename} is COBOL after scanning ${maxLines} lines (configurable via ${ExtensionDefaults.defaultEditorConfig}.pre_scan_line_limit setting)`);
                    } else {
                        // if the number of lines is low, don't comment on it
                        if (linesinSource > 5) {
                            this.externalFeatures.logMessage(` Unable to determine if ${filename} is COBOL and how it is used`);
                        }
                    }
                }
            }

            // drop out early
            if (!sourceLooksLikeCOBOL) {
                return;
            }

        } else {
            sourceLooksLikeCOBOL = true;
        }
        sourceHandler.resetCommentCount();

        this.sourceFormat = SourceFormat.get(sourceHandler, configHandler);
        switch (this.sourceFormat) {
            case ESourceFormat.free: sourceHandler.setDumpAreaBOnwards(false);
                break;
            case ESourceFormat.variable: sourceHandler.setDumpAreaBOnwards(false);
                break;
            case ESourceFormat.fixed:
                sourceHandler.setDumpAreaA(true);
                sourceHandler.setDumpAreaBOnwards(true);
                break;
            case ESourceFormat.terminal:
                sourceHandler.setSourceFormat(this.sourceFormat);
                break;
        }

        prevToken = StreamTokens.Blank;
        sourceHandler.resetCommentCount();

        let sourceTimeout = externalFeatures.getSourceTimeout(this.configHandler);
        if (externalFeatures)
            for (let l = 0; l < sourceHandler.getLineCount(); l++) {
                try {
                    state.currentLineIsComment = false;
                    const processedLine: string | undefined = sourceHandler.getLine(l, false);

                    // eof
                    if (processedLine === undefined) {
                        break;
                    }

                    // don't process line
                    if (processedLine.length === 0 && state.currentLineIsComment) {
                        continue;
                    }

                    const line = processedLine.trimEnd();

                    // don't parse a empty line
                    if (line.length > 0) {
                        const prevTokenToParse = prevToken.endsWithDot === false ? prevToken : StreamTokens.Blank;
                        prevToken = this.parseLineByLine(l, prevTokenToParse, line);
                    }

                    // if we are not debugging..
                    if (externalFeatures.isDebuggerActive() === false) {
                        // check for timeout every 1000 lines
                        if (l % 1000 !== 0) {
                            const elapsedTime = externalFeatures.performance_now() - this.sourceReferences.startTime;
                            if (elapsedTime > sourceTimeout) {
                                this.externalFeatures.logMessage(`Aborted scanning ${this.filename} after ${elapsedTime}`);
                                this.clearScanData();
                                return;
                            }
                        }
                    }

                    // only do this for "COBOL" language
                    if (this.usePortationSourceScanner) {
                        const portResult = this.sourcePorter.isDirectiveChangeRequired(this.filename, l, line);
                        if (portResult !== undefined) {
                            this.portWarnings.push(portResult);
                        }
                    }
                }
                catch (e) {
                    this.externalFeatures.logException("COBOLScannner - Parse error", e as Error);
                }
            }

        if (this.sourceReferences.topLevel) {
            const lastLineCount = sourceHandler.getLineCount();
            const lastLineLengthU = sourceHandler.getLine(lastLineCount, true);
            const lastLineLength = lastLineLengthU === undefined ? 0 : lastLineLengthU.length;
            if (state.programs.length !== 0) {
                for (let cp = 0; cp < state.programs.length; cp++) {
                    const currentProgram = state.programs.pop();
                    if (currentProgram !== undefined) {
                        currentProgram.rangeEndLine = sourceHandler.getLineCount();
                        currentProgram.rangeEndColumn = lastLineLength;
                    }
                }
            }

            if (state.currentDivision !== undefined) {
                state.currentDivision.rangeEndLine = sourceHandler.getLineCount();
                state.currentDivision.rangeEndColumn = lastLineLength;
            }

            if (state.currentSection !== undefined) {
                state.currentSection.rangeEndLine = sourceHandler.getLineCount();
                state.currentSection.rangeEndColumn = lastLineLength;
            }

            if (state.currentParagraph !== undefined) {
                state.currentParagraph.rangeEndLine = sourceHandler.getLineCount();
                state.currentParagraph.rangeEndColumn = lastLineLength;
            }

            if (this.ImplicitProgramId.length !== 0) {
                const ctoken = this.newCOBOLToken(COBOLTokenStyle.ImplicitProgramId, 0, "", 0, this.ImplicitProgramId, this.ImplicitProgramId, undefined, "", false);
                ctoken.rangeEndLine = sourceHandler.getLineCount();
                ctoken.rangeStartLine = 0;
                ctoken.rangeEndColumn = lastLineLength;
                ctoken.ignoreInOutlineView = true;
                state.currentProgramTarget.token = ctoken;
                this.tokensInOrder.pop();

                this.callTargets.set(this.ImplicitProgramId, state.currentProgramTarget);
            }

            for (const [sql_declare_name, sql_declare] of this.execSQLDeclare) {
                for (const refExecToken of this.execTokensInOrder) {
                    const fileid = this.sourceReferences.getSourceFieldId(refExecToken.filename);
                    const text = refExecToken.sourceHandler.getText(refExecToken.rangeStartLine, refExecToken.rangeStartColumn, refExecToken.rangeEndLine, refExecToken.rangeEndColumn);
                    this.parseSQLDeclareForReferences(fileid, sql_declare_name, refExecToken, text, sql_declare);
                }
            }

            // setup references for unknown forward reference-items
            const unknown = [];
            for (const [strRef, sourceRefs] of this.sourceReferences.unknownReferences) {
                const possibleTokens = this.constantsOrVariables.get(strRef);
                if (possibleTokens !== undefined) {
                    let ttype: COBOLTokenStyle = COBOLTokenStyle.Variable;
                    let addReference = true;
                    for (const token of possibleTokens) {
                        if (this.configHandler.enable_text_replacement == true) {
                            addReference = true;
                        }
                        else if (token.ignoreInOutlineView == false || token.token.isFromScanCommentsForReferences) {
                            ttype = (token.tokenType === COBOLTokenStyle.Unknown) ? ttype : token.tokenType;
                        } else {
                            addReference = false;
                        }
                    }
                    if (addReference) {
                        this.transferReference(strRef, sourceRefs, this.sourceReferences.constantsOrVariablesReferences, ttype);
                    }
                }
                else if (this.isVisibleSection(strRef)) {
                    this.transferReference(strRef, sourceRefs, this.sourceReferences.targetReferences, COBOLTokenStyle.Section);
                } else if (this.isVisibleParagraph(strRef)) {
                    this.transferReference(strRef, sourceRefs, this.sourceReferences.targetReferences, COBOLTokenStyle.Paragraph);
                } else {
                    unknown.push(strRef);
                }
            }

            if (this.sourceReferences.state.skipToEndLsIgnore) {
                if (this.lastCOBOLLS !== undefined) {
                    const diagMessage = `Missing ${configHandler.scan_comment_end_ls_ignore}`;
                    this.generalWarnings.push(new COBOLFileSymbol(this.filename, this.lastCOBOLLS.startLine, diagMessage));
                    while (this.sourceReferences.ignoreLSRanges.pop() !== undefined);
                }
            }
            this.sourceReferences.unknownReferences.clear();
            this.eventHandler.finish();
        }
    }


    private clearScanData() {
        this.tokensInOrder = [];
        this.copyBooksUsed.clear();
        this.sections.clear();
        this.paragraphs.clear();
        this.constantsOrVariables.clear();
        this.callTargets.clear();
        this.functionTargets.clear();
        this.classes.clear();
        this.methods.clear();
        this.diagMissingFileWarnings.clear();
        this.cache4PerformTargets = undefined;
        this.cache4ConstantsOrVars = undefined;
        this.scanAborted = true;
        this.sourceReferences.reset(this.configHandler);
    }

    private isVisibleSection(sectionName: string): boolean {
        if (!sectionName) {
            return false;
        }
        const foundSectionToken = this.sections.get(sectionName) ?? this.sections.get(sectionName.toLowerCase());
        if (!foundSectionToken) {
            return false;
        }
        return foundSectionToken.isFromScanCommentsForReferences || !foundSectionToken.ignoreInOutlineView;
    }

    private isVisibleParagraph(paragraphName: string): boolean {
        if (!paragraphName) {
            return false;
        }
        const foundParagraph = this.paragraphs.get(paragraphName) ?? this.paragraphs.get(paragraphName.toLowerCase());
        if (!foundParagraph) {
            return false;
        }
        return foundParagraph.isFromScanCommentsForReferences || !foundParagraph.ignoreInOutlineView;
    }

    private transferReference(symbol: string, symbolRefs: SourceReference_Via_Length[], transferReferenceMap: Map<string, SourceReference_Via_Length[]>, tokenStyle: COBOLTokenStyle): void {
        if (symbol.length === 0) {
            return;
        }

        if (this.isValidKeyword(symbol)) {
            return;
        }

        const refList = transferReferenceMap.get(symbol);
        if (refList !== undefined) {
            for (const sourceRef of symbolRefs) {
                sourceRef.tokenStyle = tokenStyle;
                refList.push(sourceRef);
            }
        } else {
            for (const sourceRef of symbolRefs) {
                sourceRef.tokenStyle = tokenStyle;
            }
            transferReferenceMap.set(symbol, symbolRefs);
        }
    }

    private tokensInOrderPush(token: COBOLToken, sendEvent: boolean): void {
        this.tokensInOrder.push(token);
        if (sendEvent) {
            this.eventHandler.processToken(token);
        }
    }

    /**
     * Creates and registers a new COBOLToken, updating relevant parse state.
     * Handles token range updates for divisions, sections, paragraphs, and implicit/ignored tokens.
     */
    private newCOBOLToken(
        tokenType: COBOLTokenStyle,
        startLine: number,
        line: string,
        currentCol: number,
        token: string,
        description: string,
        parentToken: COBOLToken | undefined,
        extraInformation1: string,
        isImplicitToken: boolean
    ): COBOLToken {
        const state = this.sourceReferences.state;
        let startColumn = currentCol;

        // For most token types, try to find the token's column in the line
        if (tokenType !== COBOLTokenStyle.CopyBook && tokenType !== COBOLTokenStyle.CopyBookInOrOf) {
            startColumn = line.indexOf(token, currentCol);
            if (startColumn === -1) {
                startColumn = line.indexOf(token);
                if (startColumn === -1) {
                    startColumn = Math.min(currentCol, line.length);
                }
            }
        }

        const ctoken = new COBOLToken(
            this.sourceHandler,
            this.sourceHandler.getUriAsString(),
            this.filename,
            tokenType,
            startLine,
            startColumn,
            token,
            description,
            parentToken,
            state.inProcedureDivision,
            extraInformation1,
            this.isFromScanCommentsForReferences,
            isImplicitToken
        );
        ctoken.ignoreInOutlineView = state.ignoreInOutlineView;
        ctoken.inSection = state.currentSection;

        // Always push implicit or ignored tokens
        if (ctoken.ignoreInOutlineView || tokenType === COBOLTokenStyle.ImplicitProgramId) {
            this.tokensInOrderPush(ctoken, true);
            return ctoken;
        }

        // Helper to update rangeEnd for a token
        const updateRangeEnd = (tok: COBOLToken | undefined, line: number, col: number) => {
            if (tok) {
                tok.rangeEndLine = line;
                if (col !== 0) {
                    tok.rangeEndColumn = col - 1;
                }
            }
        };

        switch (tokenType) {
            case COBOLTokenStyle.Division:
                updateRangeEnd(state.currentDivision, startLine, ctoken.rangeStartColumn);
                updateRangeEnd(state.currentSection, startLine, ctoken.rangeStartColumn);
                state.currentSection = undefined;
                state.currentParagraph = undefined;
                this.tokensInOrderPush(ctoken, true);
                return ctoken;

            case COBOLTokenStyle.Section:
                updateRangeEnd(state.currentSection, startLine, ctoken.rangeStartColumn);
                state.currentParagraph = undefined;
                this.tokensInOrderPush(ctoken, state.inProcedureDivision);
                return ctoken;

            case COBOLTokenStyle.Paragraph:
                updateRangeEnd(state.currentSection, startLine, ctoken.rangeStartColumn);
                state.currentParagraph = ctoken;
                updateRangeEnd(state.currentDivision, startLine, ctoken.rangeStartColumn);
                this.tokensInOrderPush(ctoken, true);
                return ctoken;

            case COBOLTokenStyle.IgnoreLS:
                this.tokensInOrderPush(ctoken, false);
                return ctoken;

            default:
                // Update previous paragraph's range if present
                updateRangeEnd(state.currentParagraph, startLine, ctoken.rangeStartColumn);
                this.tokensInOrderPush(ctoken, true);
                return ctoken;
        }
    }


    private static readonly literalRegex = /^#?[a-zA-Z0-9][a-zA-Z0-9_-]*$/;

    public static isValidLiteral(id: string): boolean {
        if (!id || typeof id !== "string") {
            return false;
        }
        return COBOLSourceScanner.literalRegex.test(id);
    }

    // Use a non-global regex for paragraph name validation
    private static readonly paragraphRegex = /^[a-zA-Z0-9][a-zA-Z0-9-_]*$/;

    private isParagraph(id: string): boolean {
        if (typeof id !== "string" || id.length === 0) {
            return false;
        }
        if (!COBOLSourceScanner.paragraphRegex.test(id)) {
            return false;
        }
        return !this.constantsOrVariables.has(id.toLowerCase());
    }


    private isValidParagraphName(id: string): boolean {
        if (!id || typeof id !== "string") {
            return false;
        }
        if (!COBOLSourceScanner.paragraphRegex.test(id)) {
            return false;
        }
        return !this.constantsOrVariables.has(id.toLowerCase());
    }

    private isValidKeyword(keyword: string): boolean {
        return this.COBOLKeywordDictionary.has(keyword);
    }

    private isValidProcedureKeyword(keyword: string): boolean {
        return cobolProcedureKeywordDictionary.has(keyword);
    }

    private isValidStorageKeyword(keyword: string): boolean {
        return cobolStorageKeywordDictionary.has(keyword);
    }

    private isNumber(value: string): boolean {
        try {
            if (value.toString().length === 0) {
                return false;
            }
            return !isNaN(Number(value.toString()));
        }
        catch (e) {
            this.externalFeatures.logException("isNumber(" + value + ")", e as Error);
            return false;
        }
    }

    private containsIndex(literal: string): boolean {
        for (let pos = 0; pos < literal.length; pos++) {
            if (literal[pos] === "(" || literal[pos] === ")" &&
                literal[pos] === "[" || literal[pos] === "]") {
                return true;
            }
        }
        return false;
    }

    private trimVariableToMap(literal: string): Map<number, string> {
        const varMap = new Map<number, string>();
        let variable = "";
        let startPos = 0;
        
        for (let pos = 0; pos < literal.length; pos++) {
            const char = literal[pos];
            
            switch (char) {
                case "(":
                case "[":
                    if (variable.length !== 0) {
                        varMap.set(startPos, variable);
                        startPos = pos + 1;
                        variable = "";
                    }
                    break;
                case ")":
                case "]":
                    if (variable.length !== 0) {
                        varMap.set(startPos, variable);
                        startPos = pos + 1;
                        variable = "";
                    }
                    break;
                default:
                    variable += char;
            }
        }

        if (variable.length !== 0) {
            varMap.set(startPos, variable);
        }
        
        return varMap;
    }

    public static trimLiteral(literal: string, trimQuotes: boolean): string {
        // Handle null, undefined, or non-string inputs gracefully
        if (!literal || typeof literal !== "string") {
            return "";
        }

        let result = literal.trim();

        // Remove leading '(' and trailing ')', trailing dot, and whitespace
        // Using a single regex to handle all these operations efficiently
        result = result.replace(/^\(+|\)+$/g, "").replace(/\.$/, "").trim();

        // Remove surrounding single or double quotes if trimQuotes is true
        if (trimQuotes && result.length >= 2) {
            // Match quotes at start and end of string
            const quoteMatch = result.match(/^(['"])(.*)\1$/);
            if (quoteMatch) {
                result = quoteMatch[2];
            }
        }

        return result;
    }

    private isQuotedLiteral(literal: string): boolean {
        if (!literal) {
            return false;
        }

        // Trim the literal to handle whitespace
        const trimmedLiteral = literal.trim();

        // Check if it's empty after trimming or too short to be quoted
        if (trimmedLiteral.length < 2) {
            return false;
        }

        // Remove trailing period if present
        const noPeriodLiteral = trimmedLiteral.endsWith('.') ? trimmedLiteral.slice(0, -1) : trimmedLiteral;

        // Check if it's quoted with either single or double quotes using regex
        // This pattern ensures:
        // ^["']     - starts with a quote (single or double)
        // [^"']*    - any number of non-quote characters
        // ["']$     - ends with a matching quote
        return /^["'][^"']*["']$/.test(noPeriodLiteral);
    }

    private relaxedParseLineByLine(prevToken: StreamTokens, line: string, lineNumber: number, state: PreParseState): StreamTokens {
        const token = new StreamTokens(line, lineNumber, prevToken);
        let tokenCountPerLine = 0;
        
        do {
            try {
                const endsWithDot = token.endsWithDot;
                let current: string = token.currentToken;
                let currentLower: string = token.currentTokenLower;
                tokenCountPerLine++;

                // Remove trailing dot if present
                if (endsWithDot) {
                    current = current.substring(0, current.length - 1);
                    currentLower = currentLower.substring(0, currentLower.length - 1);
                }

                // Handle first token on line - check for numeric values
                if (tokenCountPerLine === 1) {
                    const tokenAsNumber = Number.parseInt(current, 10);
                    if (!isNaN(tokenAsNumber)) {
                        state.numberTokensInHeader++;
                        continue;
                    }
                }

                // Process tokens based on their content
                switch (currentLower) {
                    case "section":
                        this.handleSectionToken(token, state);
                        break;
                    case "program-id":
                        state.divisionsInToken++;
                        state.leaveEarly = true;
                        break;
                    case "division":
                        this.handleDivisionToken(token, state);
                        break;
                    default:
                        this.handleDefaultTokens(currentLower, state);
                        break;
                }
                
                // Skip empty tokens
                if (current.length === 0) {
                    continue;
                }
            }
            catch (e) {
                this.externalFeatures.logException("COBOLScannner relaxedParseLineByLine line error: ", e as Error);
            }
        }
        while (token.moveToNextToken() === false);

        return token;
    }

    private handleSectionToken(token: StreamTokens, state: PreParseState): void {
        if (token.prevToken.length !== 0) {
            switch (token.prevTokenLower) {
                case "working-storage":
                case "file":
                case "linkage":
                case "screen":
                case "input-output":
                    state.sectionsInToken++;
                    state.leaveEarly = true;
                    break;
                default:
                    if (!this.isValidProcedureKeyword(token.prevTokenLower)) {
                        state.procedureDivisionRelatedTokens++;
                    }
                    break;
            }
        }
    }

    private handleDivisionToken(token: StreamTokens, state: PreParseState): void {
        switch (token.prevTokenLower) {
            case "identification":
                state.divisionsInToken++;
                state.leaveEarly = true;
                break;
            case "procedure":
                state.procedureDivisionRelatedTokens++;
                state.leaveEarly = true;
                break;
        }
    }

    private handleDefaultTokens(currentLower: string, state: PreParseState): void {
        if (this.isValidProcedureKeyword(currentLower)) {
            state.procedureDivisionRelatedTokens++;
        }

        if (this.isValidStorageKeyword(currentLower)) {
            state.workingStorageRelatedTokens++;
        }
    }

    private addVariableReference(
        referencesMap: Map<string, SourceReference_Via_Length[]>, 
        lowerCaseVariable: string, 
        line: number, 
        column: number, 
        tokenStyle: COBOLTokenStyle
    ): boolean {
        // Early validation checks
        if (lowerCaseVariable.length === 0) {
            return false;
        }

        if (this.isValidKeyword(lowerCaseVariable)) {
            return false;
        }

        if (!COBOLSourceScanner.isValidLiteral(lowerCaseVariable)) {
            return false;
        }

        // Check if reference already exists to avoid duplicates
        const existingReferences = referencesMap.get(lowerCaseVariable);
        if (existingReferences !== undefined) {
            // Use find() for better readability and performance than manual loop
            const isDuplicate = existingReferences.some(ref => 
                ref.line === line && 
                ref.column === column && 
                ref.length === lowerCaseVariable.length
            );
            
            if (!isDuplicate) {
                existingReferences.push(new SourceReference_Via_Length(
                    this.sourceFileId, 
                    line, 
                    column, 
                    lowerCaseVariable.length, 
                    tokenStyle, 
                    this.isFromScanCommentsForReferences, 
                    lowerCaseVariable, 
                    ""
                ));
            }
            return true;
        }

        // Create new reference entry
        const newReferences: SourceReference_Via_Length[] = [
            new SourceReference_Via_Length(
                this.sourceFileId, 
                line, 
                column, 
                lowerCaseVariable.length, 
                tokenStyle, 
                this.isFromScanCommentsForReferences, 
                lowerCaseVariable, 
                ""
            )
        ];
        
        referencesMap.set(lowerCaseVariable, newReferences);
        return true;
    }

    private addTargetReference(
        referencesMap: Map<string, SourceReference_Via_Length[]>, 
        _targetReference: string, 
        line: number, 
        column: number, 
        tokenStyle: COBOLTokenStyle, 
        reason: string
    ): boolean {
        // Early exit conditions
        if (_targetReference.length === 0) {
            return false;
        }

        const targetReference = _targetReference.toLowerCase();
        
        // Skip keywords and invalid literals
        if (this.isValidKeyword(targetReference) || !COBOLSourceScanner.isValidLiteral(targetReference)) {
            return false;
        }

        // Create source reference
        const srl = new SourceReference_Via_Length(
            this.sourceFileId, 
            line, 
            column, 
            _targetReference.length, 
            tokenStyle, 
            this.isFromScanCommentsForReferences, 
            _targetReference, 
            reason
        );

        // Get or create reference list for target
        let lowerCaseTargetRefs = referencesMap.get(targetReference);
        if (lowerCaseTargetRefs === undefined) {
            lowerCaseTargetRefs = [];
            referencesMap.set(targetReference, lowerCaseTargetRefs);
        }

        const state = this.sourceReferences.state;
        const inSectionOrParaToken = state.currentParagraph ?? state.currentSection;
        
        // Handle procedure division references with duplicate checking
        if (inSectionOrParaToken?.inProcedureDivision) {
            // Check for duplicates before adding
            const isDuplicate = lowerCaseTargetRefs.some(ref => 
                ref.line === line && 
                ref.column === column && 
                ref.length === _targetReference.length
            );
            
            if (!isDuplicate) {
                lowerCaseTargetRefs.push(srl);
                
                // Track cross-references for current section/paragraph
                const sectionOutRefs = state.currentSectionOutRefs;
                const tokenName = inSectionOrParaToken.tokenNameLower;
                
                let csr = sectionOutRefs.get(tokenName);
                if (csr === undefined) {
                    csr = [];
                    sectionOutRefs.set(tokenName, csr);
                }
                csr.push(srl);
            }
            
            return true;
        }

        // Add reference for non-procedure division context
        lowerCaseTargetRefs.push(srl);
        return true;
    }

    private addVariableOrConstant(lowerCaseVariable: string, cobolToken: COBOLToken) {
        if (lowerCaseVariable.length === 0) {
            return;
        }

        if (this.isValidKeyword(lowerCaseVariable)) {
            return;
        }

        if (this.addVariableReference(this.sourceReferences.constantsOrVariablesReferences, lowerCaseVariable, cobolToken.startLine, cobolToken.startColumn, cobolToken.tokenType) == false) {
            return;
        }

        const constantsOrVariablesToken = this.constantsOrVariables.get(lowerCaseVariable);
        if (constantsOrVariablesToken !== undefined) {
            constantsOrVariablesToken.push(new COBOLVariable(cobolToken));
            return;
        }

        const tokens: COBOLVariable[] = [];
        tokens.push(new COBOLVariable(cobolToken));
        this.constantsOrVariables.set(lowerCaseVariable, tokens);
    }

    private cleanupReplaceToken(token: string): [string, boolean] {
        if (token.endsWith(",")) {
            token = token.substring(0, token.length - 1);
        }

        if (token.startsWith("==") && token.endsWith("==")) {
            return [token.substring(2, token.length - 2), true];
        }

        return [token, false];
    }

    private parseLineByLine(lineNumber: number, prevToken: StreamTokens, line: string): StreamTokens {
        const token = new StreamTokens(line, lineNumber, prevToken);

        const state = this.sourceReferences.state;
        let stream = this.processToken(lineNumber, token, line, state.replaceMap.size !== 0);

        // update current group to always include the end of current line
        if (state.current01Group !== undefined && line.length !== 0 && state.inCopy === false) {
            state.current01Group.rangeEndLine = lineNumber;
            state.current01Group.rangeEndColumn = line.length;
        }

        // update current variable to always include the end of current line
        if (state.currentLevel?.tokenType === COBOLTokenStyle.Variable) {
            state.currentLevel.rangeEndLine = lineNumber;
            state.currentLevel.rangeEndColumn = line.length;
        }

        return stream;
    }

    private processToken(lineNumber: number, token: StreamTokens, line: string, replaceOn: boolean): StreamTokens {
        const state: ParseState = this.sourceReferences.state;

        do {
            try {
                // console.log(`DEBUG: ${line}`);
                let _tcurrent: string = token.currentToken;
                // continue now
                if (_tcurrent.length === 0) {
                    continue;
                }

                // if skip to end lsp
                if (state.skipToEndLsIgnore) {
                    // add skipped token into sourceref range to colour as a comment
                    if (this.configHandler.enable_semantic_token_provider) {
                        const sr = new SourceReference(this.sourceFileId, token.currentLineNumber, token.currentCol, token.currentLineNumber, token.currentCol + token.currentToken.length, COBOLTokenStyle.IgnoreLS);
                        this.sourceReferences.ignoreLSRanges.push(sr);
                    }

                    continue;
                }

                // skip one token
                if (state.skipNextToken) {
                    state.skipNextToken = false;

                    // time to leave?  
                    //  [handles "pic x.", where x. is the token to be skipped]
                    if (state.skipToDot && _tcurrent.endsWith(".")) {
                        state.skipToDot = false;
                        // drop through, so all other flags can be reset or contune past
                    } else {
                        continue;
                    }
                }

                // fakeup a replace algorithmf
                if (replaceOn) {
                    let rightLine = line.substring(token.currentCol);
                    const rightLineOrg = line.substring(token.currentCol);
                    for (const [k, r] of state.replaceMap) {
                        rightLine = rightLine.replace(r.pattern, k);
                    }

                    if (rightLine !== rightLineOrg) {
                        try {
                            const leftLine = line.substring(0, token.currentCol);

                            this.sourceHandler.setUpdatedLine(lineNumber, leftLine + rightLine);
                            const lastTokenId = this.tokensInOrder.length;
                            const newToken = new StreamTokens(rightLine as string, lineNumber, new StreamTokens(token.prevToken, token.prevTokenLineNumber, undefined));
                            const retToken = this.processToken(lineNumber, newToken, rightLine, false);

                            // ensure any new token match the original soure
                            if (lastTokenId !== this.tokensInOrder.length) {
                                for (let ltid = lastTokenId; ltid < this.tokensInOrder.length; ltid++) {
                                    const addedToken = this.tokensInOrder[ltid];
                                    addedToken.rangeStartColumn = token.currentCol;
                                    addedToken.rangeEndColumn = token.currentCol + _tcurrent.length;
                                }
                            }
                            return retToken;
                        } catch (e) {
                            this.externalFeatures.logException("replace", e as Error);
                        }
                    }
                }


                // HACK for "set x to entry"
                if (token.prevTokenLower === "to" && token.currentTokenLower === "entry") {
                    token.moveToNextToken();
                    continue;
                }

                if (_tcurrent.endsWith(",")) {
                    _tcurrent = _tcurrent.substring(0, _tcurrent.length - 1);
                    state.prevEndsWithDot = state.endsWithDot;
                    state.endsWithDot = false;
                } else if (token.endsWithDot) {
                    _tcurrent = _tcurrent.substring(0, _tcurrent.length - 1);
                    state.prevEndsWithDot = state.endsWithDot;
                    state.endsWithDot = true;
                } else {
                    state.prevEndsWithDot = state.endsWithDot;
                    state.endsWithDot = false;
                }

                const currentCol = token.currentCol;
                const current: string = _tcurrent;
                const currentLower: string = _tcurrent.toLowerCase();

                // if pickUpUsing
                if (state.pickUpUsing) {
                    if (state.endsWithDot) {
                        state.pickUpUsing = false;
                    }

                    if (token.prevTokenLower === "type") {
                        continue;
                    }

                    switch (currentLower) {
                        case "signed":
                            break;
                        case "unsigned":
                            break;
                        case "as":
                            break;
                        case "type":
                            break;
                        case "using":
                            state.using = UsingState.BY_REF;
                            break;
                        case "by":
                            break;
                        case "reference":
                            state.using = UsingState.BY_REF;
                            break;
                        case "value":
                            state.using = UsingState.BY_VALUE;
                            break;
                        case "output":
                            state.using = UsingState.BY_OUTPUT;
                            break;
                        case "returning":
                            state.using = UsingState.RETURNING;
                            break;
                        default:
                            // of are after a returning statement with a "." then parsing must end now
                            if (state.using === UsingState.UNKNOWN) {
                                state.pickUpUsing = false;
                                break;
                            }

                            if (this.sourceReferences !== undefined) {
                                if (currentLower === "any") {
                                    state.parameters.push(new COBOLParameter(state.using, current));
                                } else if (currentLower.length > 0 && this.isValidKeyword(currentLower) === false && this.isNumber(currentLower) === false) {
                                    // no forward validation can be done, as this is a one pass scanner
                                    if (this.addVariableReference(this.sourceReferences.unknownReferences, currentLower, lineNumber, token.currentCol, COBOLTokenStyle.Variable)) {
                                        state.parameters.push(new COBOLParameter(state.using, current));
                                    }
                                }
                            }

                            if (state.using === UsingState.RETURNING) {
                                state.using = UsingState.UNKNOWN;
                            }

                            // if entry or procedure division does not have "." then parsing must end now
                            if (currentLower !== "any" && this.isValidKeyword(currentLower)) {
                                state.currentProgramTarget.callParameters = state.parameters;
                                state.using = UsingState.UNKNOWN;
                                state.pickUpUsing = false
                                if (this.configHandler.linter_ignore_malformed_using === false) {
                                    const diagMessage = `Unexpected keyword '${current}' when scanning USING parameters`;
                                    let nearestLine = state.currentProgramTarget.token !== undefined ? state.currentProgramTarget.token.startLine : lineNumber;
                                    this.generalWarnings.push(new COBOLFileSymbol(this.filename, nearestLine, diagMessage));
                                }
                                break;
                            }


                        // logMessage(`INFO: using parameter : ${tcurrent}`);
                    }
                    if (state.endsWithDot) {
                        state.currentProgramTarget.callParameters = state.parameters;
                        state.pickUpUsing = false;
                    }

                    continue;
                }


                // if skipToDot and not the end of the statement.. swallow
                if (state.skipToDot) {
                    const trimTokenLower = COBOLSourceScanner.trimLiteral(currentLower, false);
                    const tokenIsKeyword = this.isValidKeyword(trimTokenLower);

                    if (state.addReferencesDuringSkipToTag) {
                        if (this.sourceReferences !== undefined) {
                            if (COBOLSourceScanner.isValidLiteral(trimTokenLower) && !this.isNumber(trimTokenLower) && tokenIsKeyword === false) {
                                // no forward validation can be done, as this is a one pass scanner
                                if (token.prevTokenLower !== "pic" && token.prevTokenLower !== "picture") {
                                    this.addVariableReference(this.sourceReferences.unknownReferences, trimTokenLower, lineNumber, token.currentCol, COBOLTokenStyle.Unknown);
                                }
                            }
                        }
                    }

                    if (trimTokenLower === "value") {
                        state.inValueClause = true;
                        state.addVariableDuringStipToTag = false;
                        continue;
                    }

                    // turn off at keyword
                    if (trimTokenLower === "pic" || trimTokenLower === "picture") {
                        state.skipNextToken = true;
                        state.addVariableDuringStipToTag = false;
                        continue;
                    }

                    // if we are in a to.. or indexed
                    if (token.prevTokenLower === "to") {
                        state.addVariableDuringStipToTag = false;
                        continue;
                    }

                    if (token.prevTokenLower === "indexed" && token.currentTokenLower === "by") {
                        state.addVariableDuringStipToTag = true;
                    }

                    if (token.prevTokenLower === "depending" && token.currentTokenLower === "on") {
                        state.addVariableDuringStipToTag = false;
                    }

                    if (state.addVariableDuringStipToTag && tokenIsKeyword === false && state.inValueClause === false) {
                        if (COBOLSourceScanner.isValidLiteral(trimTokenLower) && !this.isNumber(trimTokenLower)) {
                            const trimToken = COBOLSourceScanner.trimLiteral(current, false);
                            const variableToken = this.newCOBOLToken(COBOLTokenStyle.Variable, lineNumber, line, currentCol, trimToken, trimToken, state.currentDivision, token.prevToken, false);
                            this.addVariableOrConstant(trimTokenLower, variableToken);
                        }
                    }

                    if (state.inReplace) {
                        switch (currentLower) {
                            case "by":
                                state.captureReplaceLeft = false;
                                break;
                            case "off":
                                state.skipToDot = false;
                                state.inReplace = false;
                                state.replaceMap.clear();
                                break;
                            default:
                                if (state.captureReplaceLeft) {
                                    if (state.replaceLeft.length !== 0) {
                                        state.replaceLeft += " ";
                                    }
                                    state.replaceLeft += current;
                                } else {
                                    if (state.replaceRight.length !== 0) {
                                        state.replaceRight += " ";
                                    }
                                    state.replaceRight += current;
                                }
                                if (!state.captureReplaceLeft && current.endsWith("==")) {
                                    let cleanedUpRight = this.cleanupReplaceToken("" + state.replaceRight);
                                    let cleanedUpLeft = this.cleanupReplaceToken("" + state.replaceLeft);
                                    let rs = new replaceState();
                                    if (cleanedUpLeft[1] || cleanedUpRight[1])
                                        rs.isPseudoTextDelimiter = true;
                                    else
                                        rs.isPseudoTextDelimiter = false;

                                    state.replaceMap.set(cleanedUpRight[0], new ReplaceToken(cleanedUpLeft[0], rs));
                                    state.replaceLeft = state.replaceRight = "";
                                    state.captureReplaceLeft = true;
                                }
                                break;
                        }

                    }

                    if (state.inCopy) {
                        const cbState = state.copybook_state;
                        switch (currentLower) {
                            case "":
                                break;
                            case "suppress":
                            case "resource":
                            case "indexed":
                                break;
                            case "leading":
                                cbState.isLeading = true;
                                break;
                            case "trailing":
                                cbState.isTrailing = true;
                                break;
                            case "of": cbState.isOf = true;
                                break;
                            case "in": cbState.isIn = true;
                                break;
                            case "replacing":
                                cbState.isReplacingBy = false;
                                cbState.isReplacing = true;
                                cbState.isLeading = false;
                                cbState.isTrailing = false;
                                break;
                            case "by":
                                cbState.isReplacingBy = true;
                                break;
                            default: {
                                if (cbState.isIn && cbState.literal2.length === 0) {
                                    cbState.literal2 = current;
                                    break;
                                }
                                if (cbState.isOf && cbState.library_name.length === 0) {
                                    cbState.library_name = current;
                                    break;
                                }
                                if (cbState.isReplacing && cbState.replaceLeft.length === 0) {
                                    cbState.replaceLeft = current;
                                    break;
                                }
                                if (cbState.isReplacingBy) {
                                    if (this.configHandler.enable_text_replacement) {
                                        let cleanedUpRight = this.cleanupReplaceToken("" + current);
                                        let cleanedUpLeft = this.cleanupReplaceToken("" + cbState.replaceLeft);
                                        let rs = new replaceState();
                                        if (cleanedUpLeft[1] || cleanedUpRight[1])
                                            rs.isPseudoTextDelimiter = true;
                                        else
                                            rs.isPseudoTextDelimiter = false;

                                        cbState.copyReplaceMap.set(cleanedUpRight[0], new ReplaceToken(cleanedUpLeft[0], rs));
                                    }
                                    cbState.isReplacingBy = false;
                                    cbState.isReplacing = true;
                                    cbState.isLeading = false;
                                    cbState.isTrailing = false;
                                    cbState.replaceLeft = "";
                                    break;
                                }
                                if (currentLower.length > 0 && !cbState.isOf && !cbState.isIn && !cbState.isReplacing) {
                                    cbState.copyBook = current;
                                    cbState.trimmedCopyBook = COBOLSourceScanner.trimLiteral(current, true);
                                    cbState.startLineNumber = lineNumber;
                                    cbState.startCol = state.inCopyStartColumn; // stored when 'copy' is seen
                                    cbState.line = line;
                                    break;
                                }
                            }
                        }
                    }

                    //reset and process anything if necessary
                    if (state.endsWithDot === true) {
                        state.inReplace = false;
                        state.skipToDot = false;
                        state.inValueClause = false;
                        state.addReferencesDuringSkipToTag = false;

                        if (state.inCopy) {
                            state.copybook_state.endLineNumber = lineNumber;
                            state.copybook_state.endCol = token.currentCol + token.currentToken.length;
                            if (this.processCopyBook(state.copybook_state) === false) {
                                let extra = "";
                                if (state.copybook_state.copybookDepths.length >= this.configHandler.copybook_scan_depth) {
                                    extra = `due to copybook processing depth limit (${this.configHandler.copybook_scan_depth})`;
                                }
                                const trimmedCopyBook = state.copybook_state.trimmedCopyBook;
                                const diagMessage = `Unable to process copybook ${extra}: ${trimmedCopyBook}`;
                                // if (this.configHandler.linter_ignore_missing_copybook) {
                                this.externalFeatures.logMessage(diagMessage);
                                // TODO: make configureable if required
                                // } else {
                                // this.generalWarnings.push(new COBOLFileSymbol(this.filename, state.copybook_state.startLineNumber, diagMessage));
                                // }
                            }
                            state.current01Group = state.copybook_state.saved01Group;
                        }
                        state.inCopy = false;
                    }
                    continue;
                }


                const prevToken = COBOLSourceScanner.trimLiteral(token.prevToken, false);
                const prevTokenLowerUntrimmed = token.prevTokenLower.trim();
                const prevTokenLower = COBOLSourceScanner.trimLiteral(prevTokenLowerUntrimmed, false);
                const prevCurrentCol = token.prevCurrentCol;

                const prevPlusCurrent = token.prevToken + " " + current;

                if (currentLower === "exec") {
                    this.currentExec = token.nextSTokenIndex(1).currentToken;
                    this.currentExecVerb = token.nextSTokenIndex(2).currentToken;
                    state.currentToken = this.newCOBOLToken(COBOLTokenStyle.Exec, lineNumber, line, 0, current, `EXEC ${this.currentExec} ${this.currentExecVerb}`, state.currentDivision, "", false);
                    this.currentExecToken = state.currentToken;
                    token.moveToNextToken();
                    continue;
                }

                // do we have a split line exec
                if (this.currentExecToken !== undefined) {
                    if (this.currentExecVerb.length === 0) {
                        this.currentExecVerb = current;
                        this.currentExecToken.description = `EXEC ${this.currentExec} ${this.currentExecVerb}`;
                    }
                }

                /* finish processing end-exec */
                if (currentLower === "end-exec") {
                    if (this.currentExecToken !== undefined) {
                        if (state.currentToken !== undefined) {
                            this.currentExecToken.rangeEndLine = token.currentLineNumber;
                            this.currentExecToken.rangeEndColumn = token.currentCol + token.currentToken.length;
                            this.execTokensInOrder.push(this.currentExecToken);  // remember token
                        }

                        if (this.configHandler.enable_exec_sql_cursors) {
                            const text = this.currentExecToken.sourceHandler.getText(this.currentExecToken.rangeStartLine, this.currentExecToken.rangeStartColumn, this.currentExecToken.rangeEndLine, this.currentExecToken.rangeEndColumn);
                            this.parseExecStatement(this.currentExec, this.currentExecToken, text);
                        }
                    }
                    this.currentExec = "";
                    this.currentExecVerb = "";
                    state.currentToken = undefined;
                    state.prevEndsWithDot = state.endsWithDot;
                    state.endsWithDot = true;
                    token.endsWithDot = state.endsWithDot;

                    this.currentExecToken = undefined;
                    continue;
                }

                /* skip everything in between exec .. end-exec */
                if (state.currentToken !== undefined && this.currentExec.length !== 0) {
                    if (currentLower === "include" && this.currentExec.toLowerCase() === "sql") {
                        const sqlCopyBook = token.nextSTokenOrBlank().currentToken;
                        const trimmedCopyBook = COBOLSourceScanner.trimLiteral(sqlCopyBook, true);
                        let insertInSection = this.copybookNestedInSection ? state.currentSection : state.currentDivision;
                        if (insertInSection === undefined) {
                            insertInSection = state.currentDivision;
                        }
                        this.tokensInOrder.pop();
                        const copyToken = this.newCOBOLToken(COBOLTokenStyle.CopyBook, lineNumber, line, currentCol, trimmedCopyBook, "EXEC SQL INCLUDE " + sqlCopyBook, insertInSection, "", false);
                        if (this.copyBooksUsed.has(trimmedCopyBook) === false) {
                            const prevState = state.copybook_state;
                            state.copybook_state = new copybookState(state.current01Group);
                            state.copybook_state.trimmedCopyBook = trimmedCopyBook;
                            state.copybook_state.copyBook = sqlCopyBook;
                            state.copybook_state.line = line;
                            state.copybook_state.startLineNumber = lineNumber;
                            state.copybook_state.endLineNumber = lineNumber;
                            state.copybook_state.startCol = currentCol;
                            state.copybook_state.endCol = line.indexOf(sqlCopyBook) + sqlCopyBook.length;
                            state.copybook_state.copybookDepths = prevState.copybookDepths;
                            const fileName = this.externalFeatures.expandLogicalCopyBookToFilenameOrEmpty(trimmedCopyBook, copyToken.extraInformation1, this.sourceHandler, this.configHandler);
                            if (fileName.length === 0) {
                                continue;
                            }
                            state.copybook_state.fileName = fileName;
                            const copybookToken = new COBOLCopybookToken(copyToken, false, state.copybook_state);
                            this.copyBooksUsed.set(trimmedCopyBook, [copybookToken]);
                            const qfile = new FileSourceHandler(this.configHandler, undefined, fileName, this.externalFeatures);
                            const prevIgnoreInOutlineView = state.ignoreInOutlineView;
                            state.ignoreInOutlineView = true;
                            const currentTopLevel = this.sourceReferences.topLevel;
                            this.sourceReferences.topLevel = false;
                            const prevRepMap = state.replaceMap;

                            // eslint-disable-next-line @typescript-eslint/no-unused-vars
                            const isOkay = this.ScanUncachedInlineCopybook(qfile, this, false)
                            state.ignoreInOutlineView = prevIgnoreInOutlineView;
                            state.replaceMap = prevRepMap;
                            this.sourceReferences.topLevel = currentTopLevel;
                            copybookToken.scanComplete = true;
                            state.inCopy = false;
                            if (!isOkay) {
                                const diagMessage = `Unable perform inline sql include ${trimmedCopyBook}`;
                                this.diagMissingFileWarnings.set(diagMessage, new COBOLFileSymbol(this.filename, copyToken.startLine, trimmedCopyBook));
                            }
                        }
                        continue;
                    }

                    // tweak exec to include verb
                    if (this.currentExecVerb.length === 0) {
                        this.currentExecVerb = token.currentToken;
                        state.currentToken.description += " " + this.currentExecVerb;
                        continue;
                    }

                    /* is this a reference to a variable? */
                    const varTokens = this.constantsOrVariables.get(currentLower);
                    if (varTokens !== undefined) {
                        let ctype: COBOLTokenStyle = COBOLTokenStyle.Variable;
                        let addReference = true;
                        for (const varToken of varTokens) {
                            if (varToken.ignoreInOutlineView === false || varToken.token.isFromScanCommentsForReferences) {
                                ctype = (varToken.tokenType === COBOLTokenStyle.Unknown) ? ctype : varToken.tokenType;
                            } else {
                                addReference = false;
                            }
                        }
                        if (addReference) {
                            this.addVariableReference(this.sourceReferences.constantsOrVariablesReferences, currentLower, lineNumber, token.currentCol, ctype);
                        }
                    }
                    continue;
                }

                if (prevToken === "$") {
                    if (currentLower === "region") {
                        const trimmedCurrent = COBOLSourceScanner.trimLiteral(current, false);
                        const restOfLine = line.substring(token.currentCol);
                        const ctoken = this.newCOBOLToken(COBOLTokenStyle.Region, lineNumber, line, prevCurrentCol, prevToken+trimmedCurrent, restOfLine, state.currentDivision, "", false);

                        this.activeRegions.push(ctoken);
                        token.endToken();
                        continue;
                    }

                    if (currentLower === "end-region") {
                        if (this.activeRegions.length > 0) {
                            const ctoken = this.activeRegions.pop();
                            if (ctoken !== undefined) {
                                ctoken.rangeEndLine = lineNumber;
                                ctoken.rangeEndColumn = line.toLowerCase().indexOf(currentLower) + currentLower.length;
                                this.regions.push(ctoken);
                            }
                        }
                    }
                }

                if (state.declaratives !== undefined && prevTokenLower === "end" && currentLower === "declaratives") {
                    state.declaratives.rangeEndLine = lineNumber;
                    state.declaratives.rangeEndColumn = line.indexOf(current);
                    state.inDeclaratives = false;
                    state.declaratives = undefined;
                    state.declaratives = this.newCOBOLToken(COBOLTokenStyle.EndDeclaratives, lineNumber, line, currentCol, current, current, state.currentDivision, "", false);

                    continue;
                }

                //remember replace
                if (state.enable_text_replacement && currentLower === "replace") {
                    state.inReplace = true;
                    state.skipToDot = true;
                    state.captureReplaceLeft = true;
                    state.replaceLeft = state.replaceRight = "";
                    continue;
                }

                // handle sections
                if (state.currentClass === undefined && prevToken.length !== 0 && currentLower === "section" && (prevTokenLower !== "exit")) {
                    if (prevTokenLower === "declare") {
                        continue;
                    }

                    // So we need to insert a fake data division?
                    if (state.currentDivision === undefined) {
                        if (prevTokenLower === "file" ||
                            prevTokenLower === "working-storage" ||
                            prevTokenLower === "local-storage" ||
                            prevTokenLower === "screen" ||
                            prevTokenLower === "linkage") {

                            if (this.ImplicitProgramId.length !== 0) {
                                const trimmedCurrent = COBOLSourceScanner.trimLiteral(this.ImplicitProgramId, true);
                                const ctoken = this.newCOBOLToken(COBOLTokenStyle.ProgramId, lineNumber, "program-id. " + this.ImplicitProgramId, 0, trimmedCurrent, prevPlusCurrent, state.currentDivision, "", true);
                                state.programs.push(ctoken);
                                ctoken.ignoreInOutlineView = true;
                                this.ImplicitProgramId = "";        /* don't need it */
                            }
                            state.currentSection = undefined;
                            state.currentParagraph = undefined;
                            state.currentDivision = this.newCOBOLToken(COBOLTokenStyle.Division, lineNumber, "Data Division", 0, "Data", "Data Division (Optional)", state.currentDivision, "", false);
                            state.currentDivision.ignoreInOutlineView = true;
                        }
                    }

                    if (prevTokenLower === "working-storage" || prevTokenLower === "linkage" ||
                        prevTokenLower === "local-storage" || prevTokenLower === "file-control" ||
                        prevTokenLower === "file" || prevTokenLower === "screen") {
                        state.pickFields = true;
                        state.inProcedureDivision = false;
                    }

                    state.currentParagraph = undefined;
                    state.currentSection = this.newCOBOLToken(COBOLTokenStyle.Section, lineNumber, line, prevCurrentCol, prevToken, prevPlusCurrent, state.currentDivision, "", false);
                    this.sections.set(prevTokenLower, state.currentSection);
                    state.current01Group = undefined;
                    state.currentLevel = undefined;

                    continue;
                }

                // handle divisions
                if (state.captureDivisions && prevTokenLower.length !== 0 && currentLower === "division") {
                    state.currentDivision = this.newCOBOLToken(COBOLTokenStyle.Division, lineNumber, line, 0, prevToken, prevPlusCurrent, undefined, "", false);

                    if (prevTokenLower === "procedure") {
                        state.inProcedureDivision = true;
                        state.pickFields = false;
                        state.procedureDivision = state.currentDivision;

                        // create implicit section/paragraph and also a duplicate "procedure division" named fake section
                        let pname = this.implicitCount === 0 ? prevToken : prevToken + "-" + this.implicitCount;
                        const pname_lower = pname.toLowerCase();
                        const newTokenParagraph = this.newCOBOLToken(COBOLTokenStyle.Paragraph, lineNumber, line, 0, pname_lower, prevPlusCurrent, state.currentDivision, "", true);
                        newTokenParagraph.ignoreInOutlineView = true;
                        this.paragraphs.set(pname_lower, newTokenParagraph);
                        this.sourceReferences.ignoreUnusedSymbol.set(newTokenParagraph.tokenNameLower, newTokenParagraph.tokenNameLower);
                        this.sourceReferences.ignoreUnusedSymbol.set(pname_lower, pname_lower);

                        state.currentParagraph = newTokenParagraph;
                        state.currentSection = undefined;
                        state.current01Group = undefined;
                        state.currentLevel = undefined;
                        if (state.endsWithDot === false) {
                            state.pickUpUsing = true;
                        }
                        this.implicitCount++;
                    }

                    continue;
                }

                // handle entries
                if (prevTokenLowerUntrimmed === "entry") {
                    let entryStatement = prevPlusCurrent;
                    const trimmedCurrent = COBOLSourceScanner.trimLiteral(current, true);
                    const nextSTokenOrBlank = token.nextSTokenOrBlank().currentToken;
                    if (nextSTokenOrBlank === "&") {
                        entryStatement = prevToken + " " + token.compoundItems(trimmedCurrent);
                    }
                    const ctoken = this.newCOBOLToken(COBOLTokenStyle.EntryPoint, lineNumber, line, currentCol, trimmedCurrent, entryStatement, state.currentDivision, "", false);

                    state.entryPointCount++;
                    state.parameters = [];
                    state.currentProgramTarget = new CallTargetInformation(current, ctoken, true, []);
                    this.callTargets.set(trimmedCurrent, state.currentProgramTarget);
                    state.pickUpUsing = true;
                    continue;
                }

                // handle program-id
                if (prevTokenLower === "program-id") {
                    const trimmedCurrent = COBOLSourceScanner.trimLiteral(current, true);
                    const ctoken = this.newCOBOLToken(COBOLTokenStyle.ProgramId, lineNumber, line, prevCurrentCol, trimmedCurrent, prevPlusCurrent, state.currentDivision, "", false);
                    ctoken.rangeStartColumn = prevCurrentCol;
                    state.programs.push(ctoken);
                    if (state.currentDivision !== undefined) {
                        state.currentDivision.rangeEndLine = ctoken.endLine;
                        state.currentDivision.rangeEndColumn = ctoken.endColumn;
                    }
                    if (trimmedCurrent.indexOf(" ") === -1 && token.isTokenPresent("external") === false) {
                        state.parameters = [];
                        state.currentProgramTarget = new CallTargetInformation(current, ctoken, false, []);
                        this.callTargets.set(trimmedCurrent, state.currentProgramTarget);
                        this.ProgramId = trimmedCurrent;
                    }
                    this.ImplicitProgramId = "";        /* don't need it */
                    continue;
                }

                // handle class-id
                if (prevTokenLower === "class-id") {
                    const trimmedCurrent = COBOLSourceScanner.trimLiteral(current, true);
                    state.currentClass = this.newCOBOLToken(COBOLTokenStyle.ClassId, lineNumber, line, currentCol, trimmedCurrent, prevPlusCurrent, state.currentDivision, "", false);
                    state.captureDivisions = false;
                    state.currentMethod = undefined;
                    state.pickFields = true;
                    this.classes.set(trimmedCurrent, state.currentClass);

                    continue;
                }

                // handle "end class, enum, valuetype"
                if (state.currentClass !== undefined && prevTokenLower === "end" &&
                    (currentLower === "class" || currentLower === "enum" || currentLower === "valuetype" || currentLower === "interface")) {
                    state.currentClass.rangeEndLine = lineNumber;
                    state.currentClass.rangeEndColumn = line.toLowerCase().indexOf(currentLower) + currentLower.length;

                    state.currentClass = undefined;
                    state.captureDivisions = true;
                    state.currentMethod = undefined;
                    state.pickFields = false;
                    state.inProcedureDivision = false;
                    continue;
                }

                // handle enum-id
                if (prevTokenLower === "enum-id") {
                    state.currentClass = this.newCOBOLToken(COBOLTokenStyle.EnumId, lineNumber, line, currentCol, COBOLSourceScanner.trimLiteral(current, true), prevPlusCurrent, undefined, "", false);

                    state.captureDivisions = false;
                    state.currentMethod = undefined;
                    state.pickFields = true;
                    continue;
                }

                // handle interface-id
                if (prevTokenLower === "interface-id") {
                    state.currentClass = this.newCOBOLToken(COBOLTokenStyle.InterfaceId, lineNumber, line, currentCol, COBOLSourceScanner.trimLiteral(current, true), prevPlusCurrent, state.currentDivision, "", false);

                    state.pickFields = true;
                    state.captureDivisions = false;
                    state.currentMethod = undefined;
                    continue;
                }

                // handle valuetype-id
                if (prevTokenLower === "valuetype-id") {
                    state.currentClass = this.newCOBOLToken(COBOLTokenStyle.ValueTypeId, lineNumber, line, currentCol, COBOLSourceScanner.trimLiteral(current, true), prevPlusCurrent, state.currentDivision, "", false);

                    state.pickFields = true;
                    state.captureDivisions = false;
                    state.currentMethod = undefined;
                    continue;
                }

                // handle function-id
                if (prevTokenLower === "function-id") {
                    const trimmedCurrent = COBOLSourceScanner.trimLiteral(current, true);
                    state.currentFunctionId = this.newCOBOLToken(COBOLTokenStyle.FunctionId, lineNumber, line, currentCol, trimmedCurrent, prevPlusCurrent, state.currentDivision, "", false);
                    state.captureDivisions = true;
                    state.pickFields = true;
                    state.parameters = [];
                    state.currentProgramTarget = new CallTargetInformation(current, state.currentFunctionId, false, []);
                    this.functionTargets.set(trimmedCurrent, state.currentProgramTarget);

                    continue;
                }

                // handle method-id
                if (prevTokenLower === "method-id") {
                    const currentLowerTrim = COBOLSourceScanner.trimLiteral(currentLower, true);
                    const style = currentLowerTrim === "new" ? COBOLTokenStyle.Constructor : COBOLTokenStyle.MethodId;
                    const nextTokenLower = token.nextSTokenOrBlank().currentTokenLower;
                    const nextToken = token.nextSTokenOrBlank().currentToken;

                    if (nextTokenLower === "property") {
                        const nextPlusOneToken = token.nextSTokenIndex(2).currentToken;
                        const trimmedProperty = COBOLSourceScanner.trimLiteral(nextPlusOneToken, true);
                        state.currentMethod = this.newCOBOLToken(COBOLTokenStyle.Property, lineNumber, line, currentCol, trimmedProperty, nextToken + " " + nextPlusOneToken, state.currentDivision, "", false);
                        this.methods.set(trimmedProperty, state.currentMethod);
                    } else {
                        const trimmedCurrent = COBOLSourceScanner.trimLiteral(current, true);
                        state.currentMethod = this.newCOBOLToken(style, lineNumber, line, currentCol, trimmedCurrent, prevPlusCurrent, state.currentDivision, "", false);
                        this.methods.set(trimmedCurrent, state.currentMethod);
                    }

                    state.pickFields = true;
                    state.captureDivisions = false;
                    continue;
                }

                // handle "end method"
                if (state.currentMethod !== undefined && prevTokenLower === "end" && currentLower === "method") {
                    state.currentMethod.rangeEndLine = lineNumber;
                    state.currentMethod.rangeEndColumn = line.toLowerCase().indexOf(currentLower) + currentLower.length;

                    state.currentMethod = undefined;
                    state.pickFields = false;
                    continue;
                }

                // handle "end program"
                if (state.programs.length !== 0 && prevTokenLower === "end" && currentLower === "program") {
                    const currentProgram: COBOLToken | undefined = state.programs.pop();
                    let _safe_prevCurrentCol = prevCurrentCol == 0 ? 0 : prevCurrentCol - 1;
                    if (currentProgram !== undefined) {
                        currentProgram.rangeEndLine = lineNumber;
                        currentProgram.rangeEndColumn = line.length;
                    }

                    if (state.currentDivision !== undefined) {
                        state.currentDivision.rangeEndLine = lineNumber;
                        state.currentDivision.rangeEndColumn = _safe_prevCurrentCol;
                    }

                    if (state.currentSection !== undefined) {
                        state.currentSection.rangeEndLine = lineNumber;
                        state.currentSection.rangeEndColumn = _safe_prevCurrentCol;
                    }

                    if (state.currentParagraph !== undefined) {
                        state.currentParagraph.rangeEndLine = lineNumber;
                        state.currentParagraph.rangeEndColumn = _safe_prevCurrentCol;
                    }

                    state.currentDivision = undefined;
                    state.currentSection = undefined;
                    state.currentParagraph = undefined;
                    state.inProcedureDivision = false;
                    state.pickFields = false;
                    continue;
                }

                // handle "end function"
                if (state.currentFunctionId !== undefined && prevTokenLower === "end" && currentLower === "function") {
                    state.currentFunctionId.rangeEndLine = lineNumber;
                    state.currentFunctionId.rangeEndColumn = line.toLowerCase().indexOf(currentLower) + currentLower.length;
                    this.newCOBOLToken(COBOLTokenStyle.EndFunctionId, lineNumber, line, currentCol, prevToken, current, state.currentDivision, "", false);

                    state.pickFields = false;
                    state.inProcedureDivision = false;

                    if (state.currentDivision !== undefined) {
                        state.currentDivision.rangeEndLine = lineNumber;
                        // state.currentDivision.endColumn = 0;
                    }

                    if (state.currentSection !== undefined) {
                        state.currentSection.rangeEndLine = lineNumber;
                        // state.currentSection.endColumn = 0;
                    }
                    state.currentDivision = undefined;
                    state.currentSection = undefined;
                    state.currentParagraph = undefined;
                    state.procedureDivision = undefined;
                    continue;
                }


                if (prevTokenLower !== "end" && currentLower === "declaratives") {
                    state.declaratives = this.newCOBOLToken(COBOLTokenStyle.Declaratives, lineNumber, line, currentCol, current, current, state.currentDivision, "", false);
                    state.inDeclaratives = true;
                    // this.tokensInOrder.pop();       /* only interested it at the end */
                    continue;
                }

                //remember copy
                if (currentLower === "copy") {
                    const prevState = state.copybook_state;
                    state.copybook_state = new copybookState(state.current01Group);
                    state.current01Group = undefined;
                    state.inCopy = true;
                    state.inCopyStartColumn = token.currentCol;
                    state.skipToDot = true;
                    state.copybook_state.copyVerb = current;
                    state.copybook_state.copybookDepths = prevState.copybookDepths;
                    continue;
                }

                // we are in the procedure division
                if (state.captureDivisions && state.currentDivision !== undefined &&
                    state.currentDivision === state.procedureDivision && state.endsWithDot && state.prevEndsWithDot) {
                    if (!this.isValidKeyword(currentLower)) {
                        if (current.length !== 0) {
                            if (this.isParagraph(current)) {
                                if (state.currentSection !== undefined) {
                                    const newToken = this.newCOBOLToken(COBOLTokenStyle.Paragraph, lineNumber, line, currentCol, current, current, state.currentSection, "", false);
                                    this.paragraphs.set(newToken.tokenNameLower, newToken);
                                } else {
                                    const newToken = this.newCOBOLToken(COBOLTokenStyle.Paragraph, lineNumber, line, currentCol, current, current, state.currentDivision, "", false);
                                    this.paragraphs.set(newToken.tokenNameLower, newToken);
                                }
                            }
                        }
                    }
                }

                if (state.currentSection !== undefined) {
                    if (state.currentSection.tokenNameLower === "input-output") {
                        if (prevTokenLower === "fd" || prevTokenLower === "select") {
                            state.pickFields = true;
                        }
                    }

                    if (state.currentSection.tokenNameLower === "communication") {
                        if (prevTokenLower === "cd") {
                            state.pickFields = true;
                        }
                    }
                }


                // are we in the working-storage section?
                if (state.pickFields && prevToken.length > 0) {
                    /* only interesting in things that are after a number */
                    if (this.isNumber(prevToken) && !this.isNumber(current)) {
                        const isFiller: boolean = (currentLower === "filler");
                        let pickUpThisField: boolean = isFiller;
                        let trimToken = COBOLSourceScanner.trimLiteral(current, false);

                        // what other reasons do we need to pickup this line as a field?
                        if (!pickUpThisField) {
                            const compRegEx = /comp-[0-9]/;
                            // not a complete inclusive list but should cover the normal cases
                            if (currentLower === "pic" || currentLower === "picture" || compRegEx.test(currentLower) || currentLower.startsWith("binary-")) {
                                // fake up the line
                                line = prevToken + " filler ";
                                trimToken = "filler";
                                pickUpThisField = true;
                            } else {
                                // okay, current item looks like it could be a field
                                if (!this.isValidKeyword(currentLower)) {
                                    pickUpThisField = true;
                                }
                            }
                        }

                        if (pickUpThisField) {
                            if (COBOLSourceScanner.isValidLiteral(currentLower)) {
                                let style = COBOLTokenStyle.Variable;
                                const nextTokenLower = token.nextSTokenOrBlank().currentTokenLower;

                                switch (prevToken) {
                                    case "78":
                                        style = COBOLTokenStyle.Constant;
                                        break;
                                    case "88":
                                        style = COBOLTokenStyle.ConditionName;
                                        break;
                                    case "66":
                                        style = COBOLTokenStyle.RenameLevel;
                                        break;
                                }

                                let extraInfo = prevToken;
                                let redefinesPresent = false;
                                let occursPresent = false;
                                if (prevToken === "01" || prevToken === "1") {
                                    if (nextTokenLower === "redefines") {
                                        extraInfo += "-GROUP";
                                    } else if (nextTokenLower.length === 0) {
                                        extraInfo += "-GROUP";
                                    } else if (state.currentSection !== undefined && state.currentSection.tokenNameLower === "report") {
                                        extraInfo += "-GROUP";
                                    }

                                    if (token.isTokenPresent("constant")) {
                                        style = COBOLTokenStyle.Constant;
                                    }

                                    redefinesPresent = token.isTokenPresent("redefines");
                                    if (redefinesPresent) {
                                        style = COBOLTokenStyle.Union;
                                    }
                                } else {
                                    if (nextTokenLower.length === 0) {
                                        extraInfo += "-GROUP";
                                    }

                                    occursPresent = token.isTokenPresent("occurs");
                                    if (occursPresent) {
                                        extraInfo += "-OCCURS";
                                    }
                                }

                                const ctoken = this.newCOBOLToken(style, lineNumber, line, prevCurrentCol, trimToken, trimToken, state.currentDivision, extraInfo, false);
                                if (!isFiller) {
                                    this.addVariableOrConstant(currentLower, ctoken);
                                }

                                // place the 88 under the 01 item
                                if (state.currentLevel !== undefined && prevToken === "88") {
                                    state.currentLevel.rangeEndLine = ctoken.startLine;
                                    state.currentLevel.rangeEndColumn = ctoken.startColumn + ctoken.tokenName.length;
                                }

                                if (prevToken !== "88") {
                                    state.currentLevel = ctoken;
                                    // adjust group to include level number
                                    if (ctoken.rangeStartColumn > prevCurrentCol) {
                                        ctoken.rangeStartColumn = prevCurrentCol;
                                    }
                                }

                                if (prevToken === "01" || prevToken === "1" ||
                                    prevToken === "66" || prevToken === "77" || prevToken === "78") {
                                    if (nextTokenLower.length === 0 ||
                                        redefinesPresent || occursPresent ||
                                        (state.currentSection !== undefined && state.currentSection.tokenNameLower === "report" && nextTokenLower === "type")) {
                                        state.current01Group = ctoken;
                                    } else {
                                        state.current01Group = undefined;
                                    }
                                }

                                if (state.current01Group !== undefined && state.inCopy === false) {
                                    state.current01Group.rangeEndLine = ctoken.rangeStartLine;
                                    state.current01Group.rangeEndColumn = ctoken.rangeEndColumn;
                                }

                                /* if spans multiple lines, skip to dot */
                                if (state.endsWithDot === false) {
                                    state.skipToDot = true;
                                    state.addReferencesDuringSkipToTag = true;
                                }
                            }
                        }
                        continue;
                    }

                    if ((prevTokenLower === "fd"
                        || prevTokenLower === "sd"
                        || prevTokenLower === "cd"
                        || prevTokenLower === "rd"
                        || prevTokenLower === "select")
                        && !this.isValidKeyword(currentLower)) {
                        const trimToken = COBOLSourceScanner.trimLiteral(current, false);

                        if (COBOLSourceScanner.isValidLiteral(currentLower)) {
                            const variableToken = this.newCOBOLToken(COBOLTokenStyle.Variable, lineNumber, line, currentCol, trimToken, trimToken, state.currentDivision, prevTokenLower, false);
                            this.addVariableOrConstant(currentLower, variableToken);
                        }

                        if (prevTokenLower === "rd" || prevTokenLower === "select") {
                            if (prevTokenLower === "select") {
                                state.addReferencesDuringSkipToTag = true;
                            }
                            state.skipToDot = true;
                        }
                        continue;
                    }

                    // add tokens from comms section
                    if (state !== undefined && state.currentSection !== undefined && state.currentSection.tokenNameLower === "communication" && !this.isValidKeyword(currentLower)) {
                        if (prevTokenLower === "is") {
                            const trimToken = COBOLSourceScanner.trimLiteral(current, false);
                            const variableToken = this.newCOBOLToken(COBOLTokenStyle.Variable, lineNumber, line, currentCol, trimToken, trimToken, state.currentDivision, prevTokenLower, false);
                            this.addVariableOrConstant(currentLower, variableToken);
                        }
                    }
                }

                /* add reference when perform is used */
                if (this.parse4References && this.sourceReferences !== undefined) {
                    if (state.inProcedureDivision) {
                        // not interested in literals
                        if (this.isQuotedLiteral(currentLower)) {
                            continue;
                        }

                        if (this.isNumber(currentLower) === true || this.isValidKeyword(currentLower) === true) {
                            continue;
                        }

                        // if the token contain '(' or ')' then it must be a variable reference
                        if (this.containsIndex(currentLower) === false) {

                            if (prevTokenLower === "perform" || prevTokenLower === "to" || prevTokenLower === "goto" ||
                                prevTokenLower === "thru" || prevTokenLower === "through" || prevTokenLower == "procedure") {

                                /* go nn, could be "move xx to nn" or "go to nn" */
                                let sourceStyle = COBOLTokenStyle.Unknown;
                                let sharedReferences = this.sourceReferences.unknownReferences;
                                if (this.isVisibleSection(currentLower)) {
                                    sourceStyle = COBOLTokenStyle.Section;
                                    sharedReferences = this.sourceReferences.targetReferences;
                                } else if (this.isVisibleParagraph(currentLower)) {
                                    sourceStyle = COBOLTokenStyle.Paragraph;
                                    sharedReferences = this.sourceReferences.targetReferences;
                                } else if (this.constantsOrVariables.has(current)) {
                                    sourceStyle = COBOLTokenStyle.Variable;
                                    sharedReferences = this.sourceReferences.constantsOrVariablesReferences;
                                }
                                this.addTargetReference(sharedReferences, current, lineNumber, currentCol, sourceStyle, prevTokenLower);
                                continue;
                            }

                            /* is this a reference to a variable, constant or condition? */
                            const varTokens = this.constantsOrVariables.get(currentLower);
                            if (varTokens !== undefined) {
                                let ctype: COBOLTokenStyle = COBOLTokenStyle.Variable;
                                for (const varToken of varTokens) {
                                    ctype = varToken.tokenType;
                                }
                                this.addVariableReference(this.sourceReferences.constantsOrVariablesReferences, currentLower, lineNumber, token.currentCol, ctype);
                            } else {
                                // possible reference to a para or section
                                if (currentLower.length > 0) {
                                    let sourceStyle = COBOLTokenStyle.Unknown;
                                    let sharedReferences = this.sourceReferences.unknownReferences;

                                    if (this.isValidParagraphName(currentLower)) {
                                        const nextTokenLower = token.nextSTokenOrBlank().currentTokenLower;
                                        // const nextToken = token.nextSTokenOrBlank().currentToken;

                                        if (this.isVisibleSection(currentLower)) {
                                            sourceStyle = COBOLTokenStyle.Section;
                                            sharedReferences = this.sourceReferences.targetReferences;
                                            this.addTargetReference(sharedReferences, current, lineNumber, currentCol, sourceStyle, "");
                                        } else if (this.isVisibleParagraph(currentLower)) {
                                            sourceStyle = COBOLTokenStyle.Paragraph;
                                            sharedReferences = this.sourceReferences.targetReferences;
                                            this.addTargetReference(sharedReferences, current, lineNumber, currentCol, sourceStyle, "");
                                        } else {
                                            if (!(state.endsWithDot || nextTokenLower.startsWith("section"))) {
                                                this.addTargetReference(sharedReferences, current, lineNumber, currentCol, sourceStyle, "");
                                            }
                                        }
                                    }
                                }
                            }

                            continue;
                        }

                        // traverse map of possible variables that could be indexes etc.. a bit messy but works
                        for (const [pos, trimmedCurrentLower] of this.trimVariableToMap(currentLower)) {

                            if (this.isNumber(trimmedCurrentLower) === true || this.isValidKeyword(trimmedCurrentLower) === true || COBOLSourceScanner.isValidLiteral(trimmedCurrentLower) === false) {
                                continue;
                            }

                            /* is this a reference to a variable? */
                            const varTokens = this.constantsOrVariables.get(trimmedCurrentLower);
                            if (varTokens !== undefined) {
                                let ctype: COBOLTokenStyle = COBOLTokenStyle.Variable;
                                let addReference = true;
                                for (const varToken of varTokens) {
                                    if (varToken.ignoreInOutlineView === false || varToken.token.isFromScanCommentsForReferences) {
                                        ctype = (varToken.tokenType === COBOLTokenStyle.Unknown) ? ctype : varToken.tokenType;
                                    } else {
                                        addReference = false;
                                    }
                                }
                                if (addReference) {
                                    this.addVariableReference(this.sourceReferences.constantsOrVariablesReferences, trimmedCurrentLower, lineNumber, token.currentCol + pos, ctype);
                                }
                            } else {
                                this.addVariableReference(this.sourceReferences.unknownReferences, trimmedCurrentLower, lineNumber, token.currentCol + pos, COBOLTokenStyle.Unknown);
                            }
                        }
                    }

                }
            }
            catch (e) {
                this.externalFeatures.logException("COBOLScannner line error: ", e as Error);
            }
        }
        while (token.moveToNextToken() === false);

        return token;
    }

    parseExecStatement(currentExecStype: string, token: COBOLToken, lines: string) {
        if (this.currentExecToken === undefined) {
            return;
        }

        let currentLine = this.currentExecToken.startLine;
        if (currentExecStype.toLowerCase() === "sql") {
            let prevDeclare = false
            for (const line of lines.split("\n")) {

                for (const execWord of line.replace('\t', ' ').split(" ")) {
                    if (prevDeclare) {
                        //                        this.externalFeatures.logMessage(` Found declare ${execWord} at ${currentLine}`);
                        const c = new SQLDeclare(token, currentLine);
                        this.execSQLDeclare.set(execWord.toLowerCase(), c)
                        prevDeclare = false;
                    } else {
                        if (execWord.toLowerCase() === 'declare') {
                            prevDeclare = true;
                        }
                    }
                }
                currentLine++;
            }

        }
    }

    parseSQLDeclareForReferences(fileid: number, _refExecSQLDeclareName: string, refExecToken: COBOLToken, lines: string, sqldeclare: SQLDeclare) {

        let currentLine = refExecToken.startLine;
        let currentColumn = refExecToken.startColumn;
        const refExecSQLDeclareNameLower = _refExecSQLDeclareName.toLowerCase();
        for (const line of lines.split("\n")) {
            for (const execWord of line.replace('\t', ' ').split(" ")) {
                const execWordLower = execWord.toLowerCase();
                if (execWordLower === refExecSQLDeclareNameLower) {
                    currentColumn = line.toLowerCase().indexOf(execWordLower);
                    const sr = new SourceReference(fileid, currentLine, currentColumn, currentLine, currentColumn + execWord.length, COBOLTokenStyle.SQLCursor);
                    sqldeclare.sourceReferences.push(sr);
                }
            }
            currentLine++;
        }

    }



    private processCopyBook(cbInfo: copybookState): boolean {
        if (cbInfo.copybookDepths.length > this.configHandler.copybook_scan_depth) {
            return false;
        }

        cbInfo.copybookDepths.push(cbInfo);

        const state: ParseState = this.sourceReferences.state;

        let copyToken: COBOLToken | undefined = undefined;
        const isIn = cbInfo.isIn;
        const isOf = cbInfo.isOf;
        const lineNumber = cbInfo.startLineNumber;
        const tcurrentCurrentCol = cbInfo.startCol;
        const line = cbInfo.line;
        const trimmedCopyBook = cbInfo.trimmedCopyBook;
        const copyVerb = cbInfo.copyVerb;
        const copyBook = cbInfo.copyBook;

        let insertInSection = this.copybookNestedInSection ? state.currentSection : state.currentDivision;
        if (insertInSection === undefined) {
            insertInSection = state.currentDivision;
        }

        if (isIn || isOf) {
            const middleDesc = isIn ? " in " : " of ";
            const library_name_or_lit = COBOLSourceScanner.trimLiteral(cbInfo.library_name, true) + COBOLSourceScanner.trimLiteral(cbInfo.literal2, true);
            const desc: string = copyVerb + " " + copyBook + middleDesc + library_name_or_lit;
            // trim...
            copyToken = this.newCOBOLToken(COBOLTokenStyle.CopyBookInOrOf, lineNumber, line, tcurrentCurrentCol, trimmedCopyBook, desc, insertInSection, library_name_or_lit, false);
        }
        else {
            copyToken = this.newCOBOLToken(COBOLTokenStyle.CopyBook, lineNumber, line, tcurrentCurrentCol, trimmedCopyBook, copyVerb + " " + copyBook, insertInSection, "", false);
        }

        copyToken.rangeEndLine = cbInfo.endLineNumber;
        copyToken.rangeEndColumn = cbInfo.endCol;

        state.inCopy = false;

        const copybookToken = new COBOLCopybookToken(copyToken, false, cbInfo);

        const fileName = this.externalFeatures.expandLogicalCopyBookToFilenameOrEmpty(trimmedCopyBook, copyToken.extraInformation1, this.sourceHandler, this.configHandler);
        if (fileName.length === 0) {
            cbInfo.copybookDepths.pop();
            if (this.configHandler.linter_ignore_missing_copybook === false) {
                const diagMessage = `Unable to locate copybook ${trimmedCopyBook}`;
                this.diagMissingFileWarnings.set(diagMessage, new COBOLFileSymbol(this.filename, copyToken.startLine, trimmedCopyBook));
            }
            return false;
        }

        cbInfo.fileName = fileName;
        cbInfo.fileNameMod = this.externalFeatures.getFileModTimeStamp(fileName);
        if (this.copyBooksUsed.has(fileName) === false) {
            this.copyBooksUsed.set(fileName, [copybookToken]);
        } else {
            const copybooks = this.copyBooksUsed.get(fileName);
            if (copybooks != null) {
                copybooks.push(copybookToken)
            }
        }

        let count = 0;
        for (const cbi of cbInfo.copybookDepths) {
            if (cbi.fileName === cbInfo.fileName) {
                count++;
            }
        }

        if (count >= 2) {
            this.externalFeatures.logMessage(`Possible recursive COPYBOOK ${cbInfo.fileName}`);
            cbInfo.copybookDepths.pop();
            return false;
        }

        if (this.sourceReferences !== undefined) {
            if (this.parse_copybooks_for_references && fileName.length > 0) {
                cbInfo.fileName = fileName;
                const qfile = new FileSourceHandler(this.configHandler, undefined, fileName, this.externalFeatures);
                const currentTopLevel = this.sourceReferences.topLevel;
                this.sourceReferences.topLevel = false;

                const prevRepMap = state.replaceMap;

                if (this.configHandler.enable_text_replacement) {
                    state.replaceMap = new Map<string, ReplaceToken>([...cbInfo.copyReplaceMap, ...prevRepMap]);
                }
                // eslint-disable-next-line @typescript-eslint/no-unused-vars
                if (this.ScanUncachedInlineCopybook(qfile, this, false) == false) {
                    const diagMessage = `Unable perform inline copybook scan ${trimmedCopyBook}`;
                    this.diagMissingFileWarnings.set(diagMessage, new COBOLFileSymbol(this.filename, copyToken.startLine, trimmedCopyBook));
                }
                state.replaceMap = prevRepMap;
                this.sourceReferences.topLevel = currentTopLevel;
                copybookToken.scanComplete = true;
            }
        }

        cbInfo.copybookDepths.pop();
        return true;
    }

    private cobolLintLiteral = "cobol-lint";


    private processHintComments(commentLine: string, sourceFilename: string, sourceLineNumber: number) {
        const startOfTokenFor = this.configHandler.scan_comment_copybook_token;
        const startOfSourceDepIndex: number = commentLine.indexOf(startOfTokenFor);
        if (startOfSourceDepIndex !== -1) {
            const commentCommandArgs = commentLine.substring(startOfTokenFor.length + startOfSourceDepIndex).trim();
            const args = commentCommandArgs.split(" ");
            if (args.length !== 0) {
                let possRegExe: RegExp | undefined = undefined;
                for (const offset in args) {
                    const filenameTrimmed = args[offset].trim();
                    if (filenameTrimmed.startsWith("/") && filenameTrimmed.endsWith("/")) {
                        const regpart = filenameTrimmed.substring(1, filenameTrimmed.length - 1);
                        try {
                            possRegExe = new RegExp(regpart, "i");
                        }
                        catch (ex) {
                            this.generalWarnings.push(
                                new COBOLFileSymbol(this.filename, sourceLineNumber, `${ex}`)
                            )
                        }
                        continue;
                    }
                    const fileName = this.externalFeatures.expandLogicalCopyBookToFilenameOrEmpty(filenameTrimmed, "", this.sourceHandler, this.configHandler);
                    if (fileName.length > 0) {
                        if (this.copyBooksUsed.has(fileName) === false) {
                            this.copyBooksUsed.set(fileName, [COBOLCopybookToken.Null]);

                            const qfile = new FileSourceHandler(this.configHandler, possRegExe, fileName, this.externalFeatures);
                            const currentTopLevel = this.sourceReferences.topLevel;
                            this.sourceReferences.topLevel = false;

                            // eslint-disable-next-line @typescript-eslint/no-unused-vars
                            if (this.ScanUncachedInlineCopybook(qfile, this, this.configHandler.scan_comments_for_references) === false) {
                                const diagMessage = `${startOfTokenFor}: Unable to process inline copybook ${filenameTrimmed} specified in embedded comment`;
                                if (this.configHandler.linter_ignore_missing_copybook === false) {
                                    this.generalWarnings.push(new COBOLFileSymbol(sourceFilename, sourceLineNumber, diagMessage));
                                } else {
                                    this.externalFeatures.logMessage(diagMessage);
                                }
                            }
                            this.sourceReferences.topLevel = currentTopLevel;
                        }
                    } else {
                        if (this.configHandler.linter_ignore_missing_copybook === false) {
                            const diagMessage = `${startOfTokenFor}: Unable to locate copybook ${filenameTrimmed} specified in embedded comment`;
                            this.generalWarnings.push(new COBOLFileSymbol(sourceFilename, sourceLineNumber, diagMessage));
                        }
                    }
                }
            }
        }
    }

    private processCommentForLinter(commentLine: string, startOfCOBOLLint: number): void {
        const commentCommandArgs = commentLine.substring(this.cobolLintLiteral.length + startOfCOBOLLint).trim();
        let args = commentCommandArgs.split(" ");
        const command = args[0];
        args = args.slice(1);
        const commandTrimmed = command !== undefined ? command.trim() : undefined;
        if (commandTrimmed !== undefined) {
            if (commandTrimmed === CobolLinterProviderSymbols.NotReferencedMarker_external ||
                commandTrimmed === CobolLinterProviderSymbols.OLD_NotReferencedMarker_external) {
                for (const offset in args) {
                    this.sourceReferences.ignoreUnusedSymbol.set(args[offset].toLowerCase(), args[offset]);
                }
            }
        }
    }

    private lastCOBOLLS: COBOLToken | undefined = undefined;

    public processComment(configHandler: ICOBOLSettings, sourceHandler: ISourceHandlerLite, commentLine: string, sourceFilename: string, sourceLineNumber: number, startPos: number, format: ESourceFormat): void {
        this.sourceReferences.state.currentLineIsComment = true;

        // should consider other inline comments (aka terminal) and fixed position comments
        const startOfComment: number = commentLine.indexOf("*>");

        if (startOfComment !== undefined && startOfComment !== -1) {
            const trimmedLine = commentLine.substring(0, startOfComment).trimEnd();
            if (trimmedLine.length !== 0) {
                // we still have something to process
                this.sourceReferences.state.currentLineIsComment = false;
                commentLine = trimmedLine;
            }
        }

        const startOfCOBOLLint: number = commentLine.indexOf(this.cobolLintLiteral);
        if (startOfCOBOLLint !== -1) {
            this.processCommentForLinter(commentLine, startOfCOBOLLint);
        }

        if (this.scan_comments_for_hints) {
            this.processHintComments(commentLine, sourceFilename, sourceLineNumber);
        }

        // ls control
        if (this.scan_comment_for_ls_control) {
            if (this.sourceReferences.state.skipToEndLsIgnore) {
                const startOfEndLs = commentLine.indexOf(this.configHandler.scan_comment_end_ls_ignore);
                if (startOfEndLs !== -1) {
                    this.sourceReferences.state.skipToEndLsIgnore = false;
                    if (this.lastCOBOLLS !== undefined) {
                        this.lastCOBOLLS.rangeEndLine = sourceLineNumber;
                        this.lastCOBOLLS.rangeEndColumn = startOfEndLs + this.configHandler.scan_comment_begin_ls_ignore.length;

                        this.lastCOBOLLS = undefined;
                    }
                }
            } else {
                const startOfBeginLS = commentLine.indexOf(this.configHandler.scan_comment_begin_ls_ignore);
                if (startOfBeginLS !== -1) {
                    this.sourceReferences.state.skipToEndLsIgnore = true;
                    this.lastCOBOLLS = this.newCOBOLToken(COBOLTokenStyle.IgnoreLS, sourceLineNumber, commentLine, startOfBeginLS, this.configHandler.scan_comment_begin_ls_ignore, "Ignore Source", undefined, "", false);
                    this.lastCOBOLLS.rangeStartColumn = startPos;
                }
            }
        }
    }

    public findNearestSectionOrParagraph(line: number): COBOLToken | undefined {
        let nearToken: COBOLToken | undefined = undefined;
        for (const [, token] of this.sections) {
            if (line >= token.rangeStartLine && line <= token.rangeEndLine) {
                nearToken = token;
            }
        }
        for (const [, token] of this.paragraphs) {
            if (line >= token.rangeStartLine && line <= token.rangeEndLine) {
                nearToken = token;;
            }
        }

        return nearToken;
    }
}
