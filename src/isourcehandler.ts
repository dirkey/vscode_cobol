/* eslint-disable @typescript-eslint/ban-types */
import { ESourceFormat } from "./externalfeatures";
import { ICOBOLSettings } from "./iconfiguration";

export interface ICommentCallback {
    processComment(config: ICOBOLSettings, sourceHandler: ISourceHandlerLite, commentLine: string, sourceFilename: string, sourceLineNumber:number, startPos: number, format: ESourceFormat) : void;
}

export class CommentRange {
    constructor(
        public startLine: number,
        public startColumn: number,
        public length: number,
        public commentStyle: string
    ) {}
}

export interface ISourceHandlerLite {
    getLineCount(): number;
    getLanguageId():string;
    getFilename(): string;
    getLineTabExpanded(lineNumber: number):string|undefined;
    getNotedComments(): CommentRange[];
    getCommentAtLine(lineNumber: number):string;
}

export interface ISourceHandler {
    getUriAsString(): string;
    getLineCount(): number;
    getCommentCount(): number;
    resetCommentCount():void;
    getLine(lineNumber: number, raw: boolean): string|undefined;
    getLineTabExpanded(lineNumber: number):string|undefined;
    setUpdatedLine(lineNumber: number, line:string) : void;
    getUpdatedLine(linenumber: number) : string|undefined;
    setDumpAreaA(flag: boolean): void;
    setDumpAreaBOnwards(flag: boolean): void;
    getFilename(): string;
    addCommentCallback(commentCallback: ICommentCallback):void;
    getDocumentVersionId(): BigInt;
    getIsSourceInWorkSpace(): boolean;
    getShortWorkspaceFilename(): string;
    getLanguageId():string;
    setSourceFormat(format: ESourceFormat):void;
    getNotedComments(): CommentRange[];
    getCommentAtLine(lineNumber: number):string;
    getText(startLine: number, startColumn:number, endLine: number, endColumn:number): string;
}