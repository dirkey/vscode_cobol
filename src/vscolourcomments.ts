import { DecorationOptions, DecorationRenderOptions, Position, Range, TextDocument, TextEditor, TextEditorDecorationType, ThemeColor, window, workspace } from "vscode";
import { ExtensionDefaults } from "./extensionDefaults";
import { ESourceFormat } from "./externalfeatures";
import { commentRange, ICommentCallback, ISourceHandlerLite } from "./isourcehandler";
import { VSCOBOLSourceScanner } from "./vscobolscanner";
import { VSCOBOLConfiguration } from "./vsconfiguration";
import { TextLanguage, VSExtensionUtils } from "./vsextutis";
import { VSLogger } from "./vslogger";
import { VSExternalFeatures } from "./vsexternalfeatures";
import { ICOBOLSettings } from "./iconfiguration";

interface CommentTagItem {
    tag: string;
    color?: string;
    backgroundColor?: string;
    strikethrough?: boolean;
    underline?: boolean;
    undercurl?: boolean;
    bold?: boolean;
    italic?: boolean;
    reverse?: boolean;
}

export class ColourTagHandler {

    public setupTags(configElement: string, tags: Map<string, TextEditorDecorationType>): void {
        const items = workspace.getConfiguration(ExtensionDefaults.defaultEditorConfig).get<CommentTagItem[]>(configElement);
        if (!items) return;

        tags.clear();

        for (const item of items) {
            try {
                const options: DecorationRenderOptions = ColourTagHandler.buildDecorationOptions(item);
                const decoration = window.createTextEditorDecorationType(options);
                tags.set(item.tag.toUpperCase(), decoration);
            } catch (e) {
                VSLogger.logException("Invalid comments_tags entry", e as Error);
            }
        }
    }

    private static buildDecorationOptions(item: CommentTagItem): DecorationRenderOptions {
        const options: DecorationRenderOptions = {};

        if (item.color) options.color = item.color;
        if (item.backgroundColor) options.backgroundColor = item.backgroundColor;

        const decorations: string[] = [];
        if (item.strikethrough) decorations.push("line-through solid");
        if (item.underline) decorations.push("underline solid");
        if (item.undercurl) decorations.push("underline wavy");
        if (decorations.length) options.textDecoration = decorations.join(" ");

        if (item.bold) options.fontWeight = "bold";
        if (item.italic) options.fontStyle = "italic";

        if (item.reverse) {
            options.backgroundColor = new ThemeColor("editor.foreground");
            options.color = new ThemeColor("editor.background");
        }

        return options;
    }
}

class CommentColourHandlerImpl extends ColourTagHandler implements ICommentCallback {
    static readonly emptyCommentDecoration = window.createTextEditorDecorationType({});

    private tags = new Map<string, TextEditorDecorationType>();

    constructor() {
        super();
        this.setupTags();
    }

    public setupTags() {
        super.setupTags("comments_tags", this.tags);
    }

    public processComment(
        configHandler: ICOBOLSettings,
        sourceHandler: ISourceHandlerLite,
        commentLine: string,
        _sourceFilename: string,
        sourceLineNumber: number,
        startPos: number,
        _format: ESourceFormat
    ): void {
        if (!configHandler.enable_comment_tags) return;

        const commentLineUpper = commentLine.toUpperCase();
        let lowestTag: commentRange | undefined;
        let lowestPos = commentLine.length;

        for (const tag of this.tags.keys()) {
            const pos = commentLineUpper.indexOf(tag, startPos + 1);
            if (pos !== -1 && pos < lowestPos) {
                lowestPos = pos;
                lowestTag = new commentRange(
                    sourceLineNumber,
                    pos,
                    configHandler.comment_tag_word ? tag.length : commentLineUpper.length - startPos,
                    tag
                );
            }
        }

        if (lowestTag) sourceHandler.getNotedComments().push(lowestTag);
    }

    public async updateDecorations(activeTextEditor: TextEditor | undefined) {
        if (!activeTextEditor) return;

        const configHandler = VSCOBOLConfiguration.get_resource_settings(activeTextEditor.document, VSExternalFeatures);
        if (!configHandler.enable_comment_tags) return;

        // Clear all previous decorations
        const empty: DecorationOptions[] = [];
        for (const dec of this.tags.values()) {
            activeTextEditor.setDecorations(dec, empty);
        }

        const doc: TextDocument = activeTextEditor.document;
        if (VSExtensionUtils.isSupportedLanguage(doc) !== TextLanguage.COBOL) return;

        const scanner = VSCOBOLSourceScanner.getCachedObject(doc, configHandler);
        if (!scanner) return;

        const decorationsMap = new Map<string, DecorationOptions[]>();
        for (const range of scanner.sourceHandler.getNotedComments()) {
            const startPos = new Position(range.startLine, range.startColumn);
            const endPos = new Position(range.startLine, range.startColumn + range.length);
            const decoration: DecorationOptions = { range: new Range(startPos, endPos) };

            const key = range.commentStyle.toUpperCase();
            if (!decorationsMap.has(key)) decorationsMap.set(key, []);
            decorationsMap.get(key)?.push(decoration);
        }

        for (const [tag, decorationType] of this.tags) {
            activeTextEditor.setDecorations(decorationType, decorationsMap.get(tag) ?? []);
        }
    }
}

export const colourCommentHandler = new CommentColourHandlerImpl();