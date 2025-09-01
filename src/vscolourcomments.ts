import {
    DecorationOptions,
    DecorationRenderOptions,
    Position,
    Range,
    TextEditor,
    TextEditorDecorationType,
    ThemeColor,
    window,
    workspace,
} from "vscode";
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
        const items = workspace
            .getConfiguration(ExtensionDefaults.defaultEditorConfig)
            .get<CommentTagItem[]>(configElement);

        if (!items) return;
        tags.clear();

        for (const item of items) {
            try {
                const options = ColourTagHandler.buildDecorationOptions(item);
                tags.set(item.tag.toUpperCase(), window.createTextEditorDecorationType(options));
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
    private readonly tags = new Map<string, TextEditorDecorationType>();

    constructor() {
        super();
        this.setupTags();
    }

    public setupTags(): void {
        super.setupTags("comments_tags", this.tags);
    }

    public processComment(
        config: ICOBOLSettings,
        source: ISourceHandlerLite,
        commentLine: string,
        _filename: string,
        lineNumber: number,
        startPos: number,
        _format: ESourceFormat
    ): void {
        if (!config.enable_comment_tags) return;

        const commentUpper = commentLine.toUpperCase();
        let bestMatch: commentRange | undefined;
        let lowestPos = commentLine.length;

        for (const tag of this.tags.keys()) {
            const pos = commentUpper.indexOf(tag, startPos + 1);
            if (pos !== -1 && pos < lowestPos) {
                lowestPos = pos;
                bestMatch = new commentRange(
                    lineNumber,
                    pos,
                    config.comment_tag_word ? tag.length : commentUpper.length - startPos,
                    tag
                );
            }
        }

        if (bestMatch) source.getNotedComments().push(bestMatch);
    }

    public async updateDecorations(editor: TextEditor | undefined): Promise<void> {
        if (!editor) return;

        const config = VSCOBOLConfiguration.get_resource_settings(editor.document, VSExternalFeatures);
        if (!config.enable_comment_tags) return;

        this.clearAllDecorations(editor);

        if (VSExtensionUtils.isSupportedLanguage(editor.document) !== TextLanguage.COBOL) return;

        const scanner = VSCOBOLSourceScanner.getCachedObject(editor.document, config);
        if (!scanner) return;

        const decorationsMap = this.buildDecorations(scanner.sourceHandler.getNotedComments());

        for (const [tag, decorationType] of this.tags) {
            editor.setDecorations(decorationType, decorationsMap.get(tag) ?? []);
        }
    }

    private clearAllDecorations(editor: TextEditor): void {
        const empty: DecorationOptions[] = [];
        for (const dec of this.tags.values()) {
            editor.setDecorations(dec, empty);
        }
    }

    private buildDecorations(ranges: commentRange[]): Map<string, DecorationOptions[]> {
        const decorations = new Map<string, DecorationOptions[]>();

        for (const range of ranges) {
            const start = new Position(range.startLine, range.startColumn);
            const end = new Position(range.startLine, range.startColumn + range.length);
            const option: DecorationOptions = { range: new Range(start, end) };

            const key = range.commentStyle.toUpperCase();
            if (!decorations.has(key)) decorations.set(key, []);
            decorations.get(key)?.push(option);
        }

        return decorations;
    }
}

export const colourCommentHandler = new CommentColourHandlerImpl();