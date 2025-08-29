export class SplitTokenizer {
    private static readonly wordSeperator = `~!@$%^&*()=+[{]}\\|;,<>/?`;

    public static splitArgument(input: string, ret: string[]): void {
        const sepEscaped = SplitTokenizer.wordSeperator.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&");
        const pattern = new RegExp(
            // Match double quotes, single quotes, double equals, double colon, single separators, or non-separator words
            `"[^"]*"|'[^']*'|==|::|[${sepEscaped}]|[^\\s${sepEscaped}]+`,
            "g"
        );

        const matches = input.match(pattern);
        if (matches) {
            ret.push(...matches.map(m => m.trim()).filter(m => m.length > 0));
        }
    }
}