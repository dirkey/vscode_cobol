import { ICOBOLSettings, intellisenseStyle } from "./iconfiguration";

export class VSCustomIntelliseRules {
    public static Default: VSCustomIntelliseRules = new VSCustomIntelliseRules();

    private customRules = new Map<string, intellisenseStyle>();
    private customStartsWithRules = new Map<string, intellisenseStyle>();

    constructor() {
        VSCustomIntelliseRules.Default = this;
    }

    /** Refresh the rules from the given settings */
    public refreshConfiguration(settings: ICOBOLSettings): void {
        this.customRules.clear();
        this.customStartsWithRules.clear();

        for (const ruleString of settings.custom_intellisense_rules) {
            const colonPos = ruleString.indexOf(":");
            if (colonPos === -1) continue;

            const key = ruleString.substring(0, colonPos).toLowerCase();
            const styleChar = ruleString.charAt(colonPos + 1);

            let style: intellisenseStyle = intellisenseStyle.Unchanged;
            switch (styleChar) {
                case "u": style = intellisenseStyle.UpperCase; break;
                case "l": style = intellisenseStyle.LowerCase; break;
                case "c": style = intellisenseStyle.CamelCase; break;
                case "=": style = intellisenseStyle.Unchanged; break;
            }

            if (key.endsWith("*")) {
                this.customStartsWithRules.set(key.slice(0, -1), style);
            } else {
                this.customRules.set(key, style);
            }
        }
    }

    /**
     * Find the custom intellisense style for a given key.
     * Returns defaultStyle if no rule applies.
     */
    public findCustomIStyle(settings: ICOBOLSettings, key: string, defaultStyle: intellisenseStyle): intellisenseStyle {
        this.refreshConfiguration(settings);

        const keyLower = key.toLowerCase();

        // Exact match
        const exactMatch = this.customRules.get(keyLower);
        if (exactMatch !== undefined) return exactMatch;

        // Starts-with match
        for (const [prefix, style] of this.customStartsWithRules) {
            if (keyLower.startsWith(prefix)) return style;
        }

        return defaultStyle;
    }
}
