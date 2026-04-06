import powerbi from "powerbi-visuals-api";
import DataViewObjects = powerbi.DataViewObjects;
import Fill = powerbi.Fill;

export interface ColorRule {
    substring: string;
    color: string;
}

export interface DataControlsSettings {
    topN: number;
    minCount: number;
}

export interface AppearanceSettings {
    barOpacity: number;
    showRank: boolean;
    showCount: boolean;
    columnMinWidth: number;
    fontSize: number;
}

export interface VisualSettings {
    dataControls: DataControlsSettings;
    appearance: AppearanceSettings;
    colorRules: ColorRule[];
}

// Default color rules matching the HTML tool
export const DEFAULT_COLOR_RULES: ColorRule[] = [
    { substring: "Help Content",       color: "#e07b00" },
    { substring: "Cart",               color: "#d4522a" },
    { substring: "Checkout",           color: "#b91c1c" },
    { substring: "Product Page",       color: "#0d7f6e" },
    { substring: "Fitguide",           color: "#6d28a0" },
    { substring: "Category",           color: "#1d5fa8" },
    { substring: "Search",             color: "#3b7dd8" },
    { substring: "Exit | Order",       color: "#166534" },
    { substring: "Exit | No Order",    color: "#6b7280" },
    { substring: "Continued | Order",  color: "#15803d" },
    { substring: "Continued | No Order", color: "#9ca3af" },
    { substring: "Other",              color: "#b0b8c4" }
];

export const FALLBACK_COLOR = "#94a3b8";

export function parseSettings(dataView: powerbi.DataView): VisualSettings {
    const objects: DataViewObjects = dataView?.metadata?.objects ?? {};

    return {
        dataControls: {
            topN:     getNumber(objects, "dataControls", "topN",     20),
            minCount: getNumber(objects, "dataControls", "minCount", 10)
        },
        appearance: {
            barOpacity:      getNumber(objects,  "appearance", "barOpacity",      85),
            showRank:        getBoolean(objects, "appearance", "showRank",        true),
            showCount:       getBoolean(objects, "appearance", "showCount",       true),
            columnMinWidth:  getNumber(objects,  "appearance", "columnMinWidth",  300),
            fontSize:        getNumber(objects,  "appearance", "fontSize",        13)
        },
        colorRules: parseColorRules(objects)
    };
}

function getNumber(objects: DataViewObjects, objectName: string, prop: string, def: number): number {
    const val = objects?.[objectName]?.[prop];
    if (val !== undefined && val !== null && !isNaN(Number(val))) return Number(val);
    return def;
}

function getBoolean(objects: DataViewObjects, objectName: string, prop: string, def: boolean): boolean {
    const val = objects?.[objectName]?.[prop];
    if (val !== undefined && val !== null) return Boolean(val);
    return def;
}

function parseColorRules(objects: DataViewObjects): ColorRule[] {
    const rules: ColorRule[] = [];
    for (let i = 1; i <= 20; i++) {
        const obj = objects?.[`colorRule${i}`];
        if (!obj) continue;
        const substring = obj["substring"] as string;
        const fill = obj["color"] as Fill;
        const color = fill?.solid?.color;
        if (substring && color) {
            rules.push({ substring, color });
        }
    }
    // Fall back to defaults if the user has never opened the format pane
    return rules.length > 0 ? rules : DEFAULT_COLOR_RULES;
}

// Used by enumerateObjectInstances to populate the format pane
export function buildColorRuleForFormatPane(
    ruleIndex: number,
    rules: ColorRule[]
): { substring: string; color: Fill } {
    const rule = rules[ruleIndex];
    if (rule) {
        return { substring: rule.substring, color: { solid: { color: rule.color } } };
    }
    return { substring: "", color: { solid: { color: FALLBACK_COLOR } } };
}
