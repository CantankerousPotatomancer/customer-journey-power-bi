import { ColorRule, FALLBACK_COLOR } from "./settings";

/**
 * Returns the first matching color rule color for a node label,
 * or the fallback color if no rule matches.
 *
 * Matching is case-insensitive substring search.
 * Rules are evaluated in order — first match wins.
 */
export function getNodeColor(label: string, rules: ColorRule[]): string {
    if (!label || !rules?.length) return FALLBACK_COLOR;

    const lower = label.toLowerCase();
    for (const rule of rules) {
        if (rule.substring && lower.includes(rule.substring.toLowerCase())) {
            return rule.color;
        }
    }
    return FALLBACK_COLOR;
}
