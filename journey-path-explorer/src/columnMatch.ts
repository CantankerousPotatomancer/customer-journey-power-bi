/**
 * Robust column-name matching for Power BI table dataViews.
 *
 * Power BI can decorate column display names with aggregation prefixes
 * ("Sum of FromStep", "First FromStep", etc.) and may differ in casing
 * or spacing from the raw field name. This module normalises names before
 * comparing so that matches survive those transformations.
 */

/** Aggregation prefixes Power BI may prepend to a display name. */
const AGGREGATION_PREFIXES = [
    "sum of ",
    "min of ",
    "max of ",
    "average of ",
    "avg of ",
    "count of ",
    "count (distinct) of ",
    "first ",
    "last ",
];

/**
 * Normalise a column label:
 *   1. lowercase + trim
 *   2. strip a leading aggregation prefix (first match wins)
 *   3. collapse all remaining whitespace
 */
export function normalizeColName(s: string): string {
    let v = (s ?? "").toLowerCase().trim();
    for (const prefix of AGGREGATION_PREFIXES) {
        if (v.startsWith(prefix)) {
            v = v.slice(prefix.length).trimStart();
            break;
        }
    }
    return v.replace(/\s+/g, "");
}

/**
 * Return true if a DataViewMetadataColumn matches a canonical field name
 * (e.g. "FromStep", "TransitionCount").
 *
 * Checked in order:
 *   1. displayName  (after normalisation)
 *   2. last segment of queryName  ("Schema.Table.Column" → "Column")
 *   3. expr.ref     (raw field reference in the query AST)
 */
export function colMatches(col: { displayName?: string; queryName?: string; [k: string]: any }, canonical: string): boolean {
    const target = normalizeColName(canonical);

    if (normalizeColName(col.displayName ?? "") === target) return true;

    const qn = col.queryName ?? "";
    const lastSegment = qn.split(".").pop() ?? "";
    if (normalizeColName(lastSegment) === target) return true;

    const exprRef = col?.expr?.ref ?? "";
    if (normalizeColName(exprRef) === target) return true;

    return false;
}
