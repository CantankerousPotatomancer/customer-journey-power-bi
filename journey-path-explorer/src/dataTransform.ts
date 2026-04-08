import powerbi from "powerbi-visuals-api";
import { colMatches } from "./columnMatch";

export interface ColumnItem {
    label: string;
    count: number;
}

export interface TransformDiagnostics {
    rowCount: number;
    parsedRows: number;
    skippedRows: number;
    invalidFromStepCount: number;
    invalidCountCount: number;
    emptyNodeCount: number;
    distinctSteps: number[];
    distinctStep1Nodes: number;
    columnIndexes: { fromStep: number; fromNode: number; toNode: number; count: number };
    /** Raw values from first ≤10 rows for spot-checking coercion. */
    sampleRows: Array<{ fromStep: unknown; fromNode: unknown; toNode: unknown; count: unknown }>;
    /** Number of NaN count values in the first 100 rows. */
    nanCountSample: number;
}

/**
 * Nested map: fromStep → fromNode → toNode → totalCount
 *
 * This is built once from the Power BI DataView and then queried
 * on every user interaction without re-reading the data.
 */
export type StepTransitions = Map<number, Map<string, Map<string, number>>>;

export function transformDataView(dataView: powerbi.DataView): StepTransitions {
    return transformDataViewDetailed(dataView).transitions;
}

/**
 * Same as transformDataView but also returns detailed diagnostics.
 * Use this during debugging to inspect coercion and row-level data.
 */
export function transformDataViewDetailed(dataView: powerbi.DataView): {
    transitions: StepTransitions;
    diag: TransformDiagnostics;
} {
    const transitions: StepTransitions = new Map();
    const table = dataView?.table;

    const emptyDiag = (): TransformDiagnostics => ({
        rowCount: 0, parsedRows: 0, skippedRows: 0,
        invalidFromStepCount: 0, invalidCountCount: 0, emptyNodeCount: 0,
        distinctSteps: [], distinctStep1Nodes: 0,
        columnIndexes: { fromStep: -1, fromNode: -1, toNode: -1, count: -1 },
        sampleRows: [], nanCountSample: 0
    });

    if (!table?.columns || !table?.rows) return { transitions, diag: emptyDiag() };

    const columns = table.columns;
    const rows    = table.rows;

    const fromStepIdx = columns.findIndex(c => colMatches(c, "FromStep"));
    const fromNodeIdx = columns.findIndex(c => colMatches(c, "FromNode"));
    const toNodeIdx   = columns.findIndex(c => colMatches(c, "ToNode"));
    const countIdx    = columns.findIndex(c => colMatches(c, "TransitionCount"));

    const diag: TransformDiagnostics = {
        rowCount: rows.length,
        parsedRows: 0, skippedRows: 0,
        invalidFromStepCount: 0, invalidCountCount: 0, emptyNodeCount: 0,
        distinctSteps: [], distinctStep1Nodes: 0,
        columnIndexes: { fromStep: fromStepIdx, fromNode: fromNodeIdx, toNode: toNodeIdx, count: countIdx },
        sampleRows: [], nanCountSample: 0
    };

    if (fromStepIdx < 0 || fromNodeIdx < 0 || toNodeIdx < 0 || countIdx < 0) {
        return { transitions, diag };
    }

    // Capture raw values from the first ≤10 rows for coercion spot-check
    const sampleLen = Math.min(10, rows.length);
    for (let i = 0; i < sampleLen; i++) {
        const r = rows[i];
        diag.sampleRows.push({
            fromStep: r[fromStepIdx],
            fromNode: r[fromNodeIdx],
            toNode:   r[toNodeIdx],
            count:    r[countIdx]
        });
    }

    // Count NaN count values in the first 100 rows
    for (let i = 0; i < Math.min(100, rows.length); i++) {
        if (isNaN(Number(rows[i][countIdx]))) diag.nanCountSample++;
    }

    for (const row of rows) {
        const fromStep = Number(row[fromStepIdx]);
        const fromNode = String(row[fromNodeIdx] ?? "").trim();
        const toNode   = String(row[toNodeIdx]   ?? "").trim();
        const count    = Number(row[countIdx] ?? 0);

        let skip = false;
        if (!isFinite(fromStep) || fromStep <= 0) { diag.invalidFromStepCount++; skip = true; }
        if (!isFinite(count))                      { diag.invalidCountCount++;    skip = true; }
        if (!fromNode || !toNode)                  { diag.emptyNodeCount++;       skip = true; }
        if (skip) { diag.skippedRows++; continue; }

        diag.parsedRows++;

        let stepMap = transitions.get(fromStep);
        if (!stepMap) { stepMap = new Map(); transitions.set(fromStep, stepMap); }

        let nodeMap = stepMap.get(fromNode);
        if (!nodeMap) { nodeMap = new Map(); stepMap.set(fromNode, nodeMap); }

        nodeMap.set(toNode, (nodeMap.get(toNode) ?? 0) + count);
    }

    diag.distinctSteps      = Array.from(transitions.keys()).sort((a, b) => a - b);
    diag.distinctStep1Nodes = transitions.get(1)?.size ?? 0;

    return { transitions, diag };
}

/**
 * Column 1: aggregate all toNode counts per fromNode at step 1.
 * The "label" for each item is the fromNode.
 */
export function getStep1Items(
    transitions: StepTransitions,
    topN: number,
    minCount: number
): ColumnItem[] {
    const stepMap = transitions.get(1);
    if (!stepMap) return [];

    const totals = new Map<string, number>();
    stepMap.forEach((toNodes, fromNode) => {
        let total = 0;
        toNodes.forEach(c => { total += c; });
        totals.set(fromNode, total);
    });

    return sortAndFilter(Array.from(totals.entries()), topN, minCount);
}

/**
 * Column N+1 (N >= 1): show toNode options for a specific fromStep + fromNode selection.
 */
export function getNextStepItems(
    transitions: StepTransitions,
    fromStep: number,
    fromNode: string,
    topN: number,
    minCount: number
): ColumnItem[] {
    const nodeMap = transitions.get(fromStep)?.get(fromNode);
    if (!nodeMap) return [];

    return sortAndFilter(Array.from(nodeMap.entries()), topN, minCount);
}

function sortAndFilter(
    entries: [string, number][],
    topN: number,
    minCount: number
): ColumnItem[] {
    return entries
        .filter(([, count]) => count >= minCount)
        .sort((a, b) => b[1] - a[1])
        .slice(0, topN)
        .map(([label, count]) => ({ label, count }));
}

export function sumItems(items: ColumnItem[]): number {
    return items.reduce((acc, item) => acc + item.count, 0);
}
