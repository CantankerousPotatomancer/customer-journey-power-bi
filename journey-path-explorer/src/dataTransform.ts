import powerbi from "powerbi-visuals-api";

export interface ColumnItem {
    label: string;
    count: number;
}

/**
 * Nested map: fromStep → fromNode → toNode → totalCount
 *
 * This is built once from the Power BI DataView and then queried
 * on every user interaction without re-reading the data.
 */
export type StepTransitions = Map<number, Map<string, Map<string, number>>>;

export function transformDataView(dataView: powerbi.DataView): StepTransitions {
    const result: StepTransitions = new Map();

    if (!dataView?.table?.rows) return result;

    const columns = dataView.table.columns;
    const rows    = dataView.table.rows;

    // Locate each column by displayName (all columns share the single "values" role)
    const fromStepIdx        = columns.findIndex(c => c.displayName === "FromStep");
    const fromNodeIdx        = columns.findIndex(c => c.displayName === "FromNode");
    const toNodeIdx          = columns.findIndex(c => c.displayName === "ToNode");
    const countIdx           = columns.findIndex(c => c.displayName === "TransitionCount");

    if (fromStepIdx < 0 || fromNodeIdx < 0 || toNodeIdx < 0 || countIdx < 0) {
        return result;
    }

    for (const row of rows) {
        const fromStep = Number(row[fromStepIdx]);
        const fromNode = String(row[fromNodeIdx] ?? "").trim();
        const toNode   = String(row[toNodeIdx]   ?? "").trim();
        const count    = Number(row[countIdx] ?? 0);

        if (!isFinite(fromStep) || fromStep <= 0 || !fromNode || !toNode || !isFinite(count)) continue;

        let stepMap = result.get(fromStep);
        if (!stepMap) { stepMap = new Map(); result.set(fromStep, stepMap); }

        let nodeMap = stepMap.get(fromNode);
        if (!nodeMap) { nodeMap = new Map(); stepMap.set(fromNode, nodeMap); }

        nodeMap.set(toNode, (nodeMap.get(toNode) ?? 0) + count);
    }

    return result;
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
