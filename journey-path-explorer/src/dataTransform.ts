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
    const categorical = dataView?.categorical;
    if (!categorical?.categories || !categorical?.values) return result;

    // Find each category column by role name via source.roles
    const categories = categorical.categories;
    const fromStepCat = categories.find(c => c.source.roles["fromStep"]);
    const fromNodeCat = categories.find(c => c.source.roles["fromNode"]);
    const toNodeCat   = categories.find(c => c.source.roles["toNode"]);
    const countCol    = categorical.values.find(v => v.source.roles["transitionCount"]);

    if (!fromStepCat || !fromNodeCat || !toNodeCat || !countCol) return result;

    const len = fromStepCat.values.length;
    for (let i = 0; i < len; i++) {
        const fromStep = Number(fromStepCat.values[i]);
        const fromNode = String(fromNodeCat.values[i] ?? "").trim();
        const toNode   = String(toNodeCat.values[i]   ?? "").trim();
        const count    = Number(countCol.values[i] ?? 0);

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
