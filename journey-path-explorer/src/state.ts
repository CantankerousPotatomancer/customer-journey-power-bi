/**
 * Path selection state.
 *
 * `selections[i]` is the node selected in visual column (i+1).
 * Example: selections = ["Product Page", "Cart"]
 *   → Column 1: "Product Page" is highlighted
 *   → Column 2: "Cart" is highlighted
 *   → Column 3 is rendered showing next steps from "Cart"
 *
 * The array is always contiguous — there are no gaps.
 */
export interface PathState {
    selections: string[];
}

export function createInitialState(): PathState {
    return { selections: [] };
}

/**
 * Select a node at a given 0-based column index.
 * All selections after this column are discarded (path resets forward).
 */
export function selectNode(state: PathState, columnIndex: number, node: string): PathState {
    const selections = state.selections.slice(0, columnIndex);
    selections[columnIndex] = node;
    return { selections };
}

export function resetState(): PathState {
    return createInitialState();
}

/** The breadcrumb text for the current path */
export function getBreadcrumbText(state: PathState): string {
    const active = state.selections.filter(s => s);
    return active.length > 0 ? active.join(" → ") : "";
}
