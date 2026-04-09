import * as d3 from "d3";
import { PathState, getBreadcrumbText } from "./state";
import { VisualSettings } from "./settings";
import {
    StepTransitions,
    ColumnItem,
    getStep1Items,
    getNextStepItems,
    sumItems
} from "./dataTransform";
import { getNodeColor } from "./colorRules";

interface ColumnDef {
    /** Display label: "STEP 1", "STEP 2", … */
    stepLabel: string;
    /** 0-based position in the visual */
    columnIndex: number;
    /** The FromStep value used to query transitions (equals columnIndex for col >= 1) */
    fromStep: number;
    /** The FromNode filter, or null for the first column */
    fromNode: string | null;
    /** Currently selected node in this column, or null */
    selected: string | null;
}

// ─── Design tokens ────────────────────────────────────────────────────────────
const BG          = "#f5f4f0";
const SURFACE     = "#ffffff";
const BORDER      = "#e2e0d8";
const TEXT_PRI    = "#1a1915";
const TEXT_SEC    = "#6b6860";
const ACCENT      = "#c96a00";
const ACCENT_BG   = "#fff7ed";
const ROW_HOVER   = "#fafafa";
const HEADER_BG   = "#f5f4f0";

// ─── Public API ───────────────────────────────────────────────────────────────

export function render(
    container: HTMLElement,
    state: PathState,
    transitions: StepTransitions,
    settings: VisualSettings,
    onNodeClick: (columnIndex: number, node: string) => void,
    onReset: () => void
): void {
    const root = d3.select(container);
    root.selectAll("*").remove();

    d3.select(container)
        .style("background", BG)
        .style("font-family", "'Segoe UI', system-ui, sans-serif")
        .style("display", "flex")
        .style("flex-direction", "column")
        .style("height", "100%")
        .style("overflow", "hidden")
        .style("box-sizing", "border-box");

    renderTopBar(root, state, onReset);
    renderColumnArea(root, state, transitions, settings, onNodeClick);
}

// ─── Top bar: breadcrumb + reset ──────────────────────────────────────────────

function renderTopBar(
    root: d3.Selection<HTMLElement, unknown, null, undefined>,
    state: PathState,
    onReset: () => void
): void {
    const bar = root.append("div")
        .style("display", "flex")
        .style("align-items", "center")
        .style("justify-content", "space-between")
        .style("padding", "6px 12px")
        .style("background", SURFACE)
        .style("border-bottom", `1px solid ${BORDER}`)
        .style("flex-shrink", "0")
        .style("min-height", "36px")
        .style("box-sizing", "border-box");

    const crumbText = getBreadcrumbText(state);

    bar.append("div")
        .style("flex", "1")
        .style("overflow", "hidden")
        .style("text-overflow", "ellipsis")
        .style("white-space", "nowrap")
        .style("font-size", "12px")
        .style("color", crumbText ? TEXT_PRI : TEXT_SEC)
        .style("font-style", crumbText ? "normal" : "italic")
        .text(crumbText || "Select a node to begin");

    bar.append("button")
        .style("background", ACCENT)
        .style("color", SURFACE)
        .style("border", "none")
        .style("border-radius", "4px")
        .style("padding", "3px 10px")
        .style("font-size", "11px")
        .style("cursor", "pointer")
        .style("margin-left", "12px")
        .style("flex-shrink", "0")
        .style("font-family", "inherit")
        .style("letter-spacing", "0.02em")
        .text("Reset Path")
        .on("mouseover", function() {
            d3.select(this).style("background", "#a85800");
        })
        .on("mouseout", function() {
            d3.select(this).style("background", ACCENT);
        })
        .on("click", () => onReset());
}

// ─── Column scroll area ───────────────────────────────────────────────────────

function renderColumnArea(
    root: d3.Selection<HTMLElement, unknown, null, undefined>,
    state: PathState,
    transitions: StepTransitions,
    settings: VisualSettings,
    onNodeClick: (columnIndex: number, node: string) => void
): void {
    const area = root.append("div")
        .style("display", "flex")
        .style("flex-direction", "row")
        .style("flex", "1")
        .style("overflow-x", "auto")
        .style("overflow-y", "hidden")
        .style("padding", "10px")
        .style("gap", "10px")
        .style("box-sizing", "border-box");

    // Build column definitions
    const columns: ColumnDef[] = [
        {
            stepLabel: "STEP 1",
            columnIndex: 0,
            fromStep: 1,
            fromNode: null,
            selected: state.selections[0] ?? null
        }
    ];

    for (let i = 0; i < state.selections.length; i++) {
        if (!state.selections[i]) break;
        columns.push({
            stepLabel: `STEP ${i + 2}`,
            columnIndex: i + 1,
            fromStep: i + 1,
            fromNode: state.selections[i],
            selected: state.selections[i + 1] ?? null
        });
    }

    console.log("RENDER columns", {
        columnCount:  columns.length,
        selectedPath: state.selections.slice(),
        columns: columns.map(c => ({
            stepLabel:   c.stepLabel,
            columnIndex: c.columnIndex,
            fromStep:    c.fromStep,
            fromNode:    c.fromNode,
            selected:    c.selected
        }))
    });

    columns.forEach(col => {
        const items =
            col.fromNode === null
                ? getStep1Items(transitions, settings.dataControls.topN, settings.dataControls.minCount)
                : getNextStepItems(transitions, col.fromStep, col.fromNode, settings.dataControls.topN, settings.dataControls.minCount);

        console.log(`RENDER col[${col.columnIndex}] ${col.stepLabel}`, {
            fromStep:   col.fromStep,
            fromNode:   col.fromNode,
            itemCount:  items.length,
            topItems:   items.slice(0, 5).map(i => `${i.label}:${i.count}`),
        });

        renderColumn(area, col, items, settings, onNodeClick);
    });

    // Scroll the column area to the rightmost position so the newly added
    // step column is always visible without the user dragging the scrollbar.
    const areaEl = area.node();
    if (areaEl) {
        requestAnimationFrame(() => {
            areaEl.scrollLeft = areaEl.scrollWidth;
        });
    }
}

// ─── Single column ────────────────────────────────────────────────────────────

function renderColumn(
    container: d3.Selection<HTMLDivElement, unknown, null, undefined>,
    col: ColumnDef,
    items: ColumnItem[],
    settings: VisualSettings,
    onNodeClick: (columnIndex: number, node: string) => void
): void {
    const minWidth    = Math.max(180, settings.appearance.columnMinWidth);
    const opacity     = Math.min(1, Math.max(0, settings.appearance.barOpacity / 100));
    const fontSize    = settings.appearance.fontSize;
    const rankFontSz  = Math.max(10, fontSize - 2);
    const countFontSz = Math.max(10, fontSize - 1);
    const total       = sumItems(items);
    const maxCount    = items.length > 0 ? items[0].count : 1;

    const colDiv = container.append("div")
        .style("min-width", `${minWidth}px`)
        .style("max-width", `${minWidth * 1.5}px`)
        .style("flex", `0 0 ${minWidth}px`)
        .style("background", SURFACE)
        .style("border", `1px solid ${BORDER}`)
        .style("border-radius", "6px")
        .style("display", "flex")
        .style("flex-direction", "column")
        .style("overflow", "hidden")
        .style("height", "100%")
        .style("box-shadow", "0 1px 3px rgba(0,0,0,0.06)");

    // ── Column header ──
    const header = colDiv.append("div")
        .style("display", "flex")
        .style("justify-content", "space-between")
        .style("align-items", "center")
        .style("padding", "7px 10px")
        .style("background", HEADER_BG)
        .style("border-bottom", `1px solid ${BORDER}`)
        .style("flex-shrink", "0");

    header.append("span")
        .style("font-size", "10px")
        .style("font-weight", "700")
        .style("color", TEXT_SEC)
        .style("letter-spacing", "0.07em")
        .text(col.stepLabel);

    header.append("span")
        .style("font-size", "10px")
        .style("color", TEXT_SEC)
        .text(`${formatCount(total)} trans`);

    // ── Rows area (scrollable) ──
    const rowsDiv = colDiv.append("div")
        .style("overflow-y", "auto")
        .style("overflow-x", "hidden")
        .style("flex", "1");

    if (items.length === 0) {
        rowsDiv.append("div")
            .style("padding", "20px 12px")
            .style("font-size", "12px")
            .style("color", TEXT_SEC)
            .style("text-align", "center")
            .style("font-style", "italic")
            .text(col.fromNode === null ? "No data" : "No further transitions");
        return;
    }

    items.forEach((item, idx) => {
        const isSelected = item.label === col.selected;
        const barColor   = getNodeColor(item.label, settings.colorRules);
        const barWidth   = maxCount > 0 ? (item.count / maxCount * 100) : 0;
        const rowBg      = isSelected ? ACCENT_BG : SURFACE;

        const row = rowsDiv.append("div")
            .style("display", "flex")
            .style("align-items", "center")
            .style("padding", "4px 8px")
            .style("cursor", "pointer")
            .style("border-bottom", `1px solid ${BG}`)
            .style("background", rowBg)
            .style("position", "relative")
            .style("min-height", "30px")
            .style("box-sizing", "border-box")
            .on("click", () => onNodeClick(col.columnIndex, item.label))
            .on("mouseover", function() {
                if (!isSelected) d3.select(this).style("background", ROW_HOVER);
            })
            .on("mouseout", function() {
                if (!isSelected) d3.select(this).style("background", rowBg);
            });

        // Bar fill (positioned behind text)
        row.append("div")
            .style("position", "absolute")
            .style("left", "0")
            .style("top", "0")
            .style("bottom", "0")
            .style("width", `${barWidth}%`)
            .style("background", barColor)
            .style("opacity", String(opacity))
            .style("pointer-events", "none");

        // Text content layer (above bar)
        const content = row.append("div")
            .style("display", "flex")
            .style("align-items", "center")
            .style("width", "100%")
            .style("position", "relative")
            .style("z-index", "1")
            .style("gap", "4px");

        if (settings.appearance.showRank) {
            content.append("span")
                .style("font-size", `${rankFontSz}px`)
                .style("color", isSelected ? ACCENT : "#b0aaa2")
                .style("min-width", "18px")
                .style("flex-shrink", "0")
                .style("text-align", "right")
                .text(String(idx + 1));
        }

        content.append("span")
            .style("flex", "1")
            .style("font-size", `${fontSize}px`)
            .style("color", isSelected ? ACCENT : TEXT_PRI)
            .style("font-weight", isSelected ? "600" : "400")
            .style("overflow", "hidden")
            .style("text-overflow", "ellipsis")
            .style("white-space", "nowrap")
            .style("padding-left", "4px")
            .attr("title", item.label)
            .text(item.label);

        if (settings.appearance.showCount) {
            content.append("span")
                .style("font-size", `${countFontSz}px`)
                .style("color", isSelected ? ACCENT : TEXT_SEC)
                .style("flex-shrink", "0")
                .style("white-space", "nowrap")
                .text(formatCount(item.count));
        }
    });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatCount(n: number): string {
    if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`;
    if (n >= 1_000)     return `${(n / 1_000).toFixed(1)}k`;
    return String(n);
}
