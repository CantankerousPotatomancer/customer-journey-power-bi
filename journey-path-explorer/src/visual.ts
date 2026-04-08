"use strict";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions       = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                   = powerbi.extensibility.visual.IVisual;
import IVisualHost               = powerbi.extensibility.visual.IVisualHost;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance      = powerbi.VisualObjectInstance;
import DataView                  = powerbi.DataView;
import FilterAction              = powerbi.FilterAction;

import {
    AdvancedFilter,
    IAdvancedFilterCondition,
    IFilterColumnTarget
} from "powerbi-models";

import { transformDataView, StepTransitions } from "./dataTransform";
import { colMatches } from "./columnMatch";
import {
    parseSettings,
    VisualSettings,
    DEFAULT_COLOR_RULES,
    buildColorRuleForFormatPane
} from "./settings";
import { PathState, createInitialState, selectNode, resetState } from "./state";
import { render } from "./renderer";

const DEBUG = true;

/**
 * When true the visual skips transform and render entirely and acts as a pure
 * diagnostics surface.  Flip this to false once the data-shape mystery is
 * resolved so normal rendering resumes.
 */
const DIAGNOSTICS_ONLY = true;

/** Snapshot of data-shape information captured at the very top of update(). */
interface DebugState {
    seq: number;
    timestamp: string;
    operationKind: string | undefined;
    dvCount: number;
    hasDv: boolean;
    hasTable: boolean;
    hasCategorical: boolean;
    hasMatrix: boolean;
    rowCount: number;
    columns: powerbi.DataViewMetadataColumn[];
}

export class Visual implements IVisual {
    private readonly target: HTMLElement;
    private readonly host: IVisualHost;
    private state: PathState;
    private transitions: StepTransitions;
    private settings: VisualSettings;
    private lastViewport: { width: number; height: number };
    private fromStepTarget: IFilterColumnTarget | null = null;
    private fromNodeTarget: IFilterColumnTarget | null = null;
    /**
     * Number of update() calls expected from Power BI as echoes of our own
     * applyJsonFilter calls. A reset issues two calls (remove + merge), so this
     * can be 2; a selection update issues one (merge only), so it is 1.
     * Using a counter instead of a boolean prevents field-addition updates from
     * being mistaken for filter echoes when multiple updates arrive concurrently.
     */
    private filterPendingCount = 0;
    /** Retained for debug overlay; null when no dataview is present. */
    private lastDv: DataView | null = null;

    /** Monotonically increasing counter; incremented at the very top of every update(). */
    private updateSeq = 0;
    private lastUpdateSummary = "";
    /** Most recent diagnostic snapshot; stored so redraw() can re-render the overlay. */
    private lastDebugState: DebugState | null = null;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;
        this.state = createInitialState();
        this.transitions = new Map();
        this.lastViewport = { width: 0, height: 0 };
        this.settings = {
            dataControls: { topN: 20, minCount: 10 },
            appearance: {
                barOpacity: 85,
                showRank: true,
                showCount: true,
                columnMinWidth: 300,
                fontSize: 13
            },
            colorRules: DEFAULT_COLOR_RULES
        };
    }

    public update(options: VisualUpdateOptions): void {
        // ── STEP 1: capture diagnostics BEFORE any logic that could throw ──────
        this.updateSeq++;
        const seq       = this.updateSeq;
        const timestamp = new Date().toISOString();

        const dvs = options.dataViews;
        const dv: DataView | undefined = dvs?.[0];

        const currentDebugState: DebugState = {
            seq,
            timestamp,
            operationKind: (options as any).operationKind,
            dvCount:        dvs?.length ?? 0,
            hasDv:          !!dv,
            hasTable:       !!dv?.table,
            hasCategorical: !!(dv as any)?.categorical,
            hasMatrix:      !!(dv as any)?.matrix,
            rowCount:       dv?.table?.rows?.length ?? 0,
            columns:        dv?.table?.columns ?? []
        };
        this.lastDebugState = currentDebugState;

        // ── STEP 2: emit full signature to console immediately ─────────────────
        console.log("UPDATE START", {
            seq,
            timestamp,
            operationKind:  (options as any).operationKind,
            dataViews:      dvs?.length ?? 0,
            hasDv:          !!dv,
            hasTable:       !!dv?.table,
            hasCategorical: !!(dv as any)?.categorical,
            hasMatrix:      !!(dv as any)?.matrix,
            rowCount:       dv?.table?.rows?.length ?? 0,
            columns: (dv?.table?.columns ?? []).map((c, i) => ({
                i,
                displayName: c.displayName,
                queryName:   c.queryName,
                roles:       c.roles,
                isMeasure:   c.isMeasure,
                type:        c.type,
                aggregate:   (c as any).aggregate,
                ref:         (c as any)?.expr?.ref,
                entity:      (c as any)?.expr?.source?.entity
            }))
        });

        // ── STEP 3: paint overlay immediately so seq number is always live ─────
        if (DEBUG) {
            this.renderDebugOverlay(currentDebugState);
        }

        // ── STEP 4: everything else wrapped so errors appear in overlay ────────
        try {
            const viewport = options.viewport;
            this.lastViewport = { width: viewport.width, height: viewport.height };

            this.target.style.width  = `${viewport.width}px`;
            this.target.style.height = `${viewport.height}px`;

            this.lastDv = dv ?? null;
            this.captureFilterTargets(dv);

            // ── DIAGNOSTICS_ONLY mode: skip transform + render ─────────────────
            if (DIAGNOSTICS_ONLY) {
                // Wipe existing content and show only the overlay
                this.target.innerHTML = "";
                if (DEBUG) {
                    this.renderDebugOverlay(currentDebugState);
                }
                return;
            }

            if (!dvs?.length || !dv?.table?.rows?.length || !dv?.table?.columns?.length) {
                this.renderNoData();
                if (DEBUG) {
                    this.renderDebugOverlay(currentDebugState);
                }
                return;
            }

            this.settings = parseSettings(dv);

            console.log("BEFORE TRANSFORM", { seq });

            if (DEBUG) {
                const REQUIRED = ["FromStep", "FromNode", "ToNode", "TransitionCount"] as const;
                const cols = dv.table?.columns ?? [];
                const colCheck = REQUIRED.map(name => ({
                    name,
                    found: cols.some(c => colMatches(c, name))
                }));
                console.log("PRE-TRANSFORM column check:", colCheck);
                const missing = colCheck.filter(x => !x.found).map(x => x.name);
                if (missing.length > 0) {
                    console.warn("PRE-TRANSFORM missing required columns:", missing);
                }
            }

            this.transitions = transformDataView(dv);

            console.log("AFTER TRANSFORM", { seq, stepCount: this.transitions.size });

            // If this update() was triggered by one of our own applyJsonFilter calls,
            // absorb it (decrement the counter) and redraw with the new data but do
            // not re-apply the filter. Otherwise (initial load, viewport resize,
            // external slicer change with no active path) kick off the step filter.
            if (this.filterPendingCount > 0) {
                this.filterPendingCount--;
            } else if (this.state.selections.length === 0) {
                // Initial load: ensure we are scoped to FromStep = 1 only.
                this.applyPathFilter(this.state);
            }

            console.log("BEFORE RENDER", { seq });
            this.redraw();
            console.log("AFTER RENDER", { seq });

        } catch (err: unknown) {
            console.error("UPDATE CATCH", { seq, err });
            this.renderFatalDebug(err, currentDebugState);
        }
    }

    /**
     * Extract { table, column } targets for FromStep and FromNode from the
     * dataview table metadata. queryName is "Table.Column".
     * Targets are reset on every call so that removing a field from the data
     * roles panel is reflected immediately — stale targets must not persist.
     */
    private captureFilterTargets(dv: DataView | undefined): void {
        this.fromStepTarget = null;
        this.fromNodeTarget = null;

        const columns = dv?.table?.columns ?? [];
        for (const col of columns) {
            const displayName = col.displayName ?? "";
            const queryName   = col.queryName   ?? "";

            const table =
                (col as any)?.expr?.source?.entity ??
                queryName.split(".")[0] ??
                null;

            const column =
                (col as any)?.expr?.ref ??
                queryName.split(".").slice(1).join(".") ??
                displayName;

            if (!table || !column) continue;

            if (colMatches(col, "FromStep")) {
                this.fromStepTarget = { table, column };
            }

            if (colMatches(col, "FromNode")) {
                this.fromNodeTarget = { table, column };
            }
        }
    }

    /**
     * Apply a Power BI filter that scopes the dataview to only the FromStep
     * values needed for the current path:
     *   FromStep IN [1, 2, ..., selections.length + 1]
     *
     * No FromNode filter is applied here. A global AND on FromNode would also
     * restrict step-1 rows, hiding all unselected starting nodes in column 0.
     * Per-column node filtering is handled client-side in getNextStepItems().
     *
     * When no selections exist (initial load / reset) we issue two calls:
     *   1. FilterAction.remove — wipes any stale filter left by a previous
     *      session (replace is unavailable; merge alone cannot remove absent
     *      filters).
     *   2. FilterAction.merge — applies the fresh FromStep = 1 constraint.
     * Each call may trigger an update() echo, so we add 2 to the counter.
     *
     * When selections are active a single merge is sufficient — one echo.
     */
    private applyPathFilter(state: PathState): void {
        if (!this.fromStepTarget) return;

        const maxStep = state.selections.length + 1;
        const stepConditions: IAdvancedFilterCondition[] = [];
        for (let s = 1; s <= maxStep; s++) {
            stepConditions.push({ operator: "Is", value: s });
        }

        const stepFilter = new AdvancedFilter(
            this.fromStepTarget,
            "Or",
            stepConditions
        );

        if (state.selections.length === 0) {
            // Remove stale filters first, then apply the step-only constraint.
            // Both calls can produce an update() echo — count both.
            this.filterPendingCount += 2;
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
            this.host.applyJsonFilter([stepFilter], "general", "filter", FilterAction.merge);
        } else {
            // Step-range filter only — one echo expected.
            this.filterPendingCount += 1;
            this.host.applyJsonFilter([stepFilter], "general", "filter", FilterAction.merge);
        }
    }

    private redraw(): void {
        render(
            this.target,
            this.state,
            this.transitions,
            this.settings,
            (columnIndex: number, node: string) => {
                this.state = selectNode(this.state, columnIndex, node);
                this.applyPathFilter(this.state);
                this.redraw();
            },
            () => {
                this.state = resetState();
                this.applyPathFilter(this.state);
                this.redraw();
            }
        );
        // render() calls root.selectAll("*").remove() which wipes the container,
        // so we must re-append the overlay after every render pass.
        if (DEBUG && this.lastDebugState) {
            this.renderDebugOverlay(this.lastDebugState);
        }
    }

    // ── Debug rendering ────────────────────────────────────────────────────────

    private ensureOverlay(): HTMLDivElement {
        const OVERLAY_ID = "__debug_overlay__";
        let overlay = this.target.querySelector<HTMLDivElement>(`#${OVERLAY_ID}`);
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.id = OVERLAY_ID;
            overlay.style.cssText = [
                "position:absolute",
                "top:8px",
                "right:8px",
                "background:rgba(0,0,0,0.82)",
                "color:#fff",
                "font-size:11px",
                "font-family:Consolas,monospace",
                "padding:8px 10px",
                "border-radius:4px",
                "z-index:9999",
                "max-width:520px",
                "line-height:1.6",
                "pointer-events:none",
                "white-space:pre"
            ].join(";");
            this.target.style.position = "relative";
            this.target.appendChild(overlay);
        }
        return overlay;
    }

    private renderDebugOverlay(state: DebugState): void {
        const overlay = this.ensureOverlay();

        const fmtTarget = (t: IFilterColumnTarget | null) =>
            t ? `${t.table}.${t.column}` : "NOT FOUND";

        const selections = this.state.selections.length > 0
            ? this.state.selections.join(", ")
            : "empty";

        const REQUIRED = ["FromStep", "FromNode", "ToNode", "TransitionCount"] as const;
        const foundLines = REQUIRED.map(name => {
            const found = state.columns.some(c => colMatches(c, name));
            return `  ${name.padEnd(16)}: ${found ? "FOUND" : "MISSING <<<"}`;
        });

        const columnLines = state.columns.map((c, i) => {
            const dn        = c.displayName ?? "(none)";
            const qn        = c.queryName   ?? "(none)";
            const ref       = (c as any)?.expr?.ref              ?? "(none)";
            const entity    = (c as any)?.expr?.source?.entity   ?? "(none)";
            const isMeasure = c.isMeasure ? "measure" : "col";
            const typeStr   = (c.type as any)?.category
                           ?? (c.type as any)?.primitiveType
                           ?? "?";
            const rolesJson = JSON.stringify(c.roles ?? {});
            const agg       = (c as any)?.aggregate ?? "-";
            return [
                `  col[${i}] dn="${dn}" qn="${qn}"`,
                `         ref="${ref}" entity="${entity}"`,
                `         ${isMeasure} type=${typeStr} agg=${agg}`,
                `         roles=${rolesJson}`
            ].join("\n");
        });

        // Full column metadata to console for deep inspection in DevTools
        console.log("TABLE COLUMNS (overlay render)", state.columns.map((c, i) => ({
            i,
            displayName: c.displayName,
            queryName:   c.queryName,
            roles:       c.roles,
            isMeasure:   c.isMeasure,
            type:        c.type,
            aggregate:   (c as any).aggregate,
            ref:         (c as any)?.expr?.ref,
            entity:      (c as any)?.expr?.source?.entity
        })));

        // Fully replace overlay content on every call — no stale DOM survives.
        overlay.textContent = [
            `[DEBUG] seq=#${state.seq}  ${state.timestamp}`,
            `operationKind:    ${state.operationKind ?? "(none)"}`,
            `DIAGNOSTICS_ONLY: ${DIAGNOSTICS_ONLY}`,
            ``,
            `── DataView shape ──────────────────────────────`,
            `dvCount:          ${state.dvCount}`,
            `hasDv:            ${state.hasDv}`,
            `hasTable:         ${state.hasTable}`,
            `hasCategorical:   ${state.hasCategorical}`,
            `hasMatrix:        ${state.hasMatrix}`,
            `rowCount:         ${state.rowCount}`,
            ``,
            `── Columns (${state.columns.length}) ──────────────────────────────`,
            ...columnLines,
            ``,
            `── Required columns ─────────────────────────────`,
            ...foundLines,
            ``,
            `── Filter targets ───────────────────────────────`,
            `FromStep target:  ${fmtTarget(this.fromStepTarget)}`,
            `FromNode target:  ${fmtTarget(this.fromNodeTarget)}`,
            ``,
            `── State ────────────────────────────────────────`,
            `Selections:       ${selections}`,
            `FilterPending:    ${this.filterPendingCount > 0} [${this.filterPendingCount}]`
        ].join("\n");
    }

    private renderFatalDebug(err: unknown, state: DebugState): void {
        const errMsg = err instanceof Error ? err.message        : String(err);
        const stack  = err instanceof Error ? (err.stack ?? "(no stack)") : "(no stack)";

        // Wipe whatever the partial render may have left
        this.target.innerHTML = "";

        const overlay = this.ensureOverlay();

        const colSummary = state.columns
            .map((c, i) => `  col[${i}] "${c.displayName}" roles=${JSON.stringify(c.roles ?? {})}`)
            .join("\n");

        overlay.textContent = [
            `[FATAL ERROR in update #${state.seq}]`,
            `timestamp:   ${state.timestamp}`,
            `operationKind: ${state.operationKind ?? "(none)"}`,
            ``,
            `── Error ────────────────────────────────────────`,
            `message:     ${errMsg}`,
            `stack:`,
            stack,
            ``,
            `── DataView state at time of throw ──────────────`,
            `hasDv:         ${state.hasDv}`,
            `hasTable:      ${state.hasTable}`,
            `hasCategorical:${state.hasCategorical}`,
            `hasMatrix:     ${state.hasMatrix}`,
            `rowCount:      ${state.rowCount}`,
            ``,
            `── Columns seen before failure ──────────────────`,
            colSummary || "  (none)"
        ].join("\n");
    }

    private renderNoData(): void {
        this.target.innerHTML = "";
        const msg = document.createElement("div");
        msg.style.cssText = [
            "display:flex",
            "align-items:center",
            "justify-content:center",
            "height:100%",
            "color:#6b6860",
            "font-size:13px",
            "font-family:Segoe UI,sans-serif",
            "font-style:italic",
            "padding:20px",
            "text-align:center"
        ].join(";");
        msg.textContent = "Add your data to the four field wells: From Step, From Node, To Node, Transition Count.";
        this.target.appendChild(msg);
    }

    /**
     * Called by Power BI to populate each section of the Format Pane.
     * Must return current values so edits in the pane round-trip correctly.
     */
    public enumerateObjectInstances(
        options: EnumerateVisualObjectInstancesOptions
    ): VisualObjectInstance[] {
        const name = options.objectName;
        const instances: VisualObjectInstance[] = [];

        switch (name) {
            case "dataControls":
                instances.push({
                    objectName: name,
                    selector: null,
                    properties: {
                        topN:     this.settings.dataControls.topN,
                        minCount: this.settings.dataControls.minCount
                    }
                });
                break;

            case "appearance":
                instances.push({
                    objectName: name,
                    selector: null,
                    properties: {
                        barOpacity:     this.settings.appearance.barOpacity,
                        showRank:       this.settings.appearance.showRank,
                        showCount:      this.settings.appearance.showCount,
                        columnMinWidth: this.settings.appearance.columnMinWidth,
                        fontSize:       this.settings.appearance.fontSize
                    }
                });
                break;

            default: {
                const match = name.match(/^colorRule(\d+)$/);
                if (match) {
                    const idx  = parseInt(match[1], 10) - 1;
                    const rule = buildColorRuleForFormatPane(idx, this.settings.colorRules);
                    instances.push({
                        objectName: name,
                        selector: null,
                        properties: {
                            substring: rule.substring,
                            color:     rule.color
                        }
                    });
                }
                break;
            }
        }

        return instances;
    }
}
