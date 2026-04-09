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

import { transformDataViewDetailed, getStep1Items, StepTransitions, TransformDiagnostics } from "./dataTransform";
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
 * diagnostics surface.
 */
const DIAGNOSTICS_ONLY = false;

/**
 * When true, all applyJsonFilter() calls are skipped and the visual navigates
 * entirely using the data Power BI delivers on the initial (unfiltered) load.
 *
 * This is the stable path: equivalent to removing FromStep from the field
 * wells, but FromStep is still present and used by the transform.  The
 * server-side filter path has an unresolved echo-counting instability where
 * spurious Power BI updates (resize, slicer, page events) consume pending
 * echo slots, causing the merge echo to arrive as wasEcho=false and trigger
 * an unintended applyPathFilter(sel=[]) reset that clears transitions.
 *
 * Set to false only after the filter/update lifecycle is confirmed stable in
 * a controlled build.
 */
const DISABLE_SERVER_FILTER = true;

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
    /** Populated after transformDataViewDetailed() completes; null before that. */
    transformDiag: TransformDiagnostics | null;
    /** step1Items count captured just before calling render(). */
    step1ItemCount: number | null;
    /** Most recently constructed filter JSON; null if applyPathFilter has not run. */
    lastFilterJson: string | null;
}

export class Visual implements IVisual {
    private readonly target: HTMLElement;
    private readonly host: IVisualHost;
    private state: PathState;
    /**
     * Accumulated transitions across all fetched slices.
     * Never replaced on echo updates — only merged or cleared on reset.
     */
    private transitions: StepTransitions;
    private settings: VisualSettings;
    private lastViewport: { width: number; height: number };
    private fromStepTarget: IFilterColumnTarget | null = null;
    private fromNodeTarget: IFilterColumnTarget | null = null;
    /**
     * Number of update() calls expected from Power BI as echoes of our own
     * applyJsonFilter calls.
     *
     * Initial reset: 2 (remove + merge).
     * Selection update: 2 (remove + merge — prevents step-filter accumulation).
     */
    private filterPendingCount = 0;
    /** Retained for debug overlay. */
    private lastDv: DataView | null = null;

    private updateSeq = 0;
    private lastUpdateSummary = "";
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
            columns:        dv?.table?.columns ?? [],
            transformDiag:  null,
            step1ItemCount: null,
            lastFilterJson: null
        };
        this.lastDebugState = currentDebugState;

        // ── STEP 2: emit full signature to console ─────────────────────────────
        console.log("UPDATE START", {
            seq,
            timestamp,
            operationKind:      (options as any).operationKind,
            dataViews:          dvs?.length ?? 0,
            hasDv:              !!dv,
            hasTable:           !!dv?.table,
            rowCount:           dv?.table?.rows?.length ?? 0,
            filterPendingCount: this.filterPendingCount,
            selectedPath:       this.state.selections.slice(),
            accumulatedSteps:   Array.from(this.transitions.keys()).sort((a, b) => a - b),
        });

        // ── STEP 3: paint overlay immediately ─────────────────────────────────
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

            if (DIAGNOSTICS_ONLY) {
                this.target.innerHTML = "";
                if (DEBUG) this.renderDebugOverlay(currentDebugState);
                return;
            }

            // ── Echo accounting (before any early return) ──────────────────────
            // Must run here so that 0-row echoes (e.g. from the remove call)
            // correctly decrement the counter rather than being swallowed by
            // the early-return path below.
            const wasEcho = this.filterPendingCount > 0;
            if (wasEcho) {
                this.filterPendingCount--;
                console.log("UPDATE echo absorbed", { seq, filterPendingCount: this.filterPendingCount });
            }

            if (!dvs?.length || !dv?.table?.rows?.length || !dv?.table?.columns?.length) {
                // Only clear the display for a genuine (non-echo) empty update.
                // Echo updates (remove call, batched filter ops) may arrive with
                // 0 rows transiently — wiping the display for those causes the
                // visual to flicker to blank during normal navigation.
                if (!wasEcho) {
                    this.renderNoData();
                }
                if (DEBUG) this.renderDebugOverlay(currentDebugState);
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

            const { transitions: newSlice, diag: transformDiag } = transformDataViewDetailed(dv);
            currentDebugState.transformDiag = transformDiag;

            console.log("AFTER TRANSFORM", {
                seq,
                sliceStepCount:     newSlice.size,
                sliceSteps:         Array.from(newSlice.keys()).sort((a, b) => a - b),
                rowCount:           transformDiag.rowCount,
                parsedRows:         transformDiag.parsedRows,
                skippedRows:        transformDiag.skippedRows,
                distinctSteps:      transformDiag.distinctSteps,
                distinctStep1Nodes: transformDiag.distinctStep1Nodes,
                sampleRows:         transformDiag.sampleRows
            });

            // ── Merge new slice into accumulated transitions (never replace) ──────
            // On reset, applyPathFilter clears this.transitions before sending the
            // filter, so the first echo arrives into an empty map (same as replace).
            newSlice.forEach((nodeMap, step) => {
                if (!this.transitions.has(step)) {
                    this.transitions.set(step, new Map());
                }
                const existingStep = this.transitions.get(step)!;
                nodeMap.forEach((toNodes, fromNode) => {
                    // Overwrite the fromNode entry with fresh data from this slice.
                    existingStep.set(fromNode, toNodes);
                });
            });

            const accumulatedSteps = Array.from(this.transitions.keys()).sort((a, b) => a - b);

            // ── TRANSFORM RESULT diagnostic ────────────────────────────────────
            console.log("TRANSFORM RESULT", {
                seq,
                accumulatedSteps,
                step1Nodes: this.transitions.get(1)?.size ?? 0,
                step2Nodes: this.transitions.get(2)?.size ?? 0,
                step3Nodes: this.transitions.get(3)?.size ?? 0,
                step4Nodes: this.transitions.get(4)?.size ?? 0,
                selectedPath: this.state.selections.slice(),
            });

            if (DEBUG) {
                this.renderDebugOverlay(currentDebugState);
            }

            // ── Echo / filter re-trigger logic ─────────────────────────────────
            // wasEcho was set (and the counter decremented) before the early-return
            // guard above, so 0-row echoes are counted correctly.
            if (!wasEcho && this.state.selections.length === 0) {
                // Genuine new update with no active path — scope to step 1.
                console.log("UPDATE initial filter trigger", { seq });
                this.applyPathFilter(this.state);
            }

            console.log("BEFORE RENDER", {
                seq,
                selectedPath:     this.state.selections.slice(),
                accumulatedSteps,
            });
            this.redraw();
            console.log("AFTER RENDER", { seq });

        } catch (err: unknown) {
            console.error("UPDATE CATCH", { seq, err });
            this.renderFatalDebug(err, currentDebugState);
        }
    }

    /**
     * Extract { table, column } targets for FromStep and FromNode from the
     * dataview table metadata.
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
     * "Next-slice" filter strategy:
     *
     * Instead of requesting all steps 1..N cumulatively, each call requests
     * only the SINGLE next step needed, scoped to the relevant FromNode values.
     * Transitions are accumulated in this.transitions across calls so prior
     * steps are never lost.
     *
     * Filter shape by interaction:
     *   initial / reset  → FromStep = 1  (no FromNode filter)
     *   click at step N  → FromStep = N+1  AND  FromNode IN [toNodes of last selection]
     *
     * The FromNode candidates are the toNodes of the just-selected node at
     * step N — these are the likely step-(N+1) fromNodes in a sequential
     * journey. Using them keeps each query small even on large datasets.
     *
     * When no filter target is available, or when candidates cannot yet be
     * computed (fast click before prior echo), the FromNode filter is omitted
     * and only the step constraint is sent (graceful degradation).
     */
    private applyPathFilter(state: PathState): void {
        if (!this.fromStepTarget) {
            console.warn("APPLY FILTER skipped: fromStepTarget not found");
            return;
        }

        if (DISABLE_SERVER_FILTER) {
            // Server filtering disabled — navigate client-side only.
            // All step data arrives on the initial unfiltered load; no filter
            // calls are issued so the echo-counting instability cannot occur.
            console.log("APPLY FILTER skipped: DISABLE_SERVER_FILTER=true", {
                nextStep: state.selections.length + 1,
                sel:      state.selections.slice(),
            });
            return;
        }

        const sel      = state.selections;
        const nextStep = sel.length + 1;

        const fromStepCond: IAdvancedFilterCondition = { operator: "Is", value: nextStep };
        const fromStepFilter = new AdvancedFilter(this.fromStepTarget, "And", fromStepCond);

        if (sel.length === 0) {
            // ── Reset / initial: fetch step 1 only ────────────────────────────
            // Clear accumulated transitions so stale higher-step data doesn't
            // persist after a reset.
            this.transitions = new Map();

            const filterJson = JSON.stringify([fromStepFilter], null, 2);
            console.log("APPLY FILTER", {
                nextStep,
                selectedNode:   null,
                candidateCount: 0,
                isReset:        true,
                filterJson
            });
            if (this.lastDebugState) this.lastDebugState.lastFilterJson = filterJson;

            // Two calls: remove clears any stale filter, merge applies the new one.
            // Each may produce an update() echo.
            this.filterPendingCount += 2;
            this.host.applyJsonFilter(null,           "general", "filter", FilterAction.remove);
            this.host.applyJsonFilter([fromStepFilter], "general", "filter", FilterAction.merge);

        } else {
            // ── Selection: fetch the next step scoped to relevant fromNodes ────
            const selectedNode = sel[sel.length - 1];

            // Candidates = toNodes of the last selected node at step sel.length.
            // These are the expected fromNodes at step nextStep.
            // Example: sel=["Homepage"], step=1 → candidates are step-1 toNodes
            //          from "Homepage" (e.g. ["Product Page","Cart"]).
            const candidates: string[] = Array.from(
                this.transitions.get(sel.length)?.get(selectedNode)?.entries() ?? []
            )
                .sort((a, b) => b[1] - a[1])
                .map(([node]) => node);

            const filters: AdvancedFilter[] = [fromStepFilter];

            const MAX_CANDIDATES = 50;
            const candidatesAll = candidates.slice();
            const limited = candidates.slice(0, MAX_CANDIDATES);

            if (this.fromNodeTarget && limited.length > 0) {
                const nodeConds: IAdvancedFilterCondition[] = limited.map(n => ({
                    operator: "Is" as const,
                    value: n
                }));
                const fromNodeFilter = new AdvancedFilter(
                    this.fromNodeTarget,
                    nodeConds.length === 1 ? "And" : "Or",
                    ...nodeConds
                );
                filters.push(fromNodeFilter);
            }

            const mergeFilterJson = JSON.stringify(filters, null, 2);
            console.log("APPLY FILTER", {
                nextStep,
                sel: sel.slice(),
                selectedNode,
                candidatesAll:        candidatesAll.slice(),
                candidatesAfterCap:   limited.slice(),
                candidatesCapped:     candidatesAll.length > MAX_CANDIDATES,
                isReset:              false,
                fromNodeFilterApplied: filters.length > 1,
                removeFilterJson:     "null (clears step filter)",
                mergeFilterJson
            });
            if (this.lastDebugState) this.lastDebugState.lastFilterJson = mergeFilterJson;

            // Two calls: remove clears the existing step filter to prevent
            // accumulation (merge alone may AND rather than replace), then
            // merge applies the new step+node filter. Two echoes expected.
            this.filterPendingCount += 2;
            this.host.applyJsonFilter(null,    "general", "filter", FilterAction.remove);
            this.host.applyJsonFilter(filters, "general", "filter", FilterAction.merge);
        }
    }

    private redraw(): void {
        const { topN, minCount } = this.settings.dataControls;
        const step1Items = getStep1Items(this.transitions, topN, minCount);

        if (this.lastDebugState) {
            this.lastDebugState.step1ItemCount = step1Items.length;
        }

        console.log("REDRAW", {
            step1ItemCount:  step1Items.length,
            topN,
            minCount,
            selectedPath:    this.state.selections.slice(),
            accumulatedSteps: Array.from(this.transitions.keys()).sort((a, b) => a - b),
        });

        render(
            this.target,
            this.state,
            this.transitions,
            this.settings,
            (columnIndex: number, node: string) => {
                const selectedPathBefore = this.state.selections.slice();
                this.state = selectNode(this.state, columnIndex, node);
                const selectedPathAfter = this.state.selections.slice();
                console.log("CLICK", {
                    columnIndex,
                    node,
                    selectedPathBefore,
                    selectedPathAfter,
                    filterPendingCountBefore: this.filterPendingCount,
                });
                this.applyPathFilter(this.state);
                this.redraw();
            },
            () => {
                console.log("RESET", {
                    selectedPathBefore:       this.state.selections.slice(),
                    filterPendingCountBefore: this.filterPendingCount,
                });
                this.state = resetState();
                this.applyPathFilter(this.state);
                this.redraw();
            }
        );

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

        // ── Transform diagnostics section ──────────────────────────────────
        const td = state.transformDiag;
        const transformLines: string[] = [];
        if (td) {
            const ci = td.columnIndexes;
            transformLines.push(
                `── Transform diagnostics ───────────────────────`,
                `colIdx:  fromStep=${ci.fromStep} fromNode=${ci.fromNode} toNode=${ci.toNode} count=${ci.count}`,
                `rows:    ${td.rowCount} total | ${td.parsedRows} parsed | ${td.skippedRows} skipped`,
                `  invalidStep=${td.invalidFromStepCount} invalidCount=${td.invalidCountCount} emptyNode=${td.emptyNodeCount}`,
                `nanCount (first 100): ${td.nanCountSample}`,
                `distinctSteps (slice): [${td.distinctSteps.join(",")}]`,
                `step1Nodes (slice):    ${td.distinctStep1Nodes}`,
                ``
            );
            if (td.sampleRows.length > 0) {
                transformLines.push(`── Raw rows sample (first ${td.sampleRows.length}) ─────────────`);
                td.sampleRows.forEach((r, i) => {
                    transformLines.push(
                        `  [${i}] step=${JSON.stringify(r.fromStep)} fn=${JSON.stringify(r.fromNode)} tn=${JSON.stringify(r.toNode)} n=${JSON.stringify(r.count)}`
                    );
                });
                transformLines.push(``);
            }
        } else {
            transformLines.push(`── Transform diagnostics ───────────────────────`, `  (not yet run)`, ``);
        }

        // ── Accumulated transitions summary ────────────────────────────────
        const accLines: string[] = [];
        accLines.push(`── Accumulated transitions ──────────────────────`);
        const steps = Array.from(this.transitions.keys()).sort((a, b) => a - b);
        if (steps.length === 0) {
            accLines.push(`  (empty)`);
        } else {
            steps.forEach(s => {
                const sm = this.transitions.get(s)!;
                accLines.push(`  step ${s}: ${sm.size} fromNodes`);
            });
        }
        accLines.push(``);

        const step1Line = state.step1ItemCount !== null
            ? `step1Items:       ${state.step1ItemCount}  (topN=${this.settings.dataControls.topN} minCount=${this.settings.dataControls.minCount})`
            : `step1Items:       (not yet computed)`;

        // ── Filter diagnostics section ─────────────────────────────────────
        const filterLines: string[] = [];
        filterLines.push(`── Filter diagnostics ───────────────────────────`);
        filterLines.push(`FromStep target:  ${fmtTarget(this.fromStepTarget)}`);
        filterLines.push(`FromNode target:  ${fmtTarget(this.fromNodeTarget)}`);
        if (state.lastFilterJson !== null) {
            filterLines.push(`Last filter JSON:`);
            filterLines.push(state.lastFilterJson);
        } else {
            filterLines.push(`Last filter JSON: (applyPathFilter not yet called)`);
        }
        filterLines.push(``);

        overlay.textContent = [
            `[DEBUG] seq=#${state.seq}  ${state.timestamp}`,
            `operationKind:    ${state.operationKind ?? "(none)"}`,
            `DIAGNOSTICS_ONLY:       ${DIAGNOSTICS_ONLY}`,
            `DISABLE_SERVER_FILTER:  ${DISABLE_SERVER_FILTER}`,
            ``,
            `── DataView shape ──────────────────────────────`,
            `dvCount:          ${state.dvCount}`,
            `hasDv:            ${state.hasDv}`,
            `hasTable:         ${state.hasTable}`,
            `rowCount:         ${state.rowCount}`,
            ``,
            `── Columns (${state.columns.length}) ──────────────────────────────`,
            ...columnLines,
            ``,
            `── Required columns ─────────────────────────────`,
            ...foundLines,
            ``,
            ...transformLines,
            ...accLines,
            `── Render gate ──────────────────────────────────`,
            step1Line,
            ``,
            ...filterLines,
            `── State ────────────────────────────────────────`,
            `Selections:       ${selections}`,
            `FilterPending:    ${this.filterPendingCount > 0} [${this.filterPendingCount}]`
        ].join("\n");
    }

    private renderFatalDebug(err: unknown, state: DebugState): void {
        const errMsg = err instanceof Error ? err.message        : String(err);
        const stack  = err instanceof Error ? (err.stack ?? "(no stack)") : "(no stack)";

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
