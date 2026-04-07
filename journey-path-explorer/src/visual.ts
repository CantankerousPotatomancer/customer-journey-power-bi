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
import {
    parseSettings,
    VisualSettings,
    DEFAULT_COLOR_RULES,
    buildColorRuleForFormatPane
} from "./settings";
import { PathState, createInitialState, selectNode, resetState } from "./state";
import { render } from "./renderer";

export class Visual implements IVisual {
    private readonly target: HTMLElement;
    private readonly host: IVisualHost;
    private state: PathState;
    private transitions: StepTransitions;
    private settings: VisualSettings;
    private lastViewport: { width: number; height: number };
    private fromStepTarget: IFilterColumnTarget | null = null;
    private fromNodeTarget: IFilterColumnTarget | null = null;
    /** True when the next update() is the result of our own applyJsonFilter call. */
    private filterPending = false;

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
        const viewport = options.viewport;
        this.lastViewport = { width: viewport.width, height: viewport.height };

        // Size the root container to the Power BI viewport
        this.target.style.width  = `${viewport.width}px`;
        this.target.style.height = `${viewport.height}px`;

        const dvs = options.dataViews;
        if (!dvs?.length || !dvs[0]?.table) {
            this.renderNoData();
            return;
        }

        const dv: DataView = dvs[0];
        this.settings    = parseSettings(dv);
        this.transitions = transformDataView(dv);
        this.captureFilterTargets(dv);

        // If this update was triggered by our own applyJsonFilter call,
        // preserve the current path state. Otherwise (initial load, viewport
        // resize, external slicer change) leave state as-is too — the user
        // can use Reset to clear the path.
        if (this.filterPending) {
            this.filterPending = false;
        } else if (this.state.selections.length === 0) {
            // Initial load: ensure we are scoped to FromStep = 1 only.
            this.applyPathFilter(this.state);
        }

        this.redraw();
    }

    /**
     * Extract { table, column } targets for FromStep and FromNode from the
     * dataview metadata. queryName is "Table.Column".
     */
    private captureFilterTargets(dv: DataView): void {
        const columns = dv.table?.columns ?? [];
        for (const col of columns) {
            if (!col.queryName) continue;
            const dot = col.queryName.indexOf(".");
            if (dot < 0) continue;
            const table  = col.queryName.substring(0, dot);
            const column = col.queryName.substring(dot + 1);
            if (col.displayName === "FromStep") {
                this.fromStepTarget = { table, column };
            } else if (col.displayName === "FromNode") {
                this.fromNodeTarget = { table, column };
            }
        }
    }

    /**
     * Apply a Power BI filter that scopes the dataview to only the rows
     * needed for the current path:
     *   - FromStep IN [1, 2, ..., selections.length + 1]
     *   - FromNode IN [selections..., <any step-1 nodes>]  (only when path drilled)
     *
     * Power BI ANDs separate filters, so we use a single AdvancedFilter on
     * FromStep with OR'd "Is" conditions to limit step range. This dramatically
     * shrinks the dataview compared to the full table while keeping all
     * step-1 rows visible (which the visual always needs for column 1).
     */
    private applyPathFilter(state: PathState): void {
        if (!this.fromStepTarget) return;

        const maxStep = state.selections.length + 1;
        const stepConditions: IAdvancedFilterCondition[] = [];
        for (let s = 1; s <= maxStep; s++) {
            stepConditions.push({ operator: "Is", value: s });
        }

        const filter = new AdvancedFilter(
            this.fromStepTarget,
            "Or",
            stepConditions
        );

        this.filterPending = true;
        this.host.applyJsonFilter(
            filter,
            "general",
            "filter",
            FilterAction.merge
        );
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
        msg.textContent = "Add FromStep, FromNode, ToNode, and TransitionCount fields to the Journey Data bucket to display the journey explorer.";
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
