"use strict";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions       = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                   = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance      = powerbi.VisualObjectInstance;
import DataView                  = powerbi.DataView;

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
    private state: PathState;
    private transitions: StepTransitions;
    private settings: VisualSettings;
    private lastViewport: { width: number; height: number };

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
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

        // When filter context changes (slicer etc.), reset the path so the user
        // doesn't see stale selections that no longer exist in the new data.
        const prevSelections = this.state.selections;
        if (prevSelections.length > 0) {
            this.state = createInitialState();
        }

        this.redraw();
    }

    private redraw(): void {
        render(
            this.target,
            this.state,
            this.transitions,
            this.settings,
            (columnIndex: number, node: string) => {
                this.state = selectNode(this.state, columnIndex, node);
                this.redraw();
            },
            () => {
                this.state = resetState();
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
        msg.textContent = "Add From Step, From Node, To Node, and Transition Count fields to display the journey explorer.";
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
