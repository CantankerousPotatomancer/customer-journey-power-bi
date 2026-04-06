# Journey Path Explorer — Power BI Custom Visual

Interactive step-aware customer journey drilldown. Click any node to see what customers did next, column by column.

---

## Prerequisites

- Node.js 18 LTS
- Power BI Desktop (for testing)

```bash
npm install -g powerbi-visuals-tools   # install pbiviz CLI globally (once)
```

---

## Development

```bash
cd journey-path-explorer
npm install

# Live-reload development server (open in Power BI Desktop via developer visual)
npm start

# Build the distributable .pbiviz file
npm run build
# Output: dist/journeyPathExplorer.pbiviz
```

---

## Importing into Power BI Desktop

1. Open Power BI Desktop
2. In the Visualizations pane click **…** → **Import a visual from a file**
3. Select `dist/journeyPathExplorer.pbiviz`
4. Add the visual to a report page

---

## Data Model

| Field | Role | Type | Required |
|---|---|---|---|
| FromStep | Grouping | Integer | Yes |
| FromNode | Grouping | Text | Yes |
| ToNode | Grouping | Text | Yes |
| Transition Count | Measure | Integer | Yes |

Each row represents one directed transition. Example:

```
FromStep | FromNode     | ToNode    | TransitionCount
---------|--------------|-----------|----------------
1        | Product Page | Cart      | 4821
1        | Product Page | Exit …    | 2103
2        | Cart         | Checkout  | 2190
```

In the recommended setup, `HelpSessionType` and `HasOrder` filters are handled by report-level slicers — the visual consumes whatever filtered data Power BI provides automatically.

---

## Format Pane Settings

### Data Controls
| Setting | Default | Description |
|---|---|---|
| Top N per step | 20 | Max nodes shown per column |
| Min transition count | 10 | Hide nodes below this threshold |

### Appearance
| Setting | Default | Description |
|---|---|---|
| Bar fill opacity | 85% | Opacity of bar fill |
| Show rank numbers | On | Show 1, 2, 3… prefix |
| Show transition count | On | Show count on right |
| Column min width (px) | 300 | Min width per step column |
| Font size | 13 | Node label font size |

### Color Rules (1–20)
Each rule has a **Substring** and a **Color**. The first rule whose substring appears (case-insensitive) in a node label wins.

Default rules match the `path_explorer.html` tool:

| Substring | Color |
|---|---|
| Help Content | #e07b00 |
| Cart | #d4522a |
| Checkout | #b91c1c |
| Product Page | #0d7f6e |
| Fitguide | #6d28a0 |
| Category | #1d5fa8 |
| Search | #3b7dd8 |
| Exit \| Order | #166534 |
| Exit \| No Order | #6b7280 |
| Continued \| Order | #15803d |
| Continued \| No Order | #9ca3af |
| Other | #b0b8c4 |

Fallback (no rule matched): `#94a3b8`

---

## Known Limitations

**Transition counts are locally accurate, not path-accurate.**  
At Step 3, the count shown for a node reflects all users who made that transition from the Step 2 node — regardless of what path they took to reach Step 2. This is a deliberate tradeoff; true path-aware counts would require session-level data impractical at scale.

---

## File Structure

```
journey-path-explorer/
├── pbiviz.json          Visual metadata & entry point config
├── capabilities.json    Data roles and format pane schema
├── package.json
├── tsconfig.json
├── src/
│   ├── visual.ts        IVisual entry class, Power BI lifecycle hooks
│   ├── state.ts         Path selection state (immutable updates)
│   ├── dataTransform.ts DataView → StepTransitions map + column queries
│   ├── renderer.ts      D3 rendering (columns, bars, breadcrumb)
│   ├── colorRules.ts    Node color rule matching
│   └── settings.ts      Format pane settings parsing & defaults
├── style/
│   └── visual.less      Global styles (scrollbars, resets)
├── assets/
│   └── icon.png         Visual icon (20×20 px)
└── dist/
    └── *.pbiviz         Built output (gitignored)
```
