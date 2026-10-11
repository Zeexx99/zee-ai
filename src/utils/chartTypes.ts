/**
 * Every Excel chart type offered by Draw → Plot columns, grouped like Excel's Insert Chart dialog.
 *
 * mode "series": the chart is built series by series (X column + any Y columns, need not be adjacent).
 * mode "range":  the chart type only accepts a block of cells (statistical, hierarchy, stock, surface,
 *                bubble, map …). If the chosen columns aren't side by side, Zee AI builds a linked
 *                helper block on a new sheet and charts that.
 */
export interface ChartTypeInfo {
  /** Excel.ChartType value. */
  id: string;
  label: string;
  group: string;
  mode: "series" | "range";
  /** Whether the X column is used (histogram / box & whisker only take values). */
  useX: boolean;
  minY: number;
  maxY?: number;
  /** Shown under the dropdown when this type is selected. */
  hint?: string;
  /** Has category/value axes that can carry titles. */
  axes: boolean;
}

const S = (id: string, label: string, group: string, extra: Partial<ChartTypeInfo> = {}): ChartTypeInfo => ({
  id,
  label,
  group,
  mode: "series",
  useX: true,
  minY: 1,
  axes: true,
  ...extra,
});
const R = (id: string, label: string, group: string, extra: Partial<ChartTypeInfo> = {}): ChartTypeInfo => ({
  id,
  label,
  group,
  mode: "range",
  useX: true,
  minY: 1,
  axes: false,
  ...extra,
});

const PIE_HINT = "Pie charts show one data series: tick exactly one Y column. X gives the slice names.";
const ONE_Y = (what: string) => `Tick exactly one Y column. X gives the ${what}.`;

export const CHART_TYPES: ChartTypeInfo[] = [
  // Column
  S("ColumnClustered", "Clustered column", "Column"),
  S("ColumnStacked", "Stacked column", "Column"),
  S("ColumnStacked100", "100% stacked column", "Column"),
  S("3DColumnClustered", "3-D clustered column", "Column"),
  S("3DColumnStacked", "3-D stacked column", "Column"),
  S("3DColumnStacked100", "3-D 100% stacked column", "Column"),
  S("3DColumn", "3-D column", "Column"),

  // Bar
  S("BarClustered", "Clustered bar", "Bar"),
  S("BarStacked", "Stacked bar", "Bar"),
  S("BarStacked100", "100% stacked bar", "Bar"),
  S("3DBarClustered", "3-D clustered bar", "Bar"),
  S("3DBarStacked", "3-D stacked bar", "Bar"),
  S("3DBarStacked100", "3-D 100% stacked bar", "Bar"),

  // Line
  S("Line", "Line", "Line"),
  S("LineMarkers", "Line with markers", "Line"),
  S("LineStacked", "Stacked line", "Line"),
  S("LineMarkersStacked", "Stacked line with markers", "Line"),
  S("LineStacked100", "100% stacked line", "Line"),
  S("LineMarkersStacked100", "100% stacked line with markers", "Line"),
  S("3DLine", "3-D line", "Line"),

  // Area
  S("Area", "Area", "Area"),
  S("AreaStacked", "Stacked area", "Area"),
  S("AreaStacked100", "100% stacked area", "Area"),
  S("3DArea", "3-D area", "Area"),
  S("3DAreaStacked", "3-D stacked area", "Area"),
  S("3DAreaStacked100", "3-D 100% stacked area", "Area"),

  // Scatter & bubble
  S("XYScatter", "Scatter (markers only)", "X Y (scatter) & bubble", { hint: "X should hold numbers or dates." }),
  S("XYScatterLines", "Scatter with straight lines and markers", "X Y (scatter) & bubble", { hint: "X should hold numbers or dates." }),
  S("XYScatterLinesNoMarkers", "Scatter with straight lines", "X Y (scatter) & bubble", { hint: "X should hold numbers or dates." }),
  S("XYScatterSmooth", "Scatter with smooth lines and markers", "X Y (scatter) & bubble", { hint: "X should hold numbers or dates." }),
  S("XYScatterSmoothNoMarkers", "Scatter with smooth lines", "X Y (scatter) & bubble", { hint: "X should hold numbers or dates." }),
  R("Bubble", "Bubble", "X Y (scatter) & bubble", { minY: 2, maxY: 2, axes: true, hint: "Tick exactly two Y columns: the first is the Y value, the second the bubble size." }),
  R("Bubble3DEffect", "3-D bubble", "X Y (scatter) & bubble", { minY: 2, maxY: 2, axes: true, hint: "Tick exactly two Y columns: the first is the Y value, the second the bubble size." }),

  // Pie & doughnut
  S("Pie", "Pie", "Pie & doughnut", { maxY: 1, axes: false, hint: PIE_HINT }),
  S("PieExploded", "Exploded pie", "Pie & doughnut", { maxY: 1, axes: false, hint: PIE_HINT }),
  S("3DPie", "3-D pie", "Pie & doughnut", { maxY: 1, axes: false, hint: PIE_HINT }),
  S("3DPieExploded", "3-D exploded pie", "Pie & doughnut", { maxY: 1, axes: false, hint: PIE_HINT }),
  S("PieOfPie", "Pie of pie", "Pie & doughnut", { maxY: 1, axes: false, hint: PIE_HINT }),
  S("BarOfPie", "Bar of pie", "Pie & doughnut", { maxY: 1, axes: false, hint: PIE_HINT }),
  S("Doughnut", "Doughnut", "Pie & doughnut", { axes: false, hint: "Each ticked Y column becomes one ring." }),
  S("DoughnutExploded", "Exploded doughnut", "Pie & doughnut", { axes: false, hint: "Each ticked Y column becomes one ring." }),

  // Radar
  S("Radar", "Radar", "Radar", { axes: false }),
  S("RadarMarkers", "Radar with markers", "Radar", { axes: false }),
  S("RadarFilled", "Filled radar", "Radar", { axes: false }),

  // 3-D shapes
  S("CylinderColClustered", "Clustered cylinder column", "Cylinder, cone & pyramid"),
  S("CylinderColStacked", "Stacked cylinder column", "Cylinder, cone & pyramid"),
  S("CylinderColStacked100", "100% stacked cylinder column", "Cylinder, cone & pyramid"),
  S("CylinderBarClustered", "Clustered cylinder bar", "Cylinder, cone & pyramid"),
  S("CylinderBarStacked", "Stacked cylinder bar", "Cylinder, cone & pyramid"),
  S("CylinderBarStacked100", "100% stacked cylinder bar", "Cylinder, cone & pyramid"),
  S("CylinderCol", "3-D cylinder column", "Cylinder, cone & pyramid"),
  S("ConeColClustered", "Clustered cone column", "Cylinder, cone & pyramid"),
  S("ConeColStacked", "Stacked cone column", "Cylinder, cone & pyramid"),
  S("ConeColStacked100", "100% stacked cone column", "Cylinder, cone & pyramid"),
  S("ConeBarClustered", "Clustered cone bar", "Cylinder, cone & pyramid"),
  S("ConeBarStacked", "Stacked cone bar", "Cylinder, cone & pyramid"),
  S("ConeBarStacked100", "100% stacked cone bar", "Cylinder, cone & pyramid"),
  S("ConeCol", "3-D cone column", "Cylinder, cone & pyramid"),
  S("PyramidColClustered", "Clustered pyramid column", "Cylinder, cone & pyramid"),
  S("PyramidColStacked", "Stacked pyramid column", "Cylinder, cone & pyramid"),
  S("PyramidColStacked100", "100% stacked pyramid column", "Cylinder, cone & pyramid"),
  S("PyramidBarClustered", "Clustered pyramid bar", "Cylinder, cone & pyramid"),
  S("PyramidBarStacked", "Stacked pyramid bar", "Cylinder, cone & pyramid"),
  S("PyramidBarStacked100", "100% stacked pyramid bar", "Cylinder, cone & pyramid"),
  S("PyramidCol", "3-D pyramid column", "Cylinder, cone & pyramid"),

  // Statistical
  R("Histogram", "Histogram", "Statistical", { useX: false, hint: "Only the ticked Y columns are used (the X column is ignored). Excel picks the bins automatically." }),
  R("Pareto", "Pareto", "Statistical", { maxY: 1, hint: ONE_Y("categories") }),
  R("Boxwhisker", "Box & whisker", "Statistical", { useX: false, hint: "Only the ticked Y columns are used (the X column is ignored). One box per column." }),

  // Hierarchy / flow
  R("Treemap", "Treemap", "Hierarchy & flow", { maxY: 1, hint: ONE_Y("rectangle names") }),
  R("Sunburst", "Sunburst", "Hierarchy & flow", { maxY: 1, hint: ONE_Y("ring segment names") }),
  R("Waterfall", "Waterfall", "Hierarchy & flow", { maxY: 1, hint: ONE_Y("step names (positive and negative changes)") }),
  R("Funnel", "Funnel", "Hierarchy & flow", { maxY: 1, hint: ONE_Y("stage names") }),

  // Stock
  R("StockHLC", "High-Low-Close", "Stock", { minY: 3, maxY: 3, axes: true, hint: "X = dates. Tick 3 Y columns in this sheet order: High, Low, Close." }),
  R("StockOHLC", "Open-High-Low-Close", "Stock", { minY: 4, maxY: 4, axes: true, hint: "X = dates. Tick 4 Y columns in this sheet order: Open, High, Low, Close." }),
  R("StockVHLC", "Volume-High-Low-Close", "Stock", { minY: 4, maxY: 4, axes: true, hint: "X = dates. Tick 4 Y columns in this sheet order: Volume, High, Low, Close." }),
  R("StockVOHLC", "Volume-Open-High-Low-Close", "Stock", { minY: 5, maxY: 5, axes: true, hint: "X = dates. Tick 5 Y columns in this sheet order: Volume, Open, High, Low, Close." }),

  // Surface
  R("Surface", "3-D surface", "Surface", { minY: 2, hint: "Tick 2 or more Y columns. X gives one axis and the column names give the other." }),
  R("SurfaceWireframe", "Wireframe 3-D surface", "Surface", { minY: 2, hint: "Tick 2 or more Y columns. X gives one axis and the column names give the other." }),
  R("SurfaceTopView", "Contour", "Surface", { minY: 2, hint: "Tick 2 or more Y columns. X gives one axis and the column names give the other." }),
  R("SurfaceTopViewWireframe", "Wireframe contour", "Surface", { minY: 2, hint: "Tick 2 or more Y columns. X gives one axis and the column names give the other." }),

  // Map
  R("RegionMap", "Filled map", "Map", { maxY: 1, hint: "X must hold country/region/state names. Tick one Y column. Needs an internet connection." }),
];

/** 2-D series charts can put series on a right-hand (secondary) Y axis and use a log value axis. */
export function supportsSecondaryAxis(info: ChartTypeInfo): boolean {
  return info.mode === "series" && ["Column", "Bar", "Line", "Area", "X Y (scatter) & bubble"].includes(info.group) && !/^3D/.test(info.id);
}

export function supportsLogScale(info: ChartTypeInfo): boolean {
  return info.axes && !/Stacked100/.test(info.id) && !/^3D|^Cylinder|^Cone|^Pyramid/.test(info.id);
}

export const CHART_GROUPS: string[] = Array.from(new Set(CHART_TYPES.map((t) => t.group)));

export function chartTypeInfo(id: string): ChartTypeInfo {
  return CHART_TYPES.find((t) => t.id === id) ?? CHART_TYPES[0];
}

/** Returns an error message if the Y selection doesn't fit the chart type, else null. */
export function checkYCount(info: ChartTypeInfo, count: number): string | null {
  if (count < info.minY) return info.minY === 1 ? "Tick at least one Y column." : `${info.label} needs ${info.maxY === info.minY ? "exactly" : "at least"} ${info.minY} Y columns.`;
  if (info.maxY !== undefined && count > info.maxY) return `${info.label} takes ${info.maxY === info.minY ? "exactly" : "at most"} ${info.maxY} Y column${info.maxY > 1 ? "s" : ""}.`;
  return null;
}
