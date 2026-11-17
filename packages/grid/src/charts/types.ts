/** Native XLSX display model. Coordinates are zero-based sheet pixels/cells. */
export type ChartMarker = { row: number; col: number; dx: number; dy: number };
export type NativeChart = {
  id: string; part: string; sheet: string; title: string; kind: string; unsupported: string | null;
  anchor: { from: ChartMarker; to: ChartMarker | null; width: number; height: number; absolute: boolean };
  series: Array<{
    name: string; nameRef: string | null; categoriesRef: string | null; valuesRef: string | null;
    categories: Array<string | number | null>; values: Array<number | null>; color: string; width: number; markers: boolean;
    /** How the series is drawn. One chart can hold both (columns and a line).
     * Absent from an engine build that only knew lines. */
    kind?: "line" | "bar";
    /** Plotted against the second value axis (on the right). */
    secondary?: boolean;
    /** A line drawn smoothed. */
    smooth?: boolean;
  }>;
  legend: boolean; legendPosition: string; yMin: number | null; yMax: number | null; numberFormat: string; blanks: string;
  /** Bars: "col" (upright) or "bar" (lying); empty when the chart has none. */
  barDir?: string;
  /** Bars: "clustered", "stacked" or "percentStacked". */
  barGrouping?: string;
  /** Bars: the gap between categories in percent of a bar's width (150 when unsaid),
   * and how far bars side by side overlap (negative leaves a gap). */
  barGapWidth?: number | null; barOverlap?: number | null;
  /** The second value axis, when series are plotted against one. */
  y2Min?: number | null; y2Max?: number | null; numberFormat2?: string;
};
export type EditTimings = {
  evaluateMs: number; trackingMs: number; serializeMs: number;
  writeMs: number; decodeMs: number; chartsMs: number; workerMs: number;
  roundTripMs: number; applyMs: number; changedCells: number;
};
