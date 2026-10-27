/** Native XLSX display model. Coordinates are zero-based sheet pixels/cells. */
export type ChartMarker = { row: number; col: number; dx: number; dy: number };
export type NativeChart = {
  id: string; part: string; sheet: string; title: string; kind: string; unsupported: string | null;
  anchor: { from: ChartMarker; to: ChartMarker | null; width: number; height: number; absolute: boolean };
  series: Array<{
    name: string; nameRef: string | null; categoriesRef: string | null; valuesRef: string | null;
    categories: Array<string | number | null>; values: Array<number | null>; color: string; width: number; markers: boolean;
  }>;
  legend: boolean; legendPosition: string; yMin: number | null; yMax: number | null; numberFormat: string; blanks: string;
};
export type EditTimings = {
  evaluateMs: number; trackingMs: number; serializeMs: number;
  writeMs: number; decodeMs: number; chartsMs: number; workerMs: number;
  roundTripMs: number; applyMs: number; changedCells: number;
};
