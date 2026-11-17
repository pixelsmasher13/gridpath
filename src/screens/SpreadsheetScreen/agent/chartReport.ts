/**
 * What the agent is told after `create_chart`: the chart as it is drawn,
 * series by series, so a wrong range shows in the result instead of on the
 * user's screen.
 */
import type { NativeChart } from "../../../../packages/grid/src/charts/types";
import { colLetters } from "./reviewGroups";

export function describeNewChart(chart: NativeChart): { summary: string; result: Record<string, unknown> } {
  const bars = chart.series.filter((s) => s.kind === "bar").length;
  const lines = chart.series.length - bars;
  const lying = chart.barDir === "bar";
  const stacked = chart.barGrouping === "stacked" ? "stacked " : chart.barGrouping === "percentStacked" ? "100% stacked " : "";
  const shape = bars > 0 && lines > 0 ? `${stacked}columns + line` : bars > 0 ? `${stacked}${lying ? "bars" : "columns"}` : "line";
  const points = Math.max(0, ...chart.series.map((s) => s.values.length));
  const summary = `${shape}, ${chart.series.length} series, ${points} point${points === 1 ? "" : "s"}`;

  const warnings: string[] = [];
  const series = chart.series.map((s) => {
    const numbers = s.values.filter((v) => typeof v === "number").length;
    if (numbers === 0) warnings.push(`"${s.name}" (${s.valuesRef}) has no numbers, so nothing is drawn for it: check the range.`);
    else if (numbers < s.values.length) {
      warnings.push(`"${s.name}": ${s.values.length - numbers} of ${s.values.length} cells are empty or not numbers and leave gaps.`);
    }
    return {
      name: s.name,
      values: s.valuesRef,
      drawn_as: s.kind === "bar" ? (lying ? "bar" : "column") : "line",
      axis: s.secondary ? "second (right)" : "main",
      points: s.values.length,
      first: s.values[0] ?? null,
      last: s.values[s.values.length - 1] ?? null,
    };
  });
  const first = chart.series[0];
  const labels = first?.categories ?? [];
  if (first?.categoriesRef && labels.every((c) => c === null || c === "")) {
    warnings.push(`The categories range (${first.categoriesRef}) is empty: check the range.`);
  }
  return {
    summary,
    result: {
      chart: {
        sheet: chart.sheet,
        top_left: `${colLetters(chart.anchor.from.col)}${chart.anchor.from.row + 1}`,
        size_px: [Math.round(chart.anchor.width), Math.round(chart.anchor.height)],
        title: chart.title,
        type: shape,
        categories: first?.categoriesRef
          ? { range: first.categoriesRef, count: labels.length, first: labels[0] ?? null, last: labels[labels.length - 1] ?? null }
          : null,
        series,
      },
      ...(warnings.length > 0 ? { warnings } : {}),
      note:
        "The chart is on the sheet now and is saved into the file as a native Excel chart. It plots the cells, so it follows later edits to them. " +
        "To change it, tell the user it has to be removed first (rejecting or undoing this change) and made again.",
    },
  };
}
