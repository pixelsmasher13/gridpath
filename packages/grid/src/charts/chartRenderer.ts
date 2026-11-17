// Loaded only for a visible supported chart. Native OOXML remains the save format.
//
// What is drawn: lines, bars (upright or lying; side by side, stacked, or
// stacked to 100%), and the two in one chart, with a second value axis on
// the right for the series that use one.
import { init, use, type EChartsCoreOption } from "echarts/core";
import { BarChart, LineChart } from "echarts/charts";
import { GridComponent, TitleComponent, LegendComponent, TooltipComponent } from "echarts/components";
import { SVGRenderer } from "echarts/renderers";
import type { NativeChart } from "./types";

use([BarChart, LineChart, GridComponent, TitleComponent, LegendComponent, TooltipComponent, SVGRenderer]);

/** Labels in the axis' number format. A format that says nothing (General,
 * or none) leaves the numbers as they are: 0, 0.5, 1 must not all read "0"
 * or "1". */
function axisLabel(format: string): ((value: number) => string) | undefined {
  if (!format || /^general$/i.test(format.trim())) return undefined;
  const currency = format.includes("$") ? "$" : "";
  const percent = format.includes("%");
  const decimals = format.match(/\.([0#]+)/)?.[1].length ?? 0;
  return value => currency + (percent ? value * 100 : value).toLocaleString("en-US", { maximumFractionDigits: Math.min(decimals, 10) }) + (percent ? "%" : "");
}

/** How a series is drawn: an engine build that only knew lines says nothing. */
const kindOf = (chart: NativeChart, s: NativeChart["series"][number]): "line" | "bar" => s.kind ?? (chart.kind === "barChart" ? "bar" : "line");

export function options(chart: NativeChart): EChartsCoreOption {
  const verticalLegend = chart.legendPosition === "l" || chart.legendPosition === "r";
  const first = chart.series[0];
  const categories = first?.categories.length ? first.categories.map(v => v ?? "") : first?.values.map((_, i) => String(i+1));
  const hasBars = chart.series.some(s => kindOf(chart, s) === "bar");
  const lying = hasBars && chart.barDir === "bar";
  const stacked = chart.barGrouping === "stacked" || chart.barGrouping === "percentStacked";
  const percent = chart.barGrouping === "percentStacked";
  const hasSecond = !lying && chart.series.some(s => s.secondary);
  const gapWidth = Math.max(0, chart.barGapWidth ?? 150);

  // Stacked to 100%: each bar is its share of the category's total.
  const totals = percent
    ? (first?.values ?? []).map((_, i) => chart.series.filter(s => kindOf(chart, s) === "bar").reduce((sum, s) => sum + Math.abs(s.values[i] ?? 0), 0))
    : [];
  const blank = chart.blanks === "zero" ? 0 : null;
  const dataOf = (s: NativeChart["series"][number]) =>
    s.values.map((v, i) => v === null || v === undefined ? blank : percent && kindOf(chart, s) === "bar" ? (totals[i] ? v / totals[i] : 0) : v);

  const categoryAxis = {
    type: "category" as const, data: categories,
    // Bars sit between the ticks; a line alone runs from edge to edge.
    boundaryGap: hasBars,
    axisTick: { show: false }, axisLine: { lineStyle: { color: "#d0d5dd" } }, axisLabel: { color: "#667085", hideOverlap: true },
    // Excel lists a lying bar chart's first category at the bottom.
  };
  const valueAxis = (min: number | null | undefined, max: number | null | undefined, format: string, second: boolean) => ({
    type: "value" as const,
    ...(min !== null && min !== undefined ? { min } : {}), ...(max !== null && max !== undefined ? { max } : {}),
    ...(percent && !second ? { max: 1 } : {}),
    axisLabel: (() => {
      const formatter = percent && !second ? axisLabel("0%") : axisLabel(format);
      return { ...(formatter ? { formatter } : {}), color: "#667085" };
    })(),
    splitLine: second ? { show: false } : { lineStyle: { color: "#e9edf1" } },
    ...(second ? { position: "right" as const } : {}),
  });
  const main = valueAxis(chart.yMin, chart.yMax, chart.numberFormat, false);
  const second = valueAxis(chart.y2Min, chart.y2Max, chart.numberFormat2 ?? "", true);

  // A legend at the side takes the room its longest name needs (Excel
  // shrinks the plot the same way), next to the second axis' labels.
  const longest = Math.max(0, ...chart.series.map(s => s.name.length));
  const legendSpace = Math.min(230, 46 + longest * 6.8);
  const right = (chart.legend && chart.legendPosition === "r" ? legendSpace : 24) + (hasSecond ? 38 : 0);
  return {
    animation: false,
    backgroundColor: "#ffffff",
    textStyle: { fontFamily: "Arial, sans-serif", color: "#344054" },
    title: { text: chart.title, left: "center", top: 14, textStyle: { fontSize: 16, fontWeight: "normal" } },
    legend: {
      show: chart.legend, orient: verticalLegend ? "vertical" : "horizontal",
      ...(verticalLegend ? { top: "middle", [chart.legendPosition === "l" ? "left" : "right"]: 8 }
        : { left: "center", ...(chart.legendPosition === "t" ? { top: 42 } : { bottom: 8 }) }),
    },
    grid: { left: chart.legend && chart.legendPosition === "l" ? legendSpace + 40 : 62, right, top: chart.legend && chart.legendPosition === "t" ? 76 : 55, bottom: chart.legend && !verticalLegend && chart.legendPosition !== "t" ? 58 : 35 },
    tooltip: { trigger: "axis", renderMode: "richText", confine: true, ...(hasBars ? { axisPointer: { type: "shadow" } } : {}) },
    xAxis: lying ? main : categoryAxis,
    yAxis: lying ? categoryAxis : hasSecond ? [main, second] : main,
    series: chart.series.map(s => kindOf(chart, s) === "bar"
      ? {
          name: s.name, type: "bar", data: dataOf(s),
          ...(stacked ? { stack: s.secondary && hasSecond ? "second" : "total" } : {}),
          ...(s.secondary && hasSecond ? { yAxisIndex: 1 } : {}),
          itemStyle: { color: s.color },
          // Excel's gap width is the space between categories as a share of
          // a bar; here it is a share of the category's slot.
          barCategoryGap: `${Math.round((100 * gapWidth) / (100 + gapWidth))}%`,
          ...(stacked ? {} : { barGap: `${-(chart.barOverlap ?? 0)}%` }),
          emphasis: { disabled: true },
        }
      : {
          name: s.name, type: "line", data: dataOf(s),
          ...(s.secondary && hasSecond ? { yAxisIndex: 1 } : {}),
          connectNulls: chart.blanks === "span", showSymbol: s.markers, symbol: "circle", symbolSize: 5,
          smooth: !!s.smooth,
          lineStyle: { color: s.color, width: s.width }, itemStyle: { color: s.color },
          // A line over bars is drawn on top of them.
          z: 3,
          emphasis: { disabled: true },
        }),
  };
}

export function mount(element: HTMLElement, chart: NativeChart) {
  const instance = init(element, undefined, { renderer: "svg" });
  instance.setOption(options(chart), true);
  return { update: (next: NativeChart) => instance.setOption(options(next), true), resize: () => instance.resize(), dispose: () => instance.dispose() };
}
