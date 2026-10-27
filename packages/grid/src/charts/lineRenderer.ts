// Loaded only for a visible supported chart. Native OOXML remains the save format.
import { init, use, type EChartsCoreOption } from "echarts/core";
import { LineChart } from "echarts/charts";
import { GridComponent, TitleComponent, LegendComponent, TooltipComponent } from "echarts/components";
import { SVGRenderer } from "echarts/renderers";
import type { NativeChart } from "./types";

use([LineChart, GridComponent, TitleComponent, LegendComponent, TooltipComponent, SVGRenderer]);

function axisLabel(format: string): (value: number) => string {
  const currency = format.includes("$") ? "$" : "";
  const percent = format.includes("%");
  const decimals = format.match(/\.([0#]+)/)?.[1].length ?? 0;
  return value => currency + (percent ? value * 100 : value).toLocaleString("en-US", { maximumFractionDigits: Math.min(decimals, 10) }) + (percent ? "%" : "");
}

export function options(chart: NativeChart): EChartsCoreOption {
  const verticalLegend = chart.legendPosition === "l" || chart.legendPosition === "r";
  const first = chart.series[0];
  const categories = first?.categories.length ? first.categories.map(v => v ?? "") : first?.values.map((_, i) => String(i+1));
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
    grid: { left: chart.legend && chart.legendPosition === "l" ? 135 : 62, right: chart.legend && chart.legendPosition === "r" ? 135 : 24, top: chart.legend && chart.legendPosition === "t" ? 76 : 55, bottom: chart.legend && !verticalLegend && chart.legendPosition !== "t" ? 58 : 35 },
    tooltip: { trigger: "axis", renderMode: "richText", confine: true },
    xAxis: { type: "category", boundaryGap: false, data: categories, axisTick: { show: false }, axisLine: { lineStyle: { color: "#d0d5dd" } }, axisLabel: { color: "#667085", hideOverlap: true } },
    yAxis: { type: "value", ...(chart.yMin !== null ? { min: chart.yMin } : {}), ...(chart.yMax !== null ? { max: chart.yMax } : {}), axisLabel: { formatter: axisLabel(chart.numberFormat) }, splitLine: { lineStyle: { color: "#e9edf1" } } },
    series: chart.series.map(s => ({
      name: s.name, type: "line", data: s.values.map(v => v ?? (chart.blanks === "zero" ? 0 : null)),
      connectNulls: chart.blanks === "span", showSymbol: s.markers, symbol: "circle", symbolSize: 5,
      lineStyle: { color: s.color, width: s.width }, itemStyle: { color: s.color },
      emphasis: { disabled: true },
    })),
  };
}

export function mount(element: HTMLElement, chart: NativeChart) {
  const instance = init(element, undefined, { renderer: "svg" });
  instance.setOption(options(chart), true);
  return { update: (next: NativeChart) => instance.setOption(options(next), true), resize: () => instance.resize(), dispose: () => instance.dispose() };
}
