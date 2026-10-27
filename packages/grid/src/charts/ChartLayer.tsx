import React, { forwardRef, useEffect, useImperativeHandle, useLayoutEffect, useRef, useState, useSyncExternalStore } from "react";
import type { GridEngine } from "../engine";
import type { NativeChart } from "./types";
import { HEADER_H, HEADER_W, type SheetGeometry } from "../../../../src/screens/SpreadsheetScreen/viewportModel";

type View = { sheet: string; geometry: SheetGeometry; scroller: HTMLElement; scale: number };
export type ChartLayerHandle = { sync(): void };
const EMPTY: NativeChart[] = [];
const empty = () => EMPTY;
const noSubscribe = () => () => {};

function ChartView({ chart }: { chart: NativeChart }) {
  const host = useRef<HTMLDivElement>(null);
  const renderer = useRef<ReturnType<typeof import("./lineRenderer").mount> | null>(null);
  const latest = useRef(chart);
  latest.current = chart;
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    const element = host.current;
    if (!element || chart.unsupported) return;
    let cancelled = false;
    let resize: ResizeObserver | undefined;
    import("./lineRenderer").then(({ mount }) => {
      if (cancelled) return;
      renderer.current = mount(element, latest.current);
      resize = new ResizeObserver(() => renderer.current?.resize());
      resize.observe(element);
    }).catch(() => { if (!cancelled) setError("Chart display could not load. The native chart is retained on save."); });
    return () => { cancelled = true; resize?.disconnect(); renderer.current?.dispose(); renderer.current = null; };
  }, [chart.id, chart.unsupported]);
  useLayoutEffect(() => { renderer.current?.update(chart); }, [chart]);
  const reason = chart.unsupported ?? error;
  return <div ref={host} role="img" aria-label={reason ? `${chart.title || "Excel chart"}: ${reason}` : `${chart.title || "Excel chart"}. ${chart.series.length} series, ${chart.series[0]?.values.length ?? 0} points.`}
    data-native-chart={chart.id} style={{ width: "100%", height: "100%", background: "white", border: "1px solid #d0d5dd", boxSizing: "border-box", color: "#475467" }}>
    {reason && <div style={{ padding: 20, font: "13px/1.5 Arial, sans-serif" }}><strong>{chart.title || "Excel chart"}</strong><p>{reason}</p><span>Native chart retained in the workbook.</span></div>}
  </div>;
}

/** Position chart drawings with the same geometry, zoom and scroll as cells.
 * No workbook parsing or data fetching happens in the grid's draw loop. */
export const ChartLayer = forwardRef<ChartLayerHandle, { engine: GridEngine; view: () => View | null }>(function ChartLayer({ engine, view }, ref) {
  const charts = useSyncExternalStore(engine.subscribeCharts ?? noSubscribe, engine.charts ?? empty, empty);
  const root = useRef<HTMLDivElement>(null);
  const frames = useRef(new Map<string, HTMLDivElement>());
  const current = useRef({ view, charts });
  current.current = { view, charts };
  const sync = () => {
    const v = current.current.view(), el = root.current;
    if (!v || !el) return;
    const { geometry: g, scroller: sc, scale: z } = v;
    Object.assign(el.style, { left: `${sc.offsetLeft + HEADER_W*z}px`, top: `${sc.offsetTop + HEADER_H*z}px`, width: `${Math.max(0,sc.clientWidth-HEADER_W*z)}px`, height: `${Math.max(0,sc.clientHeight-HEADER_H*z)}px` });
    for (const chart of current.current.charts) {
      const frame = frames.current.get(chart.id);
      if (!frame) continue;
      const { from, to, absolute } = chart.anchor;
      // Chart-only sheets can anchor beyond the used range; guard missing geometry.
      const x = absolute ? from.dx : (g.colLeft[from.col] ?? g.totalWidth) + from.dx;
      const y = absolute ? from.dy : (g.rowTop[from.row] ?? g.totalHeight) + from.dy;
      const w = to ? (g.colLeft[to.col] ?? g.totalWidth) + to.dx-x : chart.anchor.width;
      const h = to ? (g.rowTop[to.row] ?? g.totalHeight) + to.dy-y : chart.anchor.height;
      const frozenX = !absolute && from.col < g.frozenCols;
      const frozenY = !absolute && from.row < g.frozenRows;
      const left = x*z - (frozenX ? 0 : Math.round(sc.scrollLeft));
      const top = y*z - (frozenY ? 0 : Math.round(sc.scrollTop));
      const clipLeft = frozenX ? 0 : Math.max(0, (g.colLeft[g.frozenCols] ?? 0)*z-left);
      const clipTop = frozenY ? 0 : Math.max(0, (g.rowTop[g.frozenRows] ?? 0)*z-top);
      Object.assign(frame.style, { display: chart.sheet === v.sheet && w>0 && h>0 ? "block" : "none", left: `${left}px`, top: `${top}px`, width: `${Math.max(1,w)}px`, height: `${Math.max(1,h)}px`, transform: `scale(${z})`, transformOrigin: "top left", clipPath: `inset(${clipTop/z}px 0 0 ${clipLeft/z}px)` });
    }
  };
  useImperativeHandle(ref, () => ({ sync }));
  useLayoutEffect(sync);
  useEffect(() => {
    const el = root.current;
    if (!el) return;
    const wheel = (e: WheelEvent) => {
      const sc = current.current.view()?.scroller;
      if (!sc || e.ctrlKey || e.metaKey) return;
      e.preventDefault();
      const unit = e.deltaMode === 1 ? 20 : e.deltaMode === 2 ? sc.clientHeight : 1;
      sc.scrollLeft += e.deltaX*unit;
      sc.scrollTop += e.deltaY*unit;
    };
    el.addEventListener("wheel", wheel, { passive: false });
    return () => el.removeEventListener("wheel", wheel);
  }, []);
  const sheet = view()?.sheet;
  return <div ref={root} style={{ position: "absolute", overflow: "hidden", pointerEvents: "none" }}>
    {charts.filter(c => c.sheet === sheet).map(chart => <div key={chart.id} ref={el => { if (el) frames.current.set(chart.id, el); else frames.current.delete(chart.id); }} style={{ position: "absolute", pointerEvents: "auto" }} onMouseDown={e => e.stopPropagation()}>
      <ChartView chart={chart} />
    </div>)}
  </div>;
});
