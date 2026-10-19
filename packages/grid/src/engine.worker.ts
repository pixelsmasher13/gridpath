/// <reference lib="webworker" />
/**
 * The engine worker: one workbook per worker, commands answered in arrival
 * order. Recalculation never blocks the page that draws the grid.
 */
import { EngineHost, type HostRequest, type HostResponse } from "./engineHost";

const host = new EngineHost();
const scope = self as unknown as DedicatedWorkerGlobalScope;

// Requests run one at a time: each command sees the state every earlier
// command left, the same ordering the desktop's per-tab queue gives.
let chain: Promise<void> = Promise.resolve();
scope.onmessage = (ev: MessageEvent<HostRequest>) => {
  const { id, cmd, args } = ev.data;
  chain = chain.then(async () => {
    let reply: HostResponse;
    try {
      reply = { id, ok: true, result: await host.handle(cmd, args) };
    } catch (e) {
      reply = { id, ok: false, error: e instanceof Error ? e.message : String(e) };
    }
    const r = reply.ok ? reply.result : null;
    scope.postMessage(reply, r instanceof Uint8Array ? [r.buffer] : []);
  });
};
