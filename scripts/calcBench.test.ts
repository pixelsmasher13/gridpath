/**
 * Opt-in calc-engine benchmark (skipped in normal `npm test` runs).
 *
 *   CALC_BENCH=1 CALC_BENCH_FILE=/path/to/model.xlsx npx vitest run scripts/calcBench.test.ts
 *   CALC_BENCH=1 CALC_BENCH_PRO=1 CALC_BENCH_FILE=... npx vitest run scripts/calcBench.test.ts
 */
import { describe, it } from "vitest";

const enabled = process.env.CALC_BENCH === "1";

describe.skipIf(!enabled)("calc engine benchmark", () => {
  it("measures initial calc and single-edit recalc", { timeout: 900_000 }, async () => {
    const { runBench } = await import("./calcBenchImpl");
    await runBench(process.env.CALC_BENCH_FILE!, process.env.CALC_BENCH_PRO === "1");
  });
});
