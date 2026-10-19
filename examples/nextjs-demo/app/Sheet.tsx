"use client";
// The spreadsheet runs in the browser (its engine is a Web Worker), so it
// lives in a client component.
import { Spreadsheet } from "gridpath";

export function Sheet() {
  return (
    <Spreadsheet
      sheets={["Sheet1"]}
      onReady={(sheet) => {
        (window as any).__demo = { sheet, ready: true };
      }}
      onError={(err) => {
        (window as any).__demo = { error: err.message, ready: true };
      }}
    />
  );
}
