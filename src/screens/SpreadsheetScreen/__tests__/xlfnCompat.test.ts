import { describe, expect, it } from "vitest";
import { addXlfnPrefixes, stripXlfnPrefixes } from "../xlfnCompat";

describe("stripXlfnPrefixes", () => {
  it("strips _xlfn. from post-2007 function calls", () => {
    expect(stripXlfnPrefixes('_xlfn.IFS(A1="x",B1,TRUE,C1)')).toBe('IFS(A1="x",B1,TRUE,C1)');
    expect(stripXlfnPrefixes("+_xlfn.IFS(A1,B1)+_xlfn.XLOOKUP(C1,D:D,E:E)")).toBe(
      "+IFS(A1,B1)+XLOOKUP(C1,D:D,E:E)",
    );
  });

  it("strips the combined _xlfn._xlws. prefix", () => {
    expect(stripXlfnPrefixes("_xlfn._xlws.FILTER(A:A,B:B=1)")).toBe("FILTER(A:A,B:B=1)");
  });

  it("leaves string literals and _xll./_xlpm. tokens alone", () => {
    expect(stripXlfnPrefixes('CONCATENATE("_xlfn.IFS(",A1)')).toBe('CONCATENATE("_xlfn.IFS(",A1)');
    expect(stripXlfnPrefixes("_xll.FDS(A1)")).toBe("_xll.FDS(A1)");
    expect(stripXlfnPrefixes("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A1)")).toBe(
      "LAMBDA(_xlpm.x,_xlpm.x*2)(A1)",
    );
  });

  it("is a no-op on formulas without the prefix", () => {
    const f = "SUM(A1:A10)*IF(B1>0,1,-1)";
    expect(stripXlfnPrefixes(f)).toBe(f);
  });
});

describe("addXlfnPrefixes", () => {
  it("prefixes known post-2007 functions", () => {
    expect(addXlfnPrefixes("IFS(A1,B1)")).toBe("_xlfn.IFS(A1,B1)");
    expect(addXlfnPrefixes("SUM(A1)+MAXIFS(B:B,C:C,1)")).toBe("SUM(A1)+_xlfn.MAXIFS(B:B,C:C,1)");
  });

  it("uses the _xlfn._xlws. prefix for FILTER/SORT", () => {
    expect(addXlfnPrefixes("SORT(FILTER(A:A,B:B=1))")).toBe(
      "_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A:A,B:B=1))",
    );
  });

  it("does not double-prefix or touch classic functions, refs and strings", () => {
    expect(addXlfnPrefixes("_xlfn.IFS(A1,B1)")).toBe("_xlfn.IFS(A1,B1)");
    expect(addXlfnPrefixes("SUM(A1:A10)")).toBe("SUM(A1:A10)");
    expect(addXlfnPrefixes('IF(A1="IFS(",B1,IFS_rate(C1))')).toBe('IF(A1="IFS(",B1,IFS_rate(C1))');
  });

  it("round-trips with stripXlfnPrefixes", () => {
    const stored = '+_xlfn.IFS(\'Scenarios Inputs\'!$P$3="Large",C$6,TRUE,_xlfn._xlws.SORT(D:D))';
    expect(addXlfnPrefixes(stripXlfnPrefixes(stored))).toBe(stored);
  });
});
