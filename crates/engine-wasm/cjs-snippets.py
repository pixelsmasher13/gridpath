#!/usr/bin/env python3
"""wasm-bindgen emits inline JS snippets as ESM even for --target nodejs,
where the glue `require()`s them. Rewrite each snippet to CommonJS."""
import pathlib, re, sys
root = pathlib.Path(sys.argv[1]) / "snippets"
for f in root.rglob("*.js"):
    s = f.read_text()
    names = re.findall(r"^\s*export\s+function\s+(\w+)", s, re.M)
    if not names:
        continue
    s = re.sub(r"^(\s*)export\s+function", r"\1function", s, flags=re.M)
    s += "\nmodule.exports = { " + ", ".join(names) + " };\n"
    f.write_text(s)
    print(f"cjs: {f.relative_to(root.parent)} ({', '.join(names)})")
