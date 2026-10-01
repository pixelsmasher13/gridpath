import React, { useEffect, useRef, useState } from "react";
import styled from "styled-components";
import {
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Baseline,
  PaintBucket,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  WrapText,
  Merge,
  TableProperties,
  Hash,
  Grid2x2,
  ChevronDown,
  DollarSign,
  Percent,
  AArrowUp,
  AArrowDown,
  ArrowDownAZ,
  ArrowDownZA,
  Undo2,
  Redo2,
  Paintbrush,
  Eraser,
  Pin,
  Filter,
  Search,
  Palette,
  ListChecks,
} from "lucide-react";

/**
 * Excel-style "Home" formatting toolbar. Presentational only — every control
 * calls back to SpreadsheetScreen, which reads the live Univer selection and
 * applies the change through the UniverGrid handle (same path the agent uses).
 * The toolbar owns just its own popover open/close state.
 */
export type InsertAction =
  | "rowAbove"
  | "rowBelow"
  | "colLeft"
  | "colRight"
  | "deleteRows"
  | "deleteColumns";

export type FormatToolbarProps = {
  /** No active workbook tab — controls are inert. */
  disabled: boolean;
  /** Undo / redo the most recent action (manual edit or agent batch). */
  onUndo: () => void;
  onRedo: () => void;
  /** Arm/disarm Format Painter (copies the anchor cell's format). */
  onFormatPainter: () => void;
  /** Whether Format Painter is currently armed (button shows active). */
  painterActive: boolean;
  /** Clear the selection's contents, formats, or both. */
  onClear: (kind: "all" | "formats" | "contents") => void;
  /** Freeze panes: top row, first column, at selection, or unfreeze. */
  onFreeze: (kind: "topRow" | "firstCol" | "atSelection" | "none") => void;
  /** Toggle an AutoFilter over the selection. */
  onToggleFilter: () => void;
  /** Open Univer's conditional-formatting panel for the active sheet. */
  onConditionalFormatting: () => void;
  /** Open Univer's data-validation panel. */
  onDataValidation: () => void;
  /** Open the Find & Replace dialog. */
  onFindReplace: () => void;
  /**
   * Max outline (grouping) depth of the active sheet's columns/rows, or null
   * when the sheet has no groups (buttons hidden). Mirrors Excel's "1 2 3"
   * outline level buttons.
   */
  outline: { cols: number; rows: number } | null;
  /** Collapse/expand groups to the given depth (maxDepth+1 = show all). */
  onOutlineLevel: (axis: "cols" | "rows", level: number) => void;
  /** Toggle a boolean font attribute across the selection (flips anchor state). */
  onToggle: (key: "bold" | "italic" | "underline" | "strike") => void;
  onFontColor: (color: string) => void;
  /** Pass null to clear the fill ("No fill"). */
  onFillColor: (color: string | null) => void;
  onAlign: (a: "left" | "center" | "right") => void;
  onVerticalAlign: (a: "top" | "middle" | "bottom") => void;
  /** Toggle wrap-text across the selection. */
  onWrap: () => void;
  onFontFamily: (family: string) => void;
  onFontSize: (size: number) => void;
  /** Step the font size up (+1) or down (-1) one rung on the Excel ladder. */
  onAdjustFontSize: (delta: 1 | -1) => void;
  /** Sort the selection's rows by its left column, ascending or descending. */
  onSort: (direction: "asc" | "desc") => void;
  /** Merge & center, plain merge, or unmerge the selection. */
  onMerge: (kind: "center" | "merge" | "unmerge") => void;
  /** Insert/delete rows or columns relative to the selection. */
  onInsert: (action: InsertAction) => void;
  /** Excel number-format pattern, e.g. "$#,##0.00". "General" clears it. */
  onNumberFormat: (pattern: string) => void;
  /** Bump the selection's number format by ±1 decimal place. */
  onAdjustDecimals: (delta: 1 | -1) => void;
  onBorders: (kind: "all" | "outer" | "none") => void;
};

// A compact, Excel-ish palette: greyscale row + standard colors row.
const SWATCHES: string[] = [
  "#000000", "#404040", "#808080", "#BFBFBF", "#FFFFFF",
  "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050",
  "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0",
];

const NUMBER_FORMATS: Array<{ label: string; pattern: string; hint: string }> = [
  { label: "General", pattern: "General", hint: "" },
  { label: "Number", pattern: "#,##0.00", hint: "1,234.56" },
  { label: "Currency", pattern: "$#,##0.00", hint: "$1,234.56" },
  { label: "Accounting", pattern: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', hint: "$  1,234.56" },
  { label: "Percent", pattern: "0.0%", hint: "12.3%" },
  { label: "Comma", pattern: "#,##0", hint: "1,234" },
  { label: "Scientific", pattern: "0.00E+00", hint: "1.23E+03" },
  { label: "Fraction", pattern: "# ?/?", hint: "1 1/2" },
  { label: "Short date", pattern: "mm/dd/yyyy", hint: "06/03/2026" },
  { label: "Long date", pattern: 'dddd, mmmm d, yyyy', hint: "Wed, June 3, 2026" },
  { label: "Time", pattern: "h:mm:ss AM/PM", hint: "1:30:00 PM" },
  { label: "Text", pattern: "@", hint: "as typed" },
];

const FONT_FAMILIES: string[] = [
  "Calibri", "Arial", "Aptos Narrow", "Times New Roman",
  "Georgia", "Verdana", "Courier New", "Consolas",
];

const FONT_SIZES: number[] = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

/** Excel-style outline buttons: 1..maxDepth+1, where the last shows all. */
function levelButtons(maxDepth: number): number[] {
  return Array.from({ length: Math.min(maxDepth, 7) + 1 }, (_, i) => i + 1);
}

const OutlineLabel = styled.span`
  font-size: 10px;
  opacity: 0.6;
  margin: 0 2px 0 4px;
  user-select: none;
`;

const OutlineDigit = styled.span`
  font-size: 11px;
  font-weight: 600;
  min-width: 10px;
  text-align: center;
`;

export const FormatToolbar: React.FC<FormatToolbarProps> = ({
  disabled,
  onUndo,
  onRedo,
  onFormatPainter,
  painterActive,
  onClear,
  onFreeze,
  onToggleFilter,
  onConditionalFormatting,
  onDataValidation,
  onFindReplace,
  outline,
  onOutlineLevel,
  onToggle,
  onFontColor,
  onFillColor,
  onAlign,
  onVerticalAlign,
  onWrap,
  onFontFamily,
  onFontSize,
  onAdjustFontSize,
  onSort,
  onMerge,
  onInsert,
  onNumberFormat,
  onAdjustDecimals,
  onBorders,
}) => {
  // Which popover (if any) is open. Only one at a time.
  const [open, setOpen] = useState<
    null | "font" | "fill" | "number" | "borders" | "merge" | "insert" | "sort" | "clear" | "freeze"
  >(null);
  const rootRef = useRef<HTMLDivElement>(null);

  // Close on outside-click or Escape.
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(null);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(null);
    };
    document.addEventListener("mousedown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  const toggle = (which: NonNullable<typeof open>) =>
    setOpen((cur) => (cur === which ? null : which));

  return (
    <Bar ref={rootRef} aria-disabled={disabled}>
      <Group>
        <IconBtn title="Undo (⌘Z)" disabled={disabled} onClick={() => onUndo()}>
          <Undo2 size={14} />
        </IconBtn>
        <IconBtn title="Redo (⌘⇧Z)" disabled={disabled} onClick={() => onRedo()}>
          <Redo2 size={14} />
        </IconBtn>
        <IconBtn
          title="Format painter"
          disabled={disabled}
          aria-pressed={painterActive}
          $active={painterActive}
          onClick={() => onFormatPainter()}
        >
          <Paintbrush size={14} />
        </IconBtn>
        <PopoverHost>
          <IconBtn title="Clear" disabled={disabled} onClick={() => toggle("clear")}>
            <Eraser size={14} />
            <ChevronDown size={10} style={{ marginLeft: -2, opacity: 0.7 }} />
          </IconBtn>
          {open === "clear" && (
            <Menu>
              <MenuItem onClick={() => { onClear("all"); setOpen(null); }}>Clear all</MenuItem>
              <MenuItem onClick={() => { onClear("formats"); setOpen(null); }}>Clear formats</MenuItem>
              <MenuItem onClick={() => { onClear("contents"); setOpen(null); }}>Clear contents</MenuItem>
            </Menu>
          )}
        </PopoverHost>
      </Group>

      <Divider />

      <Group>
        <Select
          title="Font"
          disabled={disabled}
          defaultValue=""
          style={{ width: 116 }}
          onChange={(e) => {
            if (e.target.value) onFontFamily(e.target.value);
            e.currentTarget.selectedIndex = 0;
          }}
        >
          <option value="" disabled hidden>Font</option>
          {FONT_FAMILIES.map((f) => (
            <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>
          ))}
        </Select>
        <Select
          title="Font size"
          disabled={disabled}
          defaultValue=""
          style={{ width: 56 }}
          onChange={(e) => {
            const n = Number(e.target.value);
            if (Number.isFinite(n) && n > 0) onFontSize(n);
            e.currentTarget.selectedIndex = 0;
          }}
        >
          <option value="" disabled hidden>Size</option>
          {FONT_SIZES.map((s) => (
            <option key={s} value={s}>{s}</option>
          ))}
        </Select>
        <IconBtn title="Increase font size" disabled={disabled} onClick={() => onAdjustFontSize(1)}>
          <AArrowUp size={15} />
        </IconBtn>
        <IconBtn title="Decrease font size" disabled={disabled} onClick={() => onAdjustFontSize(-1)}>
          <AArrowDown size={15} />
        </IconBtn>
      </Group>

      <Divider />

      <Group>
        <IconBtn title="Bold (⌘B)" disabled={disabled} onClick={() => onToggle("bold")}>
          <Bold size={14} />
        </IconBtn>
        <IconBtn title="Italic (⌘I)" disabled={disabled} onClick={() => onToggle("italic")}>
          <Italic size={14} />
        </IconBtn>
        <IconBtn title="Underline (⌘U)" disabled={disabled} onClick={() => onToggle("underline")}>
          <Underline size={14} />
        </IconBtn>
        <IconBtn title="Strikethrough" disabled={disabled} onClick={() => onToggle("strike")}>
          <Strikethrough size={14} />
        </IconBtn>
      </Group>

      <Divider />

      <Group>
        <PopoverHost>
          <IconBtn title="Font color" disabled={disabled} onClick={() => toggle("font")}>
            <Baseline size={14} />
            <ChevronDown size={10} style={{ marginLeft: -2, opacity: 0.7 }} />
          </IconBtn>
          {open === "font" && (
            <SwatchPopover
              onPick={(c) => {
                if (c !== null) onFontColor(c);
                setOpen(null);
              }}
            />
          )}
        </PopoverHost>

        <PopoverHost>
          <IconBtn title="Fill color" disabled={disabled} onClick={() => toggle("fill")}>
            <PaintBucket size={14} />
            <ChevronDown size={10} style={{ marginLeft: -2, opacity: 0.7 }} />
          </IconBtn>
          {open === "fill" && (
            <SwatchPopover
              showNoFill
              onPick={(c) => {
                onFillColor(c);
                setOpen(null);
              }}
            />
          )}
        </PopoverHost>
      </Group>

      <Divider />

      <Group>
        <IconBtn title="Align top" disabled={disabled} onClick={() => onVerticalAlign("top")}>
          <AlignVerticalJustifyStart size={14} />
        </IconBtn>
        <IconBtn title="Align middle" disabled={disabled} onClick={() => onVerticalAlign("middle")}>
          <AlignVerticalJustifyCenter size={14} />
        </IconBtn>
        <IconBtn title="Align bottom" disabled={disabled} onClick={() => onVerticalAlign("bottom")}>
          <AlignVerticalJustifyEnd size={14} />
        </IconBtn>
      </Group>

      <Divider />

      <Group>
        <IconBtn title="Align left" disabled={disabled} onClick={() => onAlign("left")}>
          <AlignLeft size={14} />
        </IconBtn>
        <IconBtn title="Align center" disabled={disabled} onClick={() => onAlign("center")}>
          <AlignCenter size={14} />
        </IconBtn>
        <IconBtn title="Align right" disabled={disabled} onClick={() => onAlign("right")}>
          <AlignRight size={14} />
        </IconBtn>
        <IconBtn title="Wrap text" disabled={disabled} onClick={() => onWrap()}>
          <WrapText size={14} />
        </IconBtn>
      </Group>

      <Divider />

      <Group>
        <PopoverHost>
          <IconBtn title="Merge cells" disabled={disabled} onClick={() => toggle("merge")}>
            <Merge size={14} />
            <ChevronDown size={10} style={{ marginLeft: -2, opacity: 0.7 }} />
          </IconBtn>
          {open === "merge" && (
            <Menu>
              <MenuItem onClick={() => { onMerge("center"); setOpen(null); }}>Merge &amp; center</MenuItem>
              <MenuItem onClick={() => { onMerge("merge"); setOpen(null); }}>Merge cells</MenuItem>
              <MenuItem onClick={() => { onMerge("unmerge"); setOpen(null); }}>Unmerge</MenuItem>
            </Menu>
          )}
        </PopoverHost>

        <PopoverHost>
          <TextBtn title="Insert / delete rows &amp; columns" disabled={disabled} onClick={() => toggle("insert")}>
            <TableProperties size={13} /> Insert <ChevronDown size={11} style={{ opacity: 0.7 }} />
          </TextBtn>
          {open === "insert" && (
            <Menu style={{ minWidth: 184 }}>
              <MenuItem onClick={() => { onInsert("rowAbove"); setOpen(null); }}>Insert row above</MenuItem>
              <MenuItem onClick={() => { onInsert("rowBelow"); setOpen(null); }}>Insert row below</MenuItem>
              <MenuItem onClick={() => { onInsert("colLeft"); setOpen(null); }}>Insert column left</MenuItem>
              <MenuItem onClick={() => { onInsert("colRight"); setOpen(null); }}>Insert column right</MenuItem>
              <MenuSep />
              <MenuItem onClick={() => { onInsert("deleteRows"); setOpen(null); }}>Delete row(s)</MenuItem>
              <MenuItem onClick={() => { onInsert("deleteColumns"); setOpen(null); }}>Delete column(s)</MenuItem>
            </Menu>
          )}
        </PopoverHost>

        <PopoverHost>
          <TextBtn title="Sort selection" disabled={disabled} onClick={() => toggle("sort")}>
            <ArrowDownAZ size={13} /> Sort <ChevronDown size={11} style={{ opacity: 0.7 }} />
          </TextBtn>
          {open === "sort" && (
            <Menu style={{ minWidth: 168 }}>
              <MenuItem onClick={() => { onSort("asc"); setOpen(null); }}>
                <span style={{ display: "inline-flex", alignItems: "center", gap: 8 }}>
                  <ArrowDownAZ size={14} /> Sort A → Z
                </span>
              </MenuItem>
              <MenuItem onClick={() => { onSort("desc"); setOpen(null); }}>
                <span style={{ display: "inline-flex", alignItems: "center", gap: 8 }}>
                  <ArrowDownZA size={14} /> Sort Z → A
                </span>
              </MenuItem>
            </Menu>
          )}
        </PopoverHost>

        <IconBtn title="Toggle filter" disabled={disabled} onClick={() => onToggleFilter()}>
          <Filter size={14} />
        </IconBtn>

        <IconBtn title="Conditional formatting" disabled={disabled} onClick={() => onConditionalFormatting()}>
          <Palette size={14} />
        </IconBtn>

        <IconBtn title="Data validation" disabled={disabled} onClick={() => onDataValidation()}>
          <ListChecks size={14} />
        </IconBtn>

        <PopoverHost>
          <TextBtn title="Freeze panes" disabled={disabled} onClick={() => toggle("freeze")}>
            <Pin size={13} /> Freeze <ChevronDown size={11} style={{ opacity: 0.7 }} />
          </TextBtn>
          {open === "freeze" && (
            <Menu style={{ minWidth: 184 }}>
              <MenuItem onClick={() => { onFreeze("atSelection"); setOpen(null); }}>Freeze panes (at selection)</MenuItem>
              <MenuItem onClick={() => { onFreeze("topRow"); setOpen(null); }}>Freeze top row</MenuItem>
              <MenuItem onClick={() => { onFreeze("firstCol"); setOpen(null); }}>Freeze first column</MenuItem>
              <MenuSep />
              <MenuItem onClick={() => { onFreeze("none"); setOpen(null); }}>Unfreeze panes</MenuItem>
            </Menu>
          )}
        </PopoverHost>
      </Group>

      <Divider />

      <Group>
        <IconBtn title="Currency ($#,##0.00)" disabled={disabled} onClick={() => onNumberFormat("$#,##0.00")}>
          <DollarSign size={14} />
        </IconBtn>
        <IconBtn title="Percent (0%)" disabled={disabled} onClick={() => onNumberFormat("0%")}>
          <Percent size={14} />
        </IconBtn>
        <IconBtn title="Comma style (#,##0.00)" disabled={disabled} onClick={() => onNumberFormat("#,##0.00")}>
          <Glyph>,</Glyph>
        </IconBtn>
        <IconBtn title="Increase decimal places" disabled={disabled} onClick={() => onAdjustDecimals(1)}>
          <Glyph>+.0</Glyph>
        </IconBtn>
        <IconBtn title="Decrease decimal places" disabled={disabled} onClick={() => onAdjustDecimals(-1)}>
          <Glyph>−.0</Glyph>
        </IconBtn>

        <PopoverHost>
          <TextBtn title="Number format" disabled={disabled} onClick={() => toggle("number")}>
            <Hash size={13} /> Format <ChevronDown size={11} style={{ opacity: 0.7 }} />
          </TextBtn>
          {open === "number" && (
            <Menu>
              {NUMBER_FORMATS.map((nf) => (
                <MenuItem
                  key={nf.label}
                  onClick={() => {
                    onNumberFormat(nf.pattern);
                    setOpen(null);
                  }}
                >
                  <span>{nf.label}</span>
                  <MenuHint>{nf.hint}</MenuHint>
                </MenuItem>
              ))}
            </Menu>
          )}
        </PopoverHost>

        <PopoverHost>
          <TextBtn title="Borders" disabled={disabled} onClick={() => toggle("borders")}>
            <Grid2x2 size={13} /> Borders <ChevronDown size={11} style={{ opacity: 0.7 }} />
          </TextBtn>
          {open === "borders" && (
            <Menu>
              <MenuItem onClick={() => { onBorders("all"); setOpen(null); }}>All borders</MenuItem>
              <MenuItem onClick={() => { onBorders("outer"); setOpen(null); }}>Outer border</MenuItem>
              <MenuItem onClick={() => { onBorders("none"); setOpen(null); }}>No border</MenuItem>
            </Menu>
          )}
        </PopoverHost>
      </Group>

      <Divider />

      <Group>
        <TextBtn title="Find & Replace (⌘F)" disabled={disabled} onClick={() => onFindReplace()}>
          <Search size={13} /> Find
        </TextBtn>
      </Group>

      {outline && (outline.cols > 0 || outline.rows > 0) && (
        <>
          <Divider />
          <Group>
            {outline.cols > 0 && (
              <>
                <OutlineLabel>Cols</OutlineLabel>
                {levelButtons(outline.cols).map((n) => (
                  <IconBtn
                    key={`c${n}`}
                    title={n > outline.cols ? "Show all column groups" : `Collapse column groups to level ${n}`}
                    disabled={disabled}
                    onClick={() => onOutlineLevel("cols", n)}
                  >
                    <OutlineDigit>{n}</OutlineDigit>
                  </IconBtn>
                ))}
              </>
            )}
            {outline.rows > 0 && (
              <>
                <OutlineLabel>Rows</OutlineLabel>
                {levelButtons(outline.rows).map((n) => (
                  <IconBtn
                    key={`r${n}`}
                    title={n > outline.rows ? "Show all row groups" : `Collapse row groups to level ${n}`}
                    disabled={disabled}
                    onClick={() => onOutlineLevel("rows", n)}
                  >
                    <OutlineDigit>{n}</OutlineDigit>
                  </IconBtn>
                ))}
              </>
            )}
          </Group>
        </>
      )}
    </Bar>
  );
};

/** Swatch grid + native custom color input (+ optional "No fill"). */
const SwatchPopover: React.FC<{
  onPick: (color: string | null) => void;
  showNoFill?: boolean;
}> = ({ onPick, showNoFill }) => {
  return (
    <Menu style={{ width: 168, padding: 8 }}>
      <SwatchGrid>
        {SWATCHES.map((c) => (
          <Swatch
            key={c}
            $color={c}
            title={c}
            onClick={() => onPick(c)}
          />
        ))}
      </SwatchGrid>
      <CustomRow>
        <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
          <input
            type="color"
            defaultValue="#000000"
            onChange={(e) => onPick(e.target.value)}
            style={{ width: 22, height: 22, padding: 0, border: "none", background: "transparent", cursor: "pointer" }}
          />
          Custom…
        </label>
        {showNoFill && (
          <MenuItem style={{ padding: "4px 6px" }} onClick={() => onPick(null)}>
            No fill
          </MenuItem>
        )}
      </CustomRow>
    </Menu>
  );
};

// ---------------------------------------------------------------------------
// Styled — light "Excel ribbon" theme. The ribbon is white to blend into the
// grid below it; the dark header/status bar above and below frame it like
// Excel's title and status chrome. Accent: #2563B0 (Excel-ish blue).
// ---------------------------------------------------------------------------

const Bar = styled.div`
  display: flex;
  align-items: center;
  /* Wrap onto extra rows when the window is too narrow to fit every control on
     one line, instead of forcing the sheet column wider (which used to squeeze
     the chat panel into a thin strip). The grid ribbon row is minmax(34px,auto)
     so it grows to match. */
  flex-wrap: wrap;
  gap: 4px 4px;
  min-height: 34px;
  padding: 4px 10px;
  background: #f7f8fa;
  color: #333537;
  font-size: 12px;
  user-select: none;
`;

const Group = styled.div`
  display: flex;
  align-items: center;
  gap: 2px;
`;

const Divider = styled.div`
  width: 1px;
  align-self: stretch;
  margin: 6px 6px;
  background: #dcdee1;
`;

const IconBtn = styled.button<{ $active?: boolean }>`
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: ${(p) => (p.$active ? "#2563B0" : "transparent")};
  border: 1px solid ${(p) => (p.$active ? "#2563B0" : "transparent")};
  color: ${(p) => (p.$active ? "#fff" : "#3a3d40")};
  border-radius: 5px;
  height: 26px;
  min-width: 28px;
  padding: 0 5px;
  cursor: pointer;
  &:hover:not(:disabled) { background: ${(p) => (p.$active ? "#2563B0" : "#e7ebf0")}; border-color: ${(p) => (p.$active ? "#2563B0" : "#d2d6db")}; }
  &:disabled { opacity: 0.4; cursor: not-allowed; }
`;

const TextBtn = styled(IconBtn)`
  gap: 5px;
  padding: 0 9px;
  font-size: 12px;
`;

const PopoverHost = styled.div`
  position: relative;
  display: inline-flex;
`;

const Menu = styled.div`
  position: absolute;
  top: 30px;
  left: 0;
  z-index: 50;
  min-width: 150px;
  background: #ffffff;
  border: 1px solid #d2d6db;
  border-radius: 6px;
  box-shadow: 0 6px 18px rgba(0, 0, 0, 0.16);
  padding: 4px;
  display: flex;
  flex-direction: column;
`;

const MenuItem = styled.button`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  background: transparent;
  border: 0;
  color: #333537;
  text-align: left;
  border-radius: 4px;
  padding: 6px 8px;
  font-size: 12px;
  cursor: pointer;
  &:hover { background: #eef2f7; }
`;

const MenuHint = styled.span`
  color: #9aa0a6;
  font-size: 11px;
`;

const MenuSep = styled.div`
  height: 1px;
  margin: 4px 2px;
  background: #e6e8eb;
`;

/** Compact text glyph for icon buttons that don't have a lucide icon. */
const Glyph = styled.span`
  font-size: 11px;
  font-weight: 700;
  line-height: 1;
  letter-spacing: -0.3px;
`;

const Select = styled.select`
  background: #ffffff;
  border: 1px solid #cdd1d6;
  color: #333537;
  border-radius: 5px;
  height: 26px;
  padding: 0 4px;
  font-size: 12px;
  cursor: pointer;
  outline: none;
  &:hover:not(:disabled) { border-color: #aab0b7; }
  &:focus { border-color: #2563B0; }
  &:disabled { opacity: 0.45; cursor: not-allowed; }
  option { background: #ffffff; color: #333537; }
`;

const SwatchGrid = styled.div`
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 4px;
`;

const Swatch = styled.button<{ $color: string }>`
  width: 24px;
  height: 20px;
  border-radius: 3px;
  border: 1px solid #c8ccd1;
  background: ${(p) => p.$color};
  cursor: pointer;
  padding: 0;
  &:hover { outline: 2px solid #2563B0; outline-offset: 0; }
`;

const CustomRow = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin-top: 8px;
  padding-top: 8px;
  border-top: 1px solid #e6e8eb;
  font-size: 12px;
  color: #4a4d50;
`;
