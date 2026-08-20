export type CellAddress = {
  sheet: string;
  row: number;
  col: number;
};

export type CellMutation = {
  type: "set_cell";
  address: CellAddress;
  old_value: string | number | null;
  new_value: string | number | null;
  old_formula?: string | null;
  new_formula?: string | null;
};

export type RangeMutation = {
  type: "set_range";
  sheet: string;
  start_row: number;
  start_col: number;
  values: (string | number | null)[][];
  old_values?: (string | number | null)[][];
};

export type CellFormat = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  font_color?: string;
  /** CSS color for the cell fill, e.g. "#1F4E79". Null/empty clears it. */
  background_color?: string | null;
  font_size?: number;
  font_family?: string;
  horizontal_align?: "left" | "center" | "right";
  vertical_align?: "top" | "middle" | "bottom";
  /** Whether long text wraps within the cell. */
  wrap_text?: boolean;
  number_format?: string;
};

export type FormatMutation = {
  type: "set_format";
  sheet: string;
  /** A1 range string as the agent supplied it (for display in the diff). */
  range: string;
  /** Expanded list of cells the format applies to. */
  cells: Array<{ row: number; col: number }>;
  old_format: Array<{ row: number; col: number; format: CellFormat | null }>;
  new_format: CellFormat;
};

export type ColumnWidthMutation = {
  type: "set_column_width";
  sheet: string;
  columns: number[];
  old_widths: Array<{ col: number; width: number | null }>;
  new_width: number;
};

export type RowHeightMutation = {
  type: "set_row_height";
  sheet: string;
  rows: number[];
  old_heights: Array<{ row: number; height: number | null }>;
  new_height: number;
};

export type MergeMutation = {
  type: "merge_cells" | "unmerge_cells";
  sheet: string;
  range: string;
  start_row: number;
  start_col: number;
  end_row: number;
  end_col: number;
};

export type SheetMutation =
  | { type: "create_sheet"; name: string; tab_color?: string | null }
  | {
      type: "delete_sheet";
      name: string;
      /**
       * Full Univer sheet snapshot captured BEFORE delete so Reject can
       * restore the sheet (cells, formats, dimensions) losslessly.
       */
      sheet_snapshot?: any | null;
    }
  | { type: "rename_sheet"; old_name: string; new_name: string };

export type ClearRangeMutation = {
  type: "clear_range";
  sheet: string;
  range: string;
  cells: Array<{ row: number; col: number; old_value: any; old_formula: string | null }>;
};

/** One non-empty cell captured before a delete_rows / delete_columns. */
export type DeletedCellSnapshot = {
  row: number;
  col: number;
  value: any;
  formula: string | null;
  format: CellFormat | null;
};

export type InsertDeleteMutation =
  | { type: "insert_rows"; sheet: string; before: number; count: number }
  | {
      type: "delete_rows";
      sheet: string;
      start: number;
      count: number;
      /** Non-empty cells in the deleted band, for Reject restore. */
      deleted_cells?: DeletedCellSnapshot[];
    }
  | { type: "insert_columns"; sheet: string; before: number; count: number }
  | {
      type: "delete_columns";
      sheet: string;
      start: number;
      count: number;
      /** Non-empty cells in the deleted band, for Reject restore. */
      deleted_cells?: DeletedCellSnapshot[];
    };

export type FreezeMutation = {
  type: "freeze_panes";
  sheet: string;
  freeze_rows: number;
  freeze_cols: number;
  /** Prior freeze state so Reject can restore. */
  old_freeze_rows?: number;
  old_freeze_cols?: number;
};

export type UnfreezeMutation = {
  type: "unfreeze_panes";
  sheet: string;
  /** Prior freeze state so Reject can restore. */
  old_freeze_rows?: number;
  old_freeze_cols?: number;
};

export type NoteMutation =
  | {
      type: "set_note";
      sheet: string;
      row: number;
      col: number;
      text: string;
      /** Prior note text so Reject can restore it (null = cell had none). */
      old_text?: string | null;
    }
  | {
      type: "delete_note";
      sheet: string;
      row: number;
      col: number;
      /** Prior note text so Reject can restore it (null = cell had none). */
      old_text?: string | null;
    };

export type HideShowMutation =
  | { type: "hide_rows"; sheet: string; rows: number[] }
  | { type: "show_rows"; sheet: string; rows: number[] }
  | { type: "hide_columns"; sheet: string; columns: number[] }
  | { type: "show_columns"; sheet: string; columns: number[] };

export type DefineNameMutation = {
  type: "define_name";
  /** Workbook-scoped defined name, e.g. "Revenue". */
  name: string;
  /** Sheet-qualified A1 ref the name points at, e.g. "Model!$B$5:$B$12". */
  ref: string;
  /**
   * The name's previous ref if it already existed (so Reject can restore the
   * prior target). `null` means the name was newly created → Reject deletes it.
   */
  old_ref?: string | null;
};

export type UniverMutation =
  | CellMutation
  | RangeMutation
  | FormatMutation
  | ColumnWidthMutation
  | RowHeightMutation
  | MergeMutation
  | SheetMutation
  | ClearRangeMutation
  | InsertDeleteMutation
  | FreezeMutation
  | UnfreezeMutation
  | HideShowMutation
  | DefineNameMutation
  | NoteMutation;

export type BatchStatus = "streaming" | "pending" | "accepted" | "rejected";

export type ChangeBatch = {
  id: string;
  prompt: string;
  justification: string;
  /**
   * Structured "memory" record the agent attaches to `done` — what it did,
   * key decisions, and key figures learned (from fetched sources etc.).
   * Written specifically to survive into future turns' prior-context digest,
   * unlike agent_text which gets truncated.
   */
  turn_summary?: string;
  mutations: UniverMutation[];
  status: BatchStatus;
  created_at: string;
  /** Streaming prose the agent produced for this batch (preserved on done). */
  agent_text?: string;
  /**
   * The model's reasoning/plan for this batch — Claude's extended-thinking
   * block streamed live, or Codex's reasoning summary. Rendered as a
   * collapsible "Plan" block above the prose. Preserved on done.
   */
  reasoning?: string;
  /** URLs the agent fetched mid-turn via fetch_web. Shown as inline chips. */
  fetched_urls?: string[];
  /**
   * The batch's effects are already accounted for by the file on disk and
   * the save baseline, so it must contribute NOTHING to future save
   * mirrors (replaying its structure ops would apply row inserts / sheet
   * deletes a second time). Set in two places: on session resume (the grid
   * was just loaded from the file), and after a successful full-export
   * save (which bakes every batch into the file and resets the baseline).
   */
  persisted?: boolean;
};

export type Workbook = {
  path: string;
  filename: string;
  dirty: boolean;
};
