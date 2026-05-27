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
  | { type: "delete_sheet"; name: string }
  | { type: "rename_sheet"; old_name: string; new_name: string };

export type ClearRangeMutation = {
  type: "clear_range";
  sheet: string;
  range: string;
  cells: Array<{ row: number; col: number; old_value: any; old_formula: string | null }>;
};

export type InsertDeleteMutation =
  | { type: "insert_rows"; sheet: string; before: number; count: number }
  | { type: "delete_rows"; sheet: string; start: number; count: number }
  | { type: "insert_columns"; sheet: string; before: number; count: number }
  | { type: "delete_columns"; sheet: string; start: number; count: number };

export type FreezeMutation = {
  type: "freeze_panes";
  sheet: string;
  freeze_rows: number;
  freeze_cols: number;
};

export type UnfreezeMutation = {
  type: "unfreeze_panes";
  sheet: string;
};

export type HideShowMutation =
  | { type: "hide_rows"; sheet: string; rows: number[] }
  | { type: "show_rows"; sheet: string; rows: number[] }
  | { type: "hide_columns"; sheet: string; columns: number[] }
  | { type: "show_columns"; sheet: string; columns: number[] };

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
  | HideShowMutation;

export type BatchStatus = "streaming" | "pending" | "accepted" | "rejected";

export type ChangeBatch = {
  id: string;
  prompt: string;
  justification: string;
  mutations: UniverMutation[];
  status: BatchStatus;
  created_at: string;
  /** Streaming prose the agent produced for this batch (preserved on done). */
  agent_text?: string;
  /** URLs the agent fetched mid-turn via fetch_web. Shown as inline chips. */
  fetched_urls?: string[];
};

export type Workbook = {
  path: string;
  filename: string;
  dirty: boolean;
};
