/**
 * @uekichinos/quire — a small, dependency-light `.xlsx` writer.
 *
 * Phases 1–2: worksheets; string / number / boolean / `Date` / formula cells;
 * a de-duplicated style pool (font, fill, border, alignment, number format);
 * row- and cell-level styling; `setCell` by A1 ref. Merges, column widths and
 * freeze panes land in Phase 3 (see PLAN.md); a reader comes later.
 */

export { createWorkbook } from './workbook'
export type {
  Workbook,
  Worksheet,
  WorksheetOptions,
  CellInput,
  CellScalar,
  FormulaValue,
  AddRowOptions,
  RowOptions,
  ColumnSpec,
  FreezeOptions,
} from './workbook'
export type {
  CellStyle,
  FontStyle,
  AlignStyle,
  BorderEdge,
  BorderStyle,
  BorderStyleInput,
} from './style'

// --- Reading (0.2.0) ---
export { readWorkbook, readWorkbookAsync } from './read'
export type {
  ReadWorkbook,
  ReadWorksheet,
  ReadCell,
  ReadCellType,
  ReadOptions,
  ReadLimits,
  ValuesOptions,
  ReadStyle,
  ReadFont,
  ReadBorderEdge,
} from './read'
export { QuireError } from './errors'
export { XlsxReadError } from './unzip'
export { XmlError } from './xml-read'

// Low-level helpers, exported for advanced use and testing.
export { colLetter, colNumber, parseRef, toRef, toRange, type CellAddress } from './address'
export { dateToSerial } from './datetime'
export { escapeXml } from './xml'
