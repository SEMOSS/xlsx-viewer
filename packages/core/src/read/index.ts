/**
 * Load pipeline: `.xlsx` / `.xlsm` / `.xltx` bytes -> `Workbook`.
 *
 * `readWorkbook` opens the OPC package, follows the relationships from the
 * package root to the workbook and from there to every part (theme, styles,
 * shared strings, sheets, tables, comments, drawings, charts, document
 * properties). Problems in individual parts are recorded in
 * `workbook.warnings` instead of failing the whole load.
 */
export type { ReadOptions } from './context';
export { XlsxFormatError, XlsxPasswordError } from './errors';
export { readWorkbook } from './read-workbook';
