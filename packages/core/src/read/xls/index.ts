/**
 * Legacy Excel 97-2003 (`.xls`, BIFF8) reader: `readXlsWorkbook` turns the
 * OLE2 compound file into the same `Workbook` model the `.xlsx` reader
 * builds, so the viewer, the calc engine and `writeWorkbook` work on it
 * unchanged. Encrypted files (FILEPASS) and Excel 5.0/95 files are rejected
 * with `XlsxPasswordError` / `XlsxFormatError`.
 */
export { readXlsWorkbook, type ReadXlsOptions } from './read-xls';
