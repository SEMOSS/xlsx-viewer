import type { Workbook } from 'xlsx-viewer-core';
import {
	sheetToCsv,
	sheetToHtml,
	workbookToJson,
	workbookToMarkdown,
	writeWorkbook,
} from 'xlsx-viewer-core';

import { printWorkbook } from './components/print/print-document';

export { printWorkbook, type PrintWorkbookOptions } from './components/print/print-document';

export type DownloadFormat = 'xlsx' | 'csv' | 'markdown' | 'json' | 'html';

const MIME: Record<DownloadFormat, string> = {
	xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	csv: 'text/csv;charset=utf-8',
	markdown: 'text/markdown;charset=utf-8',
	json: 'application/json;charset=utf-8',
	html: 'text/html;charset=utf-8',
};

const EXT: Record<DownloadFormat, string> = {
	xlsx: 'xlsx',
	csv: 'csv',
	markdown: 'md',
	json: 'json',
	html: 'html',
};

export function baseName(fileName: string): string {
	return fileName.replace(/\.(xlsx|xlsm|xltx|xltm|xls|xlt|csv)$/i, '') || 'Workbook';
}

/**
 * The name a save is written under: Office Open XML names are kept, anything
 * else (a legacy `.xls` that was opened, a CSV, no extension) becomes `.xlsx`,
 * since the viewer always writes the Open XML format.
 */
export function savedFileName(fileName: string): string {
	return /\.(xlsx|xlsm|xltx|xltm)$/i.test(fileName) ? fileName : `${baseName(fileName)}.xlsx`;
}

/** Passwords `.xlsx` exports are encrypted with, by workbook (viewer state, not the model). */
const passwords = new WeakMap<Workbook, string>();

/** Encrypt `.xlsx` exports of `workbook` with `password`; `undefined` or '' exports plain files. */
export function setWorkbookPassword(workbook: Workbook, password: string | undefined): void {
	if (password) {
		passwords.set(workbook, password);
	} else {
		passwords.delete(workbook);
	}
}

/** The password `.xlsx` exports of `workbook` are encrypted with, if any. */
export function workbookPassword(workbook: Workbook): string | undefined {
	return passwords.get(workbook);
}

/** Produce the bytes / text of an export format. */
export async function exportWorkbook(
	workbook: Workbook,
	sheetIndex: number,
	format: DownloadFormat,
): Promise<BlobPart> {
	switch (format) {
		case 'xlsx': {
			const password = passwords.get(workbook);
			return (await writeWorkbook(workbook, password ? { password } : {})) as BlobPart;
		}
		case 'csv':
			// Excel adds a BOM so other apps detect UTF-8.
			return `\uFEFF${sheetToCsv(workbook, workbook.sheets[sheetIndex])}`;
		case 'markdown':
			return workbookToMarkdown(workbook);
		case 'json':
			return JSON.stringify(workbookToJson(workbook), null, 2);
		case 'html': {
			const sheet = workbook.sheets[sheetIndex];
			return `<!doctype html><html><head><meta charset="utf-8"><title>${sheet.name}</title></head><body>${sheetToHtml(workbook, sheet)}</body></html>`;
		}
	}
}

/** Trigger a browser download. */
export function downloadBlob(data: BlobPart, fileName: string, mime: string): void {
	const blob = new Blob([data], { type: mime });
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = fileName;
	a.style.display = 'none';
	document.body.appendChild(a);
	a.click();
	setTimeout(() => {
		URL.revokeObjectURL(url);
		a.remove();
	}, 1000);
}

export async function downloadWorkbook(
	workbook: Workbook,
	sheetIndex: number,
	fileName: string,
	format: DownloadFormat,
): Promise<void> {
	const data = await exportWorkbook(workbook, sheetIndex, format);
	const name =
		format === 'csv'
			? `${baseName(fileName)} - ${workbook.sheets[sheetIndex].name}`
			: baseName(fileName);
	downloadBlob(data, `${name}.${EXT[format]}`, MIME[format]);
}

/**
 * Print a sheet with Excel's page layout (paper size, margins, scaling,
 * breaks, titles, headers and footers) through the browser's print dialog,
 * which also offers Save as PDF.
 */
export function printSheet(workbook: Workbook, sheetIndex: number): void {
	void printWorkbook(workbook, { sheets: [sheetIndex] });
}
