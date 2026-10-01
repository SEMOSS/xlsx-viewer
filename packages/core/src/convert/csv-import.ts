import { setCellInput } from '../edit/cell-input';
import { EditTransaction } from '../edit/transaction';
import { setColumnWidth } from '../edit/layout-ops';
import { createWorkbook } from '../sheet/workbook';
import type { Workbook } from '../types/workbook';

export interface CsvImportOptions {
	/** Field separator; detected from the first lines when omitted (comma, semicolon, tab or pipe). */
	delimiter?: string;
	/** Sheet name (default `Sheet1`; a file name without extension works well). */
	sheetName?: string;
	/** Keep every field as text instead of recognising numbers, dates and booleans. */
	textOnly?: boolean;
}

const CANDIDATES = [',', ';', '\t', '|'];

/** The delimiter that splits the first lines most consistently (quotes respected). */
export function detectCsvDelimiter(text: string): string {
	const sample = text
		.slice(0, 64 * 1024)
		.split(/\r\n|\n|\r/)
		.slice(0, 20);
	let best = ',';
	let bestScore = -1;
	for (const d of CANDIDATES) {
		const counts = sample
			.filter((l) => l.length > 0)
			.map((l) => parseCsvRows(l, d)[0]?.length ?? 1);
		if (counts.length === 0) {
			continue;
		}
		const first = counts[0];
		const consistent = counts.filter((c) => c === first).length;
		const score = first > 1 ? consistent * 1000 + first : 0;
		if (score > bestScore) {
			best = d;
			bestScore = score;
		}
	}
	return best;
}

/** Split CSV text into rows of fields (RFC 4180: quotes, doubled quotes, line breaks inside quotes). */
export function parseCsvRows(text: string, delimiter = ','): string[][] {
	const rows: string[][] = [];
	let row: string[] = [];
	let field = '';
	let quoted = false;
	let i = text.charCodeAt(0) === 0xfeff ? 1 : 0;
	for (; i < text.length; i++) {
		const ch = text[i];
		if (quoted) {
			if (ch === '"') {
				if (text[i + 1] === '"') {
					field += '"';
					i++;
				} else {
					quoted = false;
				}
			} else {
				field += ch;
			}
		} else if (ch === '"' && field === '') {
			quoted = true;
		} else if (text.startsWith(delimiter, i)) {
			row.push(field);
			field = '';
			i += delimiter.length - 1;
		} else if (ch === '\n' || ch === '\r') {
			row.push(field);
			rows.push(row);
			row = [];
			field = '';
			if (ch === '\r' && text[i + 1] === '\n') {
				i++;
			}
		} else {
			field += ch;
		}
	}
	if (field !== '' || row.length > 0) {
		row.push(field);
		rows.push(row);
	}
	return rows;
}

/**
 * Open CSV / TSV text as a one-sheet workbook, recognising values the way
 * Excel does when typing them (numbers, percentages, dates, booleans), and
 * sizing columns to their content.
 */
export function csvToWorkbook(text: string, options: CsvImportOptions = {}): Workbook {
	const delimiter = options.delimiter ?? detectCsvDelimiter(text);
	const rows = parseCsvRows(text, delimiter);
	const wb = createWorkbook({ sheetNames: [options.sheetName?.slice(0, 31) || 'Sheet1'] });
	const tx = new EditTransaction(wb, 'Import CSV');
	const widths: number[] = [];
	rows.forEach((fields, r) => {
		fields.forEach((field, c) => {
			if (field === '') {
				return;
			}
			// A leading apostrophe forces text, which also stops "=..." from becoming a formula.
			const input = options.textOnly || field.startsWith('=') ? `'${field}` : field;
			setCellInput(tx, wb, 0, r, c, input);
			widths[c] = Math.max(widths[c] ?? 0, Math.min(field.length, 60));
		});
	});
	widths.forEach((chars, c) => {
		if (chars > 8) {
			setColumnWidth(tx, wb, 0, c, c, chars + 2);
		}
	});
	tx.commit();
	for (const sheet of wb.sheets) {
		sheet.dirty = true;
	}
	return wb;
}
