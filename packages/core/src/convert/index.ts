import { cellKey, columnToLetters, keyCol, keyRow } from '../address';
import { usedRange } from '../sheet/worksheet';
import { StyleResolver } from '../styles/resolve-style';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';
import { displayText, sheetToTextGrid } from './text-grid';

export { displayText, sheetToTextGrid };

function csvField(text: string, delimiter: string): string {
	if (text.includes(delimiter) || /["\r\n]/.test(text) || /^\s|\s$/.test(text)) {
		return `"${text.replace(/"/g, '""')}"`;
	}
	return text;
}

/** CSV of a sheet's used range, with displayed (formatted) values like Excel's "Save As CSV". */
export function sheetToCsv(
	workbook: Workbook,
	sheet: Worksheet,
	options: { delimiter?: string } = {},
): string {
	const d = options.delimiter ?? ',';
	return sheetToTextGrid(workbook, sheet)
		.map((row) => row.map((t) => csvField(t, d)).join(d))
		.join('\r\n');
}

function mdEscape(text: string): string {
	return text.replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>');
}

/**
 * Markdown rendition of a workbook: one heading and table per visible sheet.
 * The first row of each sheet's used range becomes the table header.
 */
export function workbookToMarkdown(
	workbook: Workbook,
	options: { includeFormulas?: boolean } = {},
): string {
	const parts: string[] = [];
	for (const sheet of workbook.sheets) {
		if (sheet.state !== 'visible' || sheet.kind !== 'worksheet') {
			continue;
		}
		parts.push(`## ${sheet.name}`);
		const grid = sheetToTextGrid(workbook, sheet);
		if (grid.length === 0) {
			parts.push('_(empty sheet)_');
			continue;
		}
		const width = Math.max(...grid.map((r) => r.length));
		const header = grid[0].map(mdEscape);
		parts.push(`| ${header.join(' | ')} |`);
		parts.push(`| ${Array.from({ length: width }, () => '---').join(' | ')} |`);
		for (const row of grid.slice(1)) {
			parts.push(`| ${row.map(mdEscape).join(' | ')} |`);
		}
		if (options.includeFormulas) {
			const formulas = [...sheet.cells.entries()].filter(([, c]) => c.formula !== undefined);
			if (formulas.length) {
				parts.push('', '**Formulas**', '');
				for (const [key, cell] of formulas.sort((a, b) => a[0] - b[0])) {
					parts.push(
						`- \`${columnToLetters(keyCol(key))}${keyRow(key) + 1}\`: \`=${cell.formula}\``,
					);
				}
			}
		}
		parts.push('');
	}
	return parts.join('\n');
}

/** Plain JSON export: values, formulas and displayed text per sheet. */
export function workbookToJson(workbook: Workbook): unknown {
	const styles = new StyleResolver(workbook.styles, workbook.theme);
	return {
		properties: workbook.properties,
		definedNames: workbook.definedNames,
		sheets: workbook.sheets.map((sheet) => ({
			name: sheet.name,
			state: sheet.state,
			cells: [...sheet.cells.entries()]
				.filter(([, c]) => c.value !== null || c.formula !== undefined)
				.sort((a, b) => a[0] - b[0])
				.map(([key, c]) => ({
					ref: `${columnToLetters(keyCol(key))}${keyRow(key) + 1}`,
					value: isCellError(c.value) ? c.value.code : c.value,
					formula: c.formula,
					text: displayText(c, styles, workbook.date1904),
				})),
			merges: sheet.merges,
		})),
	};
}

function htmlEscape(text: string): string {
	return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** A simple HTML table for a sheet (used for printing and HTML export). */
export function sheetToHtml(workbook: Workbook, sheet: Worksheet): string {
	const range = usedRange(sheet);
	if (!range) {
		return '<table></table>';
	}
	const styles = new StyleResolver(workbook.styles, workbook.theme);
	const skip = new Set<number>();
	const rows: string[] = [];
	for (let r = range.top; r <= range.bottom; r++) {
		const cells: string[] = [];
		for (let c = range.left; c <= range.right; c++) {
			const key = cellKey(r, c);
			if (skip.has(key)) {
				continue;
			}
			const merge = sheet.merges.find((m) => m.top === r && m.left === c);
			if (merge) {
				for (let mr = merge.top; mr <= merge.bottom; mr++) {
					for (let mc = merge.left; mc <= merge.right; mc++) {
						if (mr !== r || mc !== c) {
							skip.add(cellKey(mr, mc));
						}
					}
				}
			}
			const cell = sheet.cells.get(key);
			const style = styles.resolve(cell?.style);
			const css: string[] = [];
			if (style.font.bold) css.push('font-weight:700');
			if (style.font.italic) css.push('font-style:italic');
			const numeric = typeof cell?.value === 'number';
			const align =
				style.alignment.horizontal && style.alignment.horizontal !== 'general'
					? style.alignment.horizontal
					: numeric
						? 'right'
						: 'left';
			css.push(`text-align:${align}`);
			const span = merge
				? ` colspan="${merge.right - merge.left + 1}" rowspan="${merge.bottom - merge.top + 1}"`
				: '';
			cells.push(
				`<td${span} style="${css.join(';')}">${htmlEscape(displayText(cell, styles, workbook.date1904))}</td>`,
			);
		}
		rows.push(`<tr>${cells.join('')}</tr>`);
	}
	return `<table border="1" cellspacing="0" cellpadding="3" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt">${rows.join('')}</table>`;
}
export * from './csv-import';
