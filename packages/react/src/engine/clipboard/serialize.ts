/**
 * Clipboard formats shared with other spreadsheet apps: tab-separated text
 * and HTML tables (what Excel, Google Sheets and LibreOffice put on and read
 * from the system clipboard).
 */

/** Quote a TSV field when it contains tabs, newlines or quotes. */
function tsvField(text: string): string {
	if (/[\t\n\r"]/.test(text)) {
		return `"${text.replace(/"/g, '""')}"`;
	}
	return text;
}

export function toTsv(rows: string[][]): string {
	return rows.map((r) => r.map(tsvField).join('\t')).join('\r\n') + '\r\n';
}

/** Parse tab-separated text with Excel's quoting rules. */
export function parseTsv(text: string): string[][] {
	const rows: string[][] = [];
	let row: string[] = [];
	let field = '';
	let i = 0;
	let quoted = false;
	const src = text.replace(/\r\n?/g, '\n');
	while (i < src.length) {
		const ch = src[i];
		if (quoted) {
			if (ch === '"') {
				if (src[i + 1] === '"') {
					field += '"';
					i += 2;
					continue;
				}
				quoted = false;
				i++;
				continue;
			}
			field += ch;
			i++;
			continue;
		}
		if (ch === '"' && field === '') {
			quoted = true;
			i++;
			continue;
		}
		if (ch === '\t') {
			row.push(field);
			field = '';
		} else if (ch === '\n') {
			row.push(field);
			rows.push(row);
			row = [];
			field = '';
		} else {
			field += ch;
		}
		i++;
	}
	if (field !== '' || row.length > 0) {
		row.push(field);
		rows.push(row);
	}
	return rows;
}

function escapeHtml(text: string): string {
	return text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

export interface HtmlCell {
	text: string;
	style?: string;
	colspan?: number;
	rowspan?: number;
}

/** An HTML table with inline styles, readable by Excel / Sheets / Word. */
export function toHtmlTable(rows: HtmlCell[][], colWidths?: number[]): string {
	const cols =
		colWidths
			?.map((w) => `<col width="${Math.round(w)}" style="width:${Math.round(w)}px">`)
			.join('') ?? '';
	const body = rows
		.map(
			(r) =>
				`<tr>${r
					.map((c) => {
						const attrs = [
							c.style ? ` style="${escapeHtml(c.style)}"` : '',
							c.colspan && c.colspan > 1 ? ` colspan="${c.colspan}"` : '',
							c.rowspan && c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : '',
						].join('');
						return `<td${attrs}>${escapeHtml(c.text).replace(/\n/g, '<br>')}</td>`;
					})
					.join('')}</tr>`,
		)
		.join('');
	return `<html><body><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse">${cols}${body}</table></body></html>`;
}

/** Extract the cell texts of the first table in an HTML fragment (merged cells are expanded). */
export function parseHtmlTable(html: string): string[][] | null {
	if (typeof DOMParser === 'undefined') {
		return null;
	}
	const doc = new DOMParser().parseFromString(html, 'text/html');
	const table = doc.querySelector('table');
	if (!table) {
		return null;
	}
	const grid: string[][] = [];
	const occupied: boolean[][] = [];
	Array.from(table.querySelectorAll('tr')).forEach((tr, r) => {
		grid[r] ??= [];
		occupied[r] ??= [];
		let c = 0;
		for (const td of Array.from(tr.children)) {
			if (td.tagName !== 'TD' && td.tagName !== 'TH') {
				continue;
			}
			while (occupied[r][c]) {
				c++;
			}
			const text = (td as HTMLElement).innerText ?? td.textContent ?? '';
			const colspan = Number(td.getAttribute('colspan') ?? 1) || 1;
			const rowspan = Number(td.getAttribute('rowspan') ?? 1) || 1;
			for (let dr = 0; dr < rowspan; dr++) {
				grid[r + dr] ??= [];
				occupied[r + dr] ??= [];
				for (let dc = 0; dc < colspan; dc++) {
					occupied[r + dr][c + dc] = true;
					grid[r + dr][c + dc] = dr === 0 && dc === 0 ? text.replace(/ /g, ' ').trim() : '';
				}
			}
			c += colspan;
		}
	});
	return grid.map((row) => Array.from(row, (v) => v ?? ''));
}
