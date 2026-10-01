/**
 * Table parts (`xl/tables/tableN.xml`). Regenerated tables start from the
 * original root so unknown attributes (`xr:uid`, query table links) and
 * column extras survive.
 */
import { formatRange } from '../address/range';
import type { Table, TableColumn } from '../types/table';
import { cloneElement, el, serializeXml, textEl } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { buildAutoFilter } from './auto-filter';
import { fileFormula } from './formulas';
import { findChild, findChildren, mergeAttrs, normalizeRoot, setOrderedChild } from './xml-utils';

const TABLE_ORDER = [
	'autoFilter',
	'sortState',
	'tableColumns',
	'tableStyleInfo',
	'extLst',
] as const;
const COLUMN_ORDER = [
	'calculatedColumnFormula',
	'totalsRowFormula',
	'xmlColumnPr',
	'extLst',
] as const;

function columnElement(col: TableColumn, orig: XmlElement | undefined, p: string): XmlElement {
	const fn =
		col.totalsRowFunction && col.totalsRowFunction !== 'none' ? col.totalsRowFunction : undefined;
	const out = el(
		'tableColumn',
		mergeAttrs(orig?.attrs, {
			id: col.id,
			name: col.name,
			totalsRowFunction: fn,
			totalsRowLabel: fn ? undefined : col.totalsRowLabel,
			headerRowDxfId: col.headerRowDxfId,
			dataDxfId: col.dataDxfId,
			totalsRowDxfId: col.totalsRowDxfId,
		}),
		orig ? cloneElement(orig).children : [],
	);
	setOrderedChild(
		out,
		'calculatedColumnFormula',
		col.calculatedColumnFormula
			? textEl('calculatedColumnFormula', fileFormula(col.calculatedColumnFormula))
			: null,
		COLUMN_ORDER,
		p,
	);
	setOrderedChild(
		out,
		'totalsRowFormula',
		fn === 'custom' && col.totalsRowFormula
			? textEl('totalsRowFormula', fileFormula(col.totalsRowFormula))
			: null,
		COLUMN_ORDER,
		p,
	);
	return out;
}

/**
 * Columns as Excel requires them: exactly one per column of the range, with
 * unique positive ids and unique non-empty names.
 */
export function normalizeTableColumns(table: Table): TableColumn[] {
	const width = table.ref.right - table.ref.left + 1;
	const cols = table.columns.slice(0, width);
	const ids = new Set<number>();
	let validIds = true;
	for (const c of cols) {
		if (!Number.isInteger(c.id) || c.id <= 0 || ids.has(c.id)) {
			validIds = false;
		}
		ids.add(c.id);
	}
	let nextId = cols.reduce((m, c) => (Number.isInteger(c.id) ? Math.max(m, c.id) : m), 0) + 1;
	const out: TableColumn[] = cols.map((c, i) => (validIds ? c : { ...c, id: i + 1 }));
	if (!validIds) {
		nextId = out.length + 1;
	}
	while (out.length < width) {
		out.push({ id: nextId++, name: `Column${out.length + 1}` });
	}
	const seen = new Set<string>();
	return out.map((c, i) => {
		let name = c.name && c.name.trim() ? c.name : `Column${i + 1}`;
		const base = name;
		for (let n = 2; seen.has(name.toLowerCase()); n++) {
			name = `${base}${n}`;
		}
		seen.add(name.toLowerCase());
		return name === c.name ? c : { ...c, name };
	});
}

/** Serialize a table part. `id` is the workbook-unique table id to write. */
export function buildTableXml(table: Table, id: number): string {
	const root = table.raw ? cloneElement(table.raw) : el('table', { xmlns: NS.main });
	const p = normalizeRoot(root, 'table', NS.main);
	const headerRows = table.headerRowCount === 0 ? 0 : 1;
	const totalsRows = table.totalsRowCount > 0 ? 1 : 0;
	root.attrs = mergeAttrs(root.attrs, {
		id,
		name: table.name,
		displayName: table.displayName || table.name,
		comment: table.comment,
		ref: formatRange(table.ref, { forceCells: true }),
		headerRowCount: headerRows === 0 ? 0 : undefined,
		totalsRowCount: totalsRows ? 1 : undefined,
		totalsRowShown: table.totalsRowShown ?? (totalsRows ? undefined : false),
		headerRowDxfId: table.headerRowDxfId,
		dataDxfId: table.dataDxfId,
		totalsRowDxfId: table.totalsRowDxfId,
	});
	const set = (name: string, value: XmlElement | null): void =>
		setOrderedChild(root, name, value, TABLE_ORDER, p);

	set('autoFilter', headerRows && table.autoFilter ? buildAutoFilter(table.autoFilter) : null);
	if (!headerRows) {
		set('sortState', null);
	}
	const origColumns = findChildren(findChild(root, 'tableColumns', p), 'tableColumn', p);
	const byId = new Map(origColumns.map((c) => [c.attrs.id, c]));
	const columns = normalizeTableColumns(table);
	set(
		'tableColumns',
		el(
			'tableColumns',
			{ count: columns.length },
			columns.map((c) => columnElement(c, byId.get(String(c.id)), p)),
		),
	);
	const style = table.style;
	set(
		'tableStyleInfo',
		style
			? el(
					'tableStyleInfo',
					mergeAttrs(findChild(root, 'tableStyleInfo', p)?.attrs, {
						name: style.name,
						showFirstColumn: style.showFirstColumn,
						showLastColumn: style.showLastColumn,
						showRowStripes: style.showRowStripes,
						showColumnStripes: style.showColumnStripes,
					}),
				)
			: null,
	);
	return serializeXml(root);
}
