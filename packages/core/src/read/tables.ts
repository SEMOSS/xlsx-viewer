import { parseRange } from '../address/range';
import { stripFormulaPrefixes } from '../formula/prefixes';
import type { Table, TableColumn, TableStyleInfo, TotalsRowFunction } from '../types/table';
import { decodeOoxmlEscapes } from '../xml/escape';
import { textContent } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { parseAutoFilter } from './auto-filter';
import { parseSpreadsheetXml } from './xml-normalize';
import { bool, compact, int, kid, kids, oneOf, str } from './xml-utils';

const TOTALS_FUNCTIONS: readonly TotalsRowFunction[] = [
	'none',
	'sum',
	'min',
	'max',
	'average',
	'count',
	'countNums',
	'stdDev',
	'var',
	'custom',
];

function formulaChild(el: XmlElement, local: string): string | undefined {
	const f = kid(el, local);
	return f ? stripFormulaPrefixes(textContent(f)) : undefined;
}

function parseColumn(el: XmlElement, index: number): TableColumn {
	return compact<TableColumn>({
		id: int(el, 'id') ?? index + 1,
		name: decodeOoxmlEscapes(str(el, 'name') ?? ''),
		totalsRowFunction: oneOf(str(el, 'totalsRowFunction'), TOTALS_FUNCTIONS),
		totalsRowLabel: str(el, 'totalsRowLabel'),
		totalsRowFormula: formulaChild(el, 'totalsRowFormula'),
		calculatedColumnFormula: formulaChild(el, 'calculatedColumnFormula'),
		dataDxfId: int(el, 'dataDxfId'),
		headerRowDxfId: int(el, 'headerRowDxfId'),
		totalsRowDxfId: int(el, 'totalsRowDxfId'),
	});
}

function parseStyleInfo(el: XmlElement | undefined): TableStyleInfo | undefined {
	if (!el) {
		return undefined;
	}
	return compact<TableStyleInfo>({
		name: str(el, 'name'),
		showFirstColumn: bool(el, 'showFirstColumn') ?? false,
		showLastColumn: bool(el, 'showLastColumn') ?? false,
		showRowStripes: bool(el, 'showRowStripes') ?? false,
		showColumnStripes: bool(el, 'showColumnStripes') ?? false,
	});
}

/** A table part (`xl/tables/tableN.xml`). Returns `undefined` when it has no valid range. */
export function parseTable(xml: string, partPath: string): Table | undefined {
	const root = parseSpreadsheetXml(xml).root;
	const ref = parseRange(str(root, 'ref') ?? '');
	if (!ref) {
		return undefined;
	}
	const name = str(root, 'name') ?? str(root, 'displayName') ?? '';
	const table: Table = compact<Table>({
		id: int(root, 'id') ?? 0,
		name,
		displayName: str(root, 'displayName') ?? name,
		ref,
		headerRowCount: int(root, 'headerRowCount') ?? 1,
		totalsRowCount: int(root, 'totalsRowCount') ?? 0,
		totalsRowShown: bool(root, 'totalsRowShown'),
		columns: kids(kid(root, 'tableColumns'), 'tableColumn').map(parseColumn),
		autoFilter: parseAutoFilter(kid(root, 'autoFilter')),
		style: parseStyleInfo(kid(root, 'tableStyleInfo')),
		headerRowDxfId: int(root, 'headerRowDxfId'),
		dataDxfId: int(root, 'dataDxfId'),
		totalsRowDxfId: int(root, 'totalsRowDxfId'),
		comment: str(root, 'comment'),
		partPath,
		raw: root,
		dirty: false,
	});
	return table;
}
