import { parseRange } from '../address/range';
import type {
	AutoFilter,
	CustomFilter,
	CustomFilterOperator,
	DateGroupItem,
	FilterColumn,
	SortCondition,
	SortState,
} from '../types/filter';
import type { XmlElement } from '../xml/types';
import { bool, compact, int, kid, kids, num, oneOf, str } from './xml-utils';

/** `autoFilter` and `sortState` (worksheets and tables). */

const OPERATORS: readonly CustomFilterOperator[] = [
	'equal',
	'lessThan',
	'lessThanOrEqual',
	'notEqual',
	'greaterThanOrEqual',
	'greaterThan',
];

const GROUPINGS = ['year', 'month', 'day', 'hour', 'minute', 'second'] as const;

function parseDateGroup(el: XmlElement): DateGroupItem | undefined {
	const grouping = oneOf(str(el, 'dateTimeGrouping'), GROUPINGS);
	const year = int(el, 'year');
	if (!grouping || year === undefined) {
		return undefined;
	}
	return compact<DateGroupItem>({
		dateTimeGrouping: grouping,
		year,
		month: int(el, 'month'),
		day: int(el, 'day'),
		hour: int(el, 'hour'),
		minute: int(el, 'minute'),
		second: int(el, 'second'),
	});
}

function parseFilterColumn(el: XmlElement): FilterColumn {
	const col: FilterColumn = compact({
		colId: int(el, 'colId') ?? 0,
		hiddenButton: bool(el, 'hiddenButton'),
		showButton: bool(el, 'showButton'),
	});
	const filters = kid(el, 'filters');
	if (filters) {
		const values = kids(filters, 'filter').map((f) => str(f, 'val') ?? '');
		const dateGroups = kids(filters, 'dateGroupItem')
			.map(parseDateGroup)
			.filter((d): d is DateGroupItem => d !== undefined);
		col.filters = compact({
			values,
			blank: bool(filters, 'blank'),
			dateGroups: dateGroups.length > 0 ? dateGroups : undefined,
		});
	}
	const custom = kid(el, 'customFilters');
	if (custom) {
		const list: CustomFilter[] = kids(custom, 'customFilter').map((f) =>
			compact<CustomFilter>({
				operator: oneOf(str(f, 'operator'), OPERATORS),
				val: str(f, 'val') ?? '',
			}),
		);
		col.customFilters = compact({ and: bool(custom, 'and'), filters: list });
	}
	const top10 = kid(el, 'top10');
	if (top10) {
		col.top10 = compact({
			top: bool(top10, 'top'),
			percent: bool(top10, 'percent'),
			val: num(top10, 'val') ?? 10,
			filterVal: num(top10, 'filterVal'),
		});
	}
	const dynamic = kid(el, 'dynamicFilter');
	if (dynamic) {
		col.dynamicFilter = compact({
			type: str(dynamic, 'type') ?? 'null',
			val: num(dynamic, 'val'),
			maxVal: num(dynamic, 'maxVal'),
		});
	}
	const color = kid(el, 'colorFilter');
	if (color) {
		col.colorFilter = compact({ dxfId: int(color, 'dxfId'), cellColor: bool(color, 'cellColor') });
	}
	const icon = kid(el, 'iconFilter');
	if (icon) {
		col.iconFilter = compact({
			iconSet: str(icon, 'iconSet') ?? '3Arrows',
			iconId: int(icon, 'iconId'),
		});
	}
	return col;
}

/** `<sortState>`. */
export function parseSortState(el: XmlElement | undefined): SortState | undefined {
	const ref = el ? parseRange(str(el, 'ref') ?? '') : null;
	if (!el || !ref) {
		return undefined;
	}
	const conditions: SortCondition[] = [];
	for (const c of kids(el, 'sortCondition')) {
		const cref = parseRange(str(c, 'ref') ?? '');
		if (!cref) {
			continue;
		}
		conditions.push(
			compact<SortCondition>({
				ref: cref,
				descending: bool(c, 'descending'),
				sortBy: oneOf(str(c, 'sortBy'), ['value', 'cellColor', 'fontColor', 'icon'] as const),
				customList: str(c, 'customList'),
				dxfId: int(c, 'dxfId'),
			}),
		);
	}
	return compact<SortState>({
		ref,
		columnSort: bool(el, 'columnSort'),
		caseSensitive: bool(el, 'caseSensitive'),
		conditions,
	});
}

/** `<autoFilter ref>` with its filter columns and nested sort state. */
export function parseAutoFilter(el: XmlElement | undefined): AutoFilter | undefined {
	const ref = el ? parseRange(str(el, 'ref') ?? '') : null;
	if (!el || !ref) {
		return undefined;
	}
	const filter: AutoFilter = { ref, columns: kids(el, 'filterColumn').map(parseFilterColumn) };
	const sort = parseSortState(kid(el, 'sortState'));
	if (sort) {
		filter.sortState = sort;
	}
	return filter;
}
