/** `<autoFilter>` and `<sortState>` elements (worksheets and tables). */
import { formatRange } from '../address/range';
import type { AutoFilter, FilterColumn, SortCondition, SortState } from '../types/filter';
import { el } from '../xml';
import type { XmlElement } from '../xml/types';
import { flag, unflag } from './xml-utils';

function ref(r: AutoFilter['ref']): string {
	return formatRange(r, { forceCells: true });
}

function filterColumn(fc: FilterColumn): XmlElement {
	const kids: XmlElement[] = [];
	if (fc.filters) {
		kids.push(
			el('filters', { blank: flag(fc.filters.blank) }, [
				...fc.filters.values.map((v) => el('filter', { val: v })),
				...(fc.filters.dateGroups ?? []).map((g) =>
					el('dateGroupItem', {
						year: g.year,
						month: g.month,
						day: g.day,
						hour: g.hour,
						minute: g.minute,
						second: g.second,
						dateTimeGrouping: g.dateTimeGrouping,
					}),
				),
			]),
		);
	} else if (fc.top10) {
		kids.push(
			el('top10', {
				top: unflag(fc.top10.top),
				percent: flag(fc.top10.percent),
				val: fc.top10.val,
				filterVal: fc.top10.filterVal,
			}),
		);
	} else if (fc.customFilters && fc.customFilters.filters.length) {
		kids.push(
			el(
				'customFilters',
				{ and: flag(fc.customFilters.and) },
				fc.customFilters.filters.slice(0, 2).map((f) =>
					el('customFilter', {
						operator: f.operator && f.operator !== 'equal' ? f.operator : undefined,
						val: f.val,
					}),
				),
			),
		);
	} else if (fc.dynamicFilter) {
		kids.push(
			el('dynamicFilter', {
				type: fc.dynamicFilter.type,
				val: fc.dynamicFilter.val,
				maxVal: fc.dynamicFilter.maxVal,
			}),
		);
	} else if (fc.colorFilter) {
		kids.push(
			el('colorFilter', {
				dxfId: fc.colorFilter.dxfId,
				cellColor: unflag(fc.colorFilter.cellColor),
			}),
		);
	} else if (fc.iconFilter) {
		kids.push(el('iconFilter', { iconSet: fc.iconFilter.iconSet, iconId: fc.iconFilter.iconId }));
	}
	return el(
		'filterColumn',
		{ colId: fc.colId, hiddenButton: flag(fc.hiddenButton), showButton: unflag(fc.showButton) },
		kids,
	);
}

function sortCondition(c: SortCondition): XmlElement {
	return el('sortCondition', {
		descending: flag(c.descending),
		sortBy: c.sortBy && c.sortBy !== 'value' ? c.sortBy : undefined,
		ref: ref(c.ref),
		customList: c.customList,
		dxfId: c.dxfId,
	});
}

/** A `<sortState>` element. */
export function buildSortState(s: SortState): XmlElement {
	return el(
		'sortState',
		{ columnSort: flag(s.columnSort), caseSensitive: flag(s.caseSensitive), ref: ref(s.ref) },
		s.conditions.slice(0, 64).map(sortCondition),
	);
}

/** An `<autoFilter>` element (filter columns plus an optional nested sort state). */
export function buildAutoFilter(af: AutoFilter): XmlElement {
	const cols = [...af.columns].sort((a, b) => a.colId - b.colId);
	return el('autoFilter', { ref: ref(af.ref) }, [
		...cols.map(filterColumn),
		af.sortState ? buildSortState(af.sortState) : undefined,
	]);
}
