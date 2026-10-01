import type { RangeAddress } from './address';

export type CustomFilterOperator =
	| 'equal'
	| 'lessThan'
	| 'lessThanOrEqual'
	| 'notEqual'
	| 'greaterThanOrEqual'
	| 'greaterThan';

export interface CustomFilter {
	operator?: CustomFilterOperator;
	/** Comparison value; may contain `*` and `?` wildcards for text. */
	val: string;
}

export interface DateGroupItem {
	dateTimeGrouping: 'year' | 'month' | 'day' | 'hour' | 'minute' | 'second';
	year: number;
	month?: number;
	day?: number;
	hour?: number;
	minute?: number;
	second?: number;
}

export interface FilterColumn {
	/** Zero-based column offset within the filter range. */
	colId: number;
	hiddenButton?: boolean;
	showButton?: boolean;
	/** Value list filter: show rows whose displayed text is one of `values`. */
	filters?: {
		values: string[];
		blank?: boolean;
		dateGroups?: DateGroupItem[];
	};
	customFilters?: {
		/** `true` = both conditions (AND), otherwise either (OR). */
		and?: boolean;
		filters: CustomFilter[];
	};
	top10?: { top?: boolean; percent?: boolean; val: number; filterVal?: number };
	dynamicFilter?: { type: string; val?: number; maxVal?: number };
	colorFilter?: { dxfId?: number; cellColor?: boolean };
	iconFilter?: { iconSet: string; iconId?: number };
}

export interface SortCondition {
	ref: RangeAddress;
	descending?: boolean;
	sortBy?: 'value' | 'cellColor' | 'fontColor' | 'icon';
	customList?: string;
	dxfId?: number;
}

export interface SortState {
	ref: RangeAddress;
	columnSort?: boolean;
	caseSensitive?: boolean;
	conditions: SortCondition[];
}

export interface AutoFilter {
	/** The filtered block including the header row. */
	ref: RangeAddress;
	columns: FilterColumn[];
	sortState?: SortState;
}
