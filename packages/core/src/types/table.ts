import type { XmlElement } from '../xml/types';
import type { RangeAddress } from './address';
import type { AutoFilter } from './filter';

export type TotalsRowFunction =
	| 'none'
	| 'sum'
	| 'min'
	| 'max'
	| 'average'
	| 'count'
	| 'countNums'
	| 'stdDev'
	| 'var'
	| 'custom';

export interface TableColumn {
	id: number;
	name: string;
	totalsRowFunction?: TotalsRowFunction;
	totalsRowLabel?: string;
	/** Custom totals formula (display form, no leading `=`). */
	totalsRowFormula?: string;
	/** Calculated column formula (display form). */
	calculatedColumnFormula?: string;
	dataDxfId?: number;
	headerRowDxfId?: number;
	totalsRowDxfId?: number;
}

export interface TableStyleInfo {
	/** Built-in (`TableStyleMedium2`) or custom table style name. */
	name?: string;
	showFirstColumn: boolean;
	showLastColumn: boolean;
	showRowStripes: boolean;
	showColumnStripes: boolean;
}

/** An Excel table (ListObject). */
export interface Table {
	id: number;
	name: string;
	displayName: string;
	/** Whole table range including header and totals rows. */
	ref: RangeAddress;
	/** 1 when the table has a header row (default), 0 otherwise. */
	headerRowCount: number;
	/** 1 when the totals row is shown. */
	totalsRowCount: number;
	totalsRowShown?: boolean;
	columns: TableColumn[];
	autoFilter?: AutoFilter;
	style?: TableStyleInfo;
	headerRowDxfId?: number;
	dataDxfId?: number;
	totalsRowDxfId?: number;
	comment?: string;
	/** Round-trip: source part path and original root element. */
	partPath?: string;
	raw?: XmlElement;
	/** Set when the model changed since load, so the part must be regenerated. */
	dirty?: boolean;
}
