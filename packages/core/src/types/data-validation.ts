import type { RangeAddress } from './address';

export type DataValidationType =
	| 'none'
	| 'whole'
	| 'decimal'
	| 'list'
	| 'date'
	| 'time'
	| 'textLength'
	| 'custom';

export type DataValidationOperator =
	| 'between'
	| 'notBetween'
	| 'equal'
	| 'notEqual'
	| 'lessThan'
	| 'lessThanOrEqual'
	| 'greaterThan'
	| 'greaterThanOrEqual';

export interface DataValidation {
	ranges: RangeAddress[];
	type: DataValidationType;
	operator?: DataValidationOperator;
	/**
	 * First criteria formula in display form (no leading `=`). For list
	 * validations this is either a quoted, comma-separated literal list
	 * (`"Yes,No"`) or a reference/formula (`$A$1:$A$5`, `MyList`).
	 */
	formula1?: string;
	formula2?: string;
	allowBlank?: boolean;
	/**
	 * Whether to show the in-cell dropdown for list validations. Note that the
	 * OOXML attribute `showDropDown="1"` means the opposite (suppress it); the
	 * reader and writer translate.
	 */
	inCellDropdown?: boolean;
	showInputMessage?: boolean;
	showErrorMessage?: boolean;
	errorStyle?: 'stop' | 'warning' | 'information';
	errorTitle?: string;
	error?: string;
	promptTitle?: string;
	prompt?: string;
	/** The validation came from the Excel 2010+ `extLst` (cross-sheet list sources). */
	fromExtension?: boolean;
}
