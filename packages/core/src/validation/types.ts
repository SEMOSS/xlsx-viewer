import type { FormulaEvaluator } from '../conditional/formula-eval';
import type { CellValue } from '../types/cell';
import type { Workbook } from '../types/workbook';

/** Where and how a data validation is checked. */
export interface DataValidationContext {
	/** Formula evaluator (e.g. `CalcEngine.evaluateFormula`). */
	evaluate: FormulaEvaluator;
	/** The cell being validated. */
	sheetIndex: number;
	row: number;
	col: number;
	/** The workbook uses the 1904 date system (for text-to-number conversion). */
	date1904: boolean;
	/**
	 * The workbook. Needed to read list sources given as references or defined
	 * names, and to evaluate custom formulas with the value being entered in
	 * place of the cell's current value (like Excel). Without it, list sources
	 * are resolved only through the evaluator's array results.
	 */
	workbook?: Workbook;
	/**
	 * Display text of a list source cell. Defaults to the cell's number format
	 * (via the numfmt engine) for numbers and the plain text otherwise.
	 */
	formatValue?: (value: CellValue, sheetIndex: number, row: number, col: number) => string;
}

export interface DataValidationResult {
	valid: boolean;
	/**
	 * Alert style for an invalid value: `stop` rejects it, `warning` asks
	 * "Continue?", `information` only informs. Absent when the value is valid
	 * or the rule does not enable the error alert (Excel accepts the value then).
	 */
	style?: 'stop' | 'warning' | 'information';
	/** Alert title (the rule's error title, or Excel's default "Microsoft Excel"). */
	title?: string;
	/** Alert message (the rule's error message, or Excel's default message). */
	message?: string;
}

/** One dropdown item of a list validation. */
export interface DataValidationListEntry {
	/** Text shown in the dropdown (the source cell's displayed value). */
	text: string;
	/** Underlying value (the source cell value, or the literal item text). */
	value: CellValue;
}

/** Title of Excel's error alert when the rule defines none. */
export const DEFAULT_VALIDATION_ERROR_TITLE = 'Microsoft Excel';

/** Excel's error alert message when the rule defines none. */
export const DEFAULT_VALIDATION_ERROR_MESSAGE =
	"This value doesn't match the data validation restrictions defined for this cell.";
