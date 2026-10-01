/** Categories used by the Insert Function dialog (Excel's function library groups). */
export type FunctionCategory =
	| 'Financial'
	| 'Date & Time'
	| 'Math & Trig'
	| 'Statistical'
	| 'Lookup & Reference'
	| 'Database'
	| 'Text'
	| 'Logical'
	| 'Information'
	| 'Engineering'
	| 'Cube'
	| 'Compatibility'
	| 'Web';

export interface FunctionArgInfo {
	/** Argument name as Excel shows it (`lookup_value`, `number1`). */
	name: string;
	description: string;
	optional?: boolean;
	/**
	 * The argument can be supplied again and again. Consecutive repeating
	 * arguments form one group that repeats as a unit: SUM marks `number2`
	 * (`number1, [number2], ...`), SUMIFS marks `criteria_range2` and
	 * `criteria2` (`[criteria_range2, criteria2], ...`).
	 */
	repeating?: boolean;
}

/** Help text for one worksheet function. */
export interface FunctionInfo {
	/** Upper-case name, e.g. `XLOOKUP`, `NORM.S.DIST`. */
	name: string;
	category: FunctionCategory;
	/** One-sentence description, as in Excel's Insert Function dialog. */
	description: string;
	args: FunctionArgInfo[];
	/** Excel version that introduced it, e.g. `2010`, `2019`, `365`. */
	since?: string;
}
