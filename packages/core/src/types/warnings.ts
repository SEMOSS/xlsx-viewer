/** Something in the file was approximated, skipped or could not be read. */
export interface XlsxWarning {
	/** Machine-readable code, e.g. `unsupported-function`, `chart-type-approximated`. */
	code: string;
	message: string;
	/** Package path of the part involved, when relevant. */
	part?: string;
	/** Sheet name, when relevant. */
	sheet?: string;
	/** Cell reference (A1), when relevant. */
	cell?: string;
}
