/** A formula that does not follow Excel's grammar. */
export class FormulaParseError extends Error {
	/** Offset in the formula text where the problem was detected. */
	readonly position: number;

	constructor(message: string, position: number) {
		super(message);
		this.name = 'FormulaParseError';
		this.position = position;
	}
}
