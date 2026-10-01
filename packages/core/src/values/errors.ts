import type { CellError, CellValue, ErrorCode } from '../types/cell';

const ERROR_CODES: readonly ErrorCode[] = [
	'#NULL!',
	'#DIV/0!',
	'#VALUE!',
	'#REF!',
	'#NAME?',
	'#NUM!',
	'#N/A',
	'#GETTING_DATA',
	'#SPILL!',
	'#CALC!',
	'#FIELD!',
	'#BLOCKED!',
	'#CONNECT!',
	'#BUSY!',
	'#UNKNOWN!',
];

const SINGLETONS = new Map<ErrorCode, CellError>(
	ERROR_CODES.map((code) => [code, Object.freeze({ type: 'error' as const, code })]),
);

/** The shared, frozen error value for a code. */
export function cellError(code: ErrorCode): CellError {
	return SINGLETONS.get(code)!;
}

/** Frozen error singletons, e.g. `ERRORS.DIV0`. */
export const ERRORS = {
	NULL: cellError('#NULL!'),
	DIV0: cellError('#DIV/0!'),
	VALUE: cellError('#VALUE!'),
	REF: cellError('#REF!'),
	NAME: cellError('#NAME?'),
	NUM: cellError('#NUM!'),
	NA: cellError('#N/A'),
	GETTING_DATA: cellError('#GETTING_DATA'),
	SPILL: cellError('#SPILL!'),
	CALC: cellError('#CALC!'),
	FIELD: cellError('#FIELD!'),
	BLOCKED: cellError('#BLOCKED!'),
	CONNECT: cellError('#CONNECT!'),
	BUSY: cellError('#BUSY!'),
	UNKNOWN: cellError('#UNKNOWN!'),
} as const;

/** Type guard for error values. */
export function isCellError(v: unknown): v is CellError {
	return typeof v === 'object' && v !== null && (v as { type?: unknown }).type === 'error';
}

/** Parse an error literal such as `#DIV/0!` (case-insensitive). */
export function parseErrorCode(text: string): ErrorCode | null {
	const upper = text.trim().toUpperCase();
	for (const code of ERROR_CODES) {
		if (code === upper) {
			return code;
		}
	}
	return null;
}

/** `ERROR.TYPE` numbering. */
export function errorTypeNumber(code: ErrorCode): number {
	switch (code) {
		case '#NULL!':
			return 1;
		case '#DIV/0!':
			return 2;
		case '#VALUE!':
			return 3;
		case '#REF!':
			return 4;
		case '#NAME?':
			return 5;
		case '#NUM!':
			return 6;
		case '#N/A':
			return 7;
		case '#GETTING_DATA':
			return 8;
		case '#SPILL!':
			return 9;
		case '#CONNECT!':
			return 10;
		case '#BLOCKED!':
			return 11;
		case '#UNKNOWN!':
			return 12;
		case '#FIELD!':
			return 13;
		case '#CALC!':
			return 14;
		default:
			return 3;
	}
}

/** Human-readable classification of a value. */
export function valueKind(v: CellValue): 'blank' | 'number' | 'string' | 'boolean' | 'error' {
	if (v === null) {
		return 'blank';
	}
	if (typeof v === 'number') {
		return 'number';
	}
	if (typeof v === 'string') {
		return 'string';
	}
	if (typeof v === 'boolean') {
		return 'boolean';
	}
	return 'error';
}
