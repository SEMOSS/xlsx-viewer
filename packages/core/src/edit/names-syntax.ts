import { lettersToColumn } from '../address/cell-address';
import { MAX_COLS, MAX_ROWS } from '../address/constants';
import type { DefinedName } from '../types/workbook';

/**
 * Lexical rules for defined names (Formulas > Define Name): which strings
 * Excel accepts as a name, and how a label is turned into one.
 */

/** Prefix of built-in names in the file (`_xlnm.Print_Area`). */
export const BUILTIN_NAME_PREFIX = '_xlnm.';

const MAX_NAME_LENGTH = 255;
/** First character: a letter of any script, `_` or `\`. */
const START_CHAR = /^[\p{L}_\\]$/u;
/** Later characters: letters, combining marks, digits, `_`, `.`, `\` and `?`. */
const PART_CHAR = /^[\p{L}\p{M}\p{N}_.\\?]$/u;
const VALID_NAME = /^[\p{L}_\\][\p{L}\p{M}\p{N}_.\\?]*$/u;

/** Name shown in the UI: built-in names lose their `_xlnm.` prefix (`Print_Area`). */
export function definedNameDisplayName(def: Pick<DefinedName, 'name'>): string {
	return def.name.toLowerCase().startsWith(BUILTIN_NAME_PREFIX)
		? def.name.slice(BUILTIN_NAME_PREFIX.length)
		: def.name;
}

function inRange(digits: string, max: number): boolean {
	if (digits === '') {
		return true;
	}
	const n = Number(digits);
	return n >= 1 && n <= max;
}

/**
 * The text reads as a cell reference, so Excel refuses it as a name: A1
 * style (`A1`, `xfd1048576`, `AB12`) or R1C1 style (`R`, `C`, `R1`, `C5`,
 * `RC`, `R2C3`).
 */
export function looksLikeCellReference(text: string): boolean {
	const a1 = /^([A-Za-z]{1,3})(\d+)$/.exec(text);
	if (a1) {
		const col = lettersToColumn(a1[1]);
		const row = Number(a1[2]);
		if (col >= 0 && row >= 1 && row <= MAX_ROWS) {
			return true;
		}
	}
	const rc = /^[Rr](\d*)[Cc](\d*)$/.exec(text);
	if (rc && inRange(rc[1], MAX_ROWS) && inRange(rc[2], MAX_COLS)) {
		return true;
	}
	const r = /^[Rr](\d*)$/.exec(text);
	if (r && inRange(r[1], MAX_ROWS)) {
		return true;
	}
	const c = /^[Cc](\d*)$/.exec(text);
	return c !== null && inRange(c[1], MAX_COLS);
}

/**
 * Excel's lexical rules for a defined name: 1 to 255 characters, starting
 * with a letter, `_` or `\`, followed by letters, digits, `_`, `.`, `\` or
 * `?`; no spaces; not a cell reference (A1 or R1C1) and not TRUE / FALSE.
 */
export function isValidDefinedNameSyntax(name: string): boolean {
	if (name.length < 1 || name.length > MAX_NAME_LENGTH || !VALID_NAME.test(name)) {
		return false;
	}
	return !looksLikeCellReference(name) && !/^(?:true|false)$/i.test(name);
}

/**
 * Turn a label into a valid name the way Excel suggests one: spaces and
 * other invalid characters become `_`, a name that cannot start with its
 * first character gets a leading `_`, and a reference look-alike gets a
 * trailing `_` (`Q1` -> `Q1_`). Returns `''` when nothing usable is left.
 */
export function sanitizeDefinedName(label: string): string {
	const chars = Array.from(label.trim());
	if (chars.length === 0) {
		return '';
	}
	let out = chars.map((ch) => (PART_CHAR.test(ch) || START_CHAR.test(ch) ? ch : '_')).join('');
	if (!START_CHAR.test(Array.from(out)[0])) {
		out = `_${out}`;
	}
	if (looksLikeCellReference(out) || /^(?:true|false)$/i.test(out)) {
		out = `${out}_`;
	}
	if (out.length > MAX_NAME_LENGTH) {
		out = Array.from(out).slice(0, MAX_NAME_LENGTH).join('');
	}
	return isValidDefinedNameSyntax(out) ? out : '';
}
