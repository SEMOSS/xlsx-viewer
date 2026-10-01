/**
 * Sheet-name quoting rules for formulas: names that are not plain identifiers
 * (spaces, punctuation, leading digits, or that look like a cell reference)
 * must be wrapped in single quotes, with embedded quotes doubled.
 */
export function quoteSheetName(name: string): string {
	if (needsQuoting(name)) {
		return `'${name.replace(/'/g, "''")}'`;
	}
	return name;
}

function needsQuoting(name: string): boolean {
	if (name.length === 0) {
		return true;
	}
	if (!/^[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*$/.test(name)) {
		return true;
	}
	// Looks like a cell reference (A1, XFD1048576) or an R1C1 reference.
	if (/^[A-Za-z]{1,3}\d+$/.test(name)) {
		return true;
	}
	if (/^[Rr](\d*)[Cc](\d*)$/.test(name) || /^[RrCc]$/.test(name)) {
		return true;
	}
	if (/^(TRUE|FALSE)$/i.test(name)) {
		return true;
	}
	return false;
}

/** Remove surrounding quotes and un-double embedded quotes. */
export function unquoteSheetName(text: string): string {
	if (text.length >= 2 && text.startsWith("'") && text.endsWith("'")) {
		return text.slice(1, -1).replace(/''/g, "'");
	}
	return text;
}

/** Characters Excel forbids in sheet names. */
const INVALID_SHEET_CHARS = /[\\/?*[\]:]/;

/** Validate a sheet name: 1-31 chars, none of `\ / ? * [ ] :`, not starting/ending with `'`. */
export function isValidSheetName(name: string): boolean {
	return (
		name.length > 0 &&
		name.length <= 31 &&
		!INVALID_SHEET_CHARS.test(name) &&
		!name.startsWith("'") &&
		!name.endsWith("'") &&
		name.toLowerCase() !== 'history'
	);
}
