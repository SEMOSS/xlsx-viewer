/**
 * Character classification for the formula tokenizer. Everything works on
 * UTF-16 char codes so the hot loops never allocate.
 */

export const CH_TAB = 9;
export const CH_LF = 10;
export const CH_CR = 13;
export const CH_SPACE = 32;
export const CH_BANG = 33; // !
export const CH_DQUOTE = 34; // "
export const CH_HASH = 35; // #
export const CH_DOLLAR = 36; // $
export const CH_PERCENT = 37; // %
export const CH_AMP = 38; // &
export const CH_SQUOTE = 39; // '
export const CH_LPAREN = 40; // (
export const CH_RPAREN = 41; // )
export const CH_STAR = 42; // *
export const CH_PLUS = 43; // +
export const CH_COMMA = 44; // ,
export const CH_MINUS = 45; // -
export const CH_DOT = 46; // .
export const CH_SLASH = 47; // /
export const CH_0 = 48;
export const CH_9 = 57;
export const CH_COLON = 58; // :
export const CH_SEMI = 59; // ;
export const CH_LT = 60; // <
export const CH_EQ = 61; // =
export const CH_GT = 62; // >
export const CH_QUESTION = 63; // ?
export const CH_AT = 64; // @
export const CH_LBRACKET = 91; // [
export const CH_BACKSLASH = 92; // \
export const CH_RBRACKET = 93; // ]
export const CH_CARET = 94; // ^
export const CH_UNDERSCORE = 95; // _
export const CH_LBRACE = 123; // {
export const CH_RBRACE = 125; // }

/** ASCII lookup: 1 = may start a name, 2 = may continue a name (bit flags). */
const ASCII = new Uint8Array(128);
const START = 1;
const PART = 2;

for (let c = 65; c <= 90; c++) {
	ASCII[c] = START | PART;
	ASCII[c + 32] = START | PART;
}
for (let c = CH_0; c <= CH_9; c++) {
	ASCII[c] = PART;
}
ASCII[CH_UNDERSCORE] = START | PART;
ASCII[CH_BACKSLASH] = START | PART;
ASCII[CH_DOT] = PART;
ASCII[CH_QUESTION] = PART;

/**
 * Excel accepts letters from any script in names, so every non-ASCII code
 * unit counts as a name character except the Unicode space separators.
 * `NaN` (reading past the end of the text) is not a name character.
 */
function isUnicodeNameUnit(code: number): boolean {
	if (!(code >= 128) || code === 0xa0 || code === 0x1680 || code === 0x202f || code === 0x205f) {
		return false;
	}
	if (code >= 0x2000 && code <= 0x200b) {
		return false;
	}
	return code !== 0x2028 && code !== 0x2029 && code !== 0x3000 && code !== 0xfeff;
}

/** First character of a name, function, sheet or table (`A`, `_`, `\`, any non-ASCII letter). */
export function isNameStart(code: number): boolean {
	if (code < 128) {
		return (ASCII[code] & START) !== 0;
	}
	return isUnicodeNameUnit(code);
}

/** Continuation character of a name (letters, digits, `_`, `.`, `?`, `\`). */
export function isNameChar(code: number): boolean {
	if (code < 128) {
		return (ASCII[code] & PART) !== 0;
	}
	return isUnicodeNameUnit(code);
}

/** Characters of a "word": name characters plus `$` (absolute reference markers). */
export function isWordChar(code: number): boolean {
	return code === CH_DOLLAR || isNameChar(code);
}

export function isDigit(code: number): boolean {
	return code >= CH_0 && code <= CH_9;
}

export function isLetter(code: number): boolean {
	const lower = code | 32;
	return lower >= 97 && lower <= 122;
}

/** Whitespace Excel accepts between formula tokens (space, tab, CR, LF). */
export function isSpace(code: number): boolean {
	return code === CH_SPACE || code === CH_LF || code === CH_CR || code === CH_TAB;
}
