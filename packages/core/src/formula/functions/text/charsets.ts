/**
 * Character tables for CHAR / CODE (Windows-1252, the ANSI code page of
 * Western-locale Excel) and ASC / DBCS (half-width and full-width forms).
 * Everything is written as code points so the source stays ASCII.
 */

/** Unicode code points of Windows-1252 bytes 0x80-0x9F (undefined bytes map to themselves). */
const CP1252_HIGH = [
	0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152,
	0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122,
	0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

const CP1252_REVERSE = new Map<number, number>(CP1252_HIGH.map((cp, i) => [cp, 0x80 + i]));

/** Character for a Windows-1252 byte (1-255). */
export function cp1252Char(byte: number): string {
	return String.fromCharCode(byte >= 0x80 && byte <= 0x9f ? CP1252_HIGH[byte - 0x80] : byte);
}

/** Windows-1252 byte of a UTF-16 code unit; unmappable characters give 63 (`?`), like Excel. */
export function cp1252Code(unit: number): number {
	if (unit < 0x80 || (unit >= 0xa0 && unit <= 0xff)) {
		return unit;
	}
	return CP1252_REVERSE.get(unit) ?? 63;
}

/** Full-width forms of the half-width katakana block U+FF61-U+FF9F, in order. */
const HALF_KATAKANA_TO_FULL = [
	0x3002, 0x300c, 0x300d, 0x3001, 0x30fb, 0x30f2, 0x30a1, 0x30a3, 0x30a5, 0x30a7, 0x30a9, 0x30e3,
	0x30e5, 0x30e7, 0x30c3, 0x30fc, 0x30a2, 0x30a4, 0x30a6, 0x30a8, 0x30aa, 0x30ab, 0x30ad, 0x30af,
	0x30b1, 0x30b3, 0x30b5, 0x30b7, 0x30b9, 0x30bb, 0x30bd, 0x30bf, 0x30c1, 0x30c4, 0x30c6, 0x30c8,
	0x30ca, 0x30cb, 0x30cc, 0x30cd, 0x30ce, 0x30cf, 0x30d2, 0x30d5, 0x30d8, 0x30db, 0x30de, 0x30df,
	0x30e0, 0x30e1, 0x30e2, 0x30e4, 0x30e6, 0x30e8, 0x30e9, 0x30ea, 0x30eb, 0x30ec, 0x30ed, 0x30ef,
	0x30f3, 0x309b, 0x309c,
];
const HALF_KATAKANA_START = 0xff61;
const HALF_VOICED_MARK = 0xff9e;
const HALF_SEMI_VOICED_MARK = 0xff9f;

/** Full-width katakana that take a voiced (dakuten) form at code point + 1. */
const VOICEABLE = new Set([
	0x30ab, 0x30ad, 0x30af, 0x30b1, 0x30b3, 0x30b5, 0x30b7, 0x30b9, 0x30bb, 0x30bd, 0x30bf, 0x30c1,
	0x30c4, 0x30c6, 0x30c8, 0x30cf, 0x30d2, 0x30d5, 0x30d8, 0x30db,
]);
/** Full-width katakana that take a semi-voiced (handakuten) form at code point + 2. */
const SEMI_VOICEABLE = new Set([0x30cf, 0x30d2, 0x30d5, 0x30d8, 0x30db]);
const KATAKANA_U = 0x30a6;
const KATAKANA_VU = 0x30f4;

const FULL_TO_HALF = new Map<number, string>();
HALF_KATAKANA_TO_FULL.forEach((full, i) => {
	if (!FULL_TO_HALF.has(full)) {
		FULL_TO_HALF.set(full, String.fromCharCode(HALF_KATAKANA_START + i));
	}
});
for (const base of VOICEABLE) {
	FULL_TO_HALF.set(base + 1, FULL_TO_HALF.get(base)! + String.fromCharCode(HALF_VOICED_MARK));
}
for (const base of SEMI_VOICEABLE) {
	FULL_TO_HALF.set(base + 2, FULL_TO_HALF.get(base)! + String.fromCharCode(HALF_SEMI_VOICED_MARK));
}
FULL_TO_HALF.set(
	KATAKANA_VU,
	FULL_TO_HALF.get(KATAKANA_U)! + String.fromCharCode(HALF_VOICED_MARK),
);

/** ASC: full-width (double-byte) characters to half-width. */
export function toHalfWidth(text: string): string {
	let out = '';
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i);
		if (code === 0x3000) {
			out += ' ';
		} else if (code >= 0xff01 && code <= 0xff5e) {
			out += String.fromCharCode(code - 0xfee0);
		} else {
			out += FULL_TO_HALF.get(code) ?? text[i];
		}
	}
	return out;
}

/** DBCS: half-width (single-byte) characters to full-width, merging voiced-sound marks. */
export function toFullWidth(text: string): string {
	let out = '';
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i);
		if (code === 0x20) {
			out += String.fromCharCode(0x3000);
		} else if (code >= 0x21 && code <= 0x7e) {
			out += String.fromCharCode(code + 0xfee0);
		} else if (
			code >= HALF_KATAKANA_START &&
			code < HALF_KATAKANA_START + HALF_KATAKANA_TO_FULL.length
		) {
			let full = HALF_KATAKANA_TO_FULL[code - HALF_KATAKANA_START];
			const next = text.charCodeAt(i + 1);
			if (next === HALF_VOICED_MARK && (VOICEABLE.has(full) || full === KATAKANA_U)) {
				full = full === KATAKANA_U ? KATAKANA_VU : full + 1;
				i++;
			} else if (next === HALF_SEMI_VOICED_MARK && SEMI_VOICEABLE.has(full)) {
				full += 2;
				i++;
			}
			out += String.fromCharCode(full);
		} else {
			out += text[i];
		}
	}
	return out;
}
