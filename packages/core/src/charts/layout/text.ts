/**
 * Text measurement for chart layout. The default measurer uses Calibri advance
 * widths (the Office default font) so layout is deterministic without a DOM;
 * hosts with a canvas can inject exact measurements instead.
 */

export interface ChartFont {
	/** Primary font family name (first entry of the CSS stack). */
	family: string;
	/** Size in pixels. */
	size: number;
	bold?: boolean;
	italic?: boolean;
}

/** Width in pixels of a single line of text. */
export type ChartTextMeasurer = (text: string, font: ChartFont) => number;

/** Calibri advance widths for ASCII 32..126 in thousandths of an em. */
const CALIBRI_ASCII = [
	226, 326, 401, 498, 507, 715, 682, 221, 303, 303, 498, 498, 250, 306, 252, 386, 507, 507, 507,
	507, 507, 507, 507, 507, 507, 507, 268, 268, 498, 498, 498, 463, 894, 579, 544, 533, 615, 488,
	459, 631, 623, 252, 319, 520, 420, 855, 646, 662, 517, 673, 543, 459, 487, 642, 567, 890, 519,
	487, 468, 307, 386, 307, 498, 498, 291, 479, 525, 423, 525, 498, 305, 471, 525, 229, 239, 455,
	229, 799, 525, 527, 525, 525, 349, 391, 335, 525, 452, 715, 433, 453, 395, 312, 460, 312, 498,
];

/** Width relative to Calibri for common families (wider fonts need more room). */
const FAMILY_FACTORS: Array<[RegExp, number]> = [
	[/calibri|carlito/i, 1],
	[/aptos/i, 1.06],
	[/arial|helvetica|liberation sans|arimo/i, 1.1],
	[/segoe/i, 1.08],
	[/verdana|tahoma/i, 1.2],
	[/cambria|georgia|times|serif/i, 1.03],
	[/courier|consolas|mono/i, 1.18],
];

function familyFactor(family: string): number {
	for (const [re, f] of FAMILY_FACTORS) {
		if (re.test(family)) {
			return f;
		}
	}
	return 1.07;
}

function isWide(code: number): boolean {
	return (
		(code >= 0x1100 && code <= 0x115f) ||
		(code >= 0x2e80 && code <= 0xa4cf) ||
		(code >= 0xac00 && code <= 0xd7a3) ||
		(code >= 0xf900 && code <= 0xfaff) ||
		(code >= 0xfe30 && code <= 0xfe4f) ||
		(code >= 0xff00 && code <= 0xff60) ||
		(code >= 0xffe0 && code <= 0xffe6)
	);
}

/** Approximate advance width of one character in em. */
function charEm(code: number): number {
	if (code >= 32 && code <= 126) {
		return CALIBRI_ASCII[code - 32] / 1000;
	}
	if (isWide(code)) {
		return 1;
	}
	if (code >= 0xc0 && code <= 0x24f) {
		return 0.52;
	}
	if (code === 0x20ac || code === 0xa3 || code === 0xa5) {
		return 0.507;
	}
	return 0.55;
}

/** Deterministic Calibri-based measurement. */
export const defaultMeasureText: ChartTextMeasurer = (text, font) => {
	let em = 0;
	for (let i = 0; i < text.length; i++) {
		em += charEm(text.charCodeAt(i));
	}
	const weight = font.bold ? 1.05 : 1;
	return em * font.size * familyFactor(font.family) * weight;
};

/** Line height (baseline to baseline) for a font size in pixels. */
export function lineHeightFor(sizePx: number): number {
	return sizePx * 1.22;
}

/** Offset from the top of a line box to its baseline. */
export function baselineOffset(sizePx: number): number {
	return (lineHeightFor(sizePx) - sizePx) / 2 + sizePx * 0.8;
}

/** Shorten `text` with an ellipsis so it fits `maxWidth`. */
export function ellipsize(
	text: string,
	maxWidth: number,
	font: ChartFont,
	measure: ChartTextMeasurer,
): string {
	if (measure(text, font) <= maxWidth) {
		return text;
	}
	let lo = 0;
	let hi = text.length;
	while (lo < hi) {
		const mid = Math.ceil((lo + hi) / 2);
		if (measure(`${text.slice(0, mid)}…`, font) <= maxWidth) {
			lo = mid;
		} else {
			hi = mid - 1;
		}
	}
	return lo === 0 ? '' : `${text.slice(0, lo).trimEnd()}…`;
}

/**
 * Word-wrap `text` to `maxWidth`. Explicit newlines are kept. Words longer
 * than a line are broken by character. At most `maxLines` lines are returned;
 * the last one is ellipsized when text remains.
 */
export function wrapText(
	text: string,
	maxWidth: number,
	font: ChartFont,
	measure: ChartTextMeasurer,
	maxLines = Infinity,
): string[] {
	const out: string[] = [];
	const paragraphs = text.split('\n');
	for (let p = 0; p < paragraphs.length; p++) {
		const words = paragraphs[p].split(/(\s+)/).filter((w) => w.length > 0);
		let line = '';
		for (const word of words) {
			const candidate = line + word;
			if (line === '' || measure(candidate.trimEnd(), font) <= maxWidth) {
				if (line === '' && measure(word, font) > maxWidth && !/^\s+$/.test(word)) {
					// Break an over-long word by characters.
					let chunk = '';
					for (const ch of word) {
						if (chunk && measure(chunk + ch, font) > maxWidth) {
							out.push(chunk);
							chunk = ch;
						} else {
							chunk += ch;
						}
					}
					line = chunk;
				} else {
					line = /^\s+$/.test(word) && line === '' ? '' : candidate;
				}
			} else {
				out.push(line.trimEnd());
				line = /^\s+$/.test(word) ? '' : word;
			}
		}
		out.push(line.trimEnd());
	}
	if (out.length > maxLines) {
		const kept = out.slice(0, maxLines);
		const rest = out.slice(maxLines - 1).join(' ');
		kept[maxLines - 1] = ellipsize(rest, maxWidth, font, measure);
		return kept;
	}
	return out;
}

/** Width of the widest line. */
export function blockWidth(
	lines: readonly string[],
	font: ChartFont,
	measure: ChartTextMeasurer,
): number {
	let w = 0;
	for (const l of lines) {
		w = Math.max(w, measure(l, font));
	}
	return w;
}
