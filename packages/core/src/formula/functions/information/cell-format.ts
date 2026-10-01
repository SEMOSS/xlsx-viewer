/**
 * CELL("format" | "color" | "parentheses") support: derive Excel's short
 * format descriptors (`G`, `F2`, `,0`, `C2-`, `P0`, `S2`, `D1`...) from a
 * number format code.
 */

/** Split a format code into its `;`-separated sections (quotes, escapes and brackets respected). */
export function formatSections(code: string): string[] {
	const sections: string[] = [];
	let current = '';
	for (let i = 0; i < code.length; i++) {
		const ch = code[i];
		if (ch === '"') {
			const end = code.indexOf('"', i + 1);
			const stop = end < 0 ? code.length : end + 1;
			current += code.slice(i, stop);
			i = stop - 1;
		} else if (ch === '\\' || ch === '_' || ch === '*') {
			current += code.slice(i, i + 2);
			i++;
		} else if (ch === '[') {
			const end = code.indexOf(']', i + 1);
			const stop = end < 0 ? code.length : end + 1;
			current += code.slice(i, stop);
			i = stop - 1;
		} else if (ch === ';') {
			sections.push(current);
			current = '';
		} else {
			current += ch;
		}
	}
	sections.push(current);
	return sections;
}

/** A section with literals, escapes, padding / fill directives and bracket codes removed. */
function bare(section: string): string {
	return section.replace(/"[^"]*"?|\\.|_.|\*.|\[[^\]]*\]?/g, '');
}

const COLOR = /\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]/i;

function decimals(section: string): number {
	const m = /\.([0#?]+)/.exec(bare(section));
	return m ? m[1].length : 0;
}

function dateCode(lower: string): string {
	if (/[dy]/.test(lower)) {
		const hasDay = lower.includes('d');
		const hasYear = lower.includes('y');
		if (/mmm/.test(lower)) {
			if (hasDay && hasYear) {
				return 'D1';
			}
			return hasDay ? 'D2' : 'D3';
		}
		return hasYear ? 'D4' : 'D5';
	}
	const ampm = /am\/pm|a\/p/.test(lower);
	const seconds = lower.includes('s');
	if (ampm) {
		return seconds ? 'D6' : 'D7';
	}
	return seconds ? 'D8' : 'D9';
}

/** The negative values of the format are shown in color. */
export function hasNegativeColor(code: string | undefined): boolean {
	if (!code) {
		return false;
	}
	const sections = formatSections(code);
	return COLOR.test(sections.length > 1 ? sections[1] : sections[0]);
}

/** Positive (or all) values are shown in parentheses. */
export function hasParentheses(code: string | undefined): boolean {
	if (!code) {
		return false;
	}
	return bare(formatSections(code)[0]).includes('(');
}

/** Excel's CELL("format") descriptor of a number format code. */
export function cellFormatCode(code: string | undefined): string {
	if (!code || /^general$/i.test(code.trim())) {
		return 'G';
	}
	const first = formatSections(code)[0];
	const plain = bare(first);
	const lower = plain.toLowerCase();
	let result: string;
	if (/general/.test(lower) || plain.includes('@') || /[#0?]\s*\/\s*[#0?\d]/.test(plain)) {
		result = 'G';
	} else if (/[dmyhs]/.test(lower)) {
		result = dateCode(lower);
	} else if (plain.includes('%')) {
		result = `P${decimals(first)}`;
	} else if (/e[+-]/.test(lower)) {
		result = `S${decimals(first)}`;
	} else if (/[$\u20ac\u00a3\u00a5]/.test(first.replace(/\[\$-[0-9A-Fa-f]+\]/g, ''))) {
		result = `C${decimals(first)}`;
	} else if (/[#0?],[#0?]/.test(plain)) {
		result = `,${decimals(first)}`;
	} else {
		result = `F${decimals(first)}`;
	}
	if (result.startsWith('D') || result === 'G') {
		return result;
	}
	return result + (hasNegativeColor(code) ? '-' : '') + (hasParentheses(code) ? '()' : '');
}
