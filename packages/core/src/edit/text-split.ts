/**
 * Text to Columns parsing: delimited and fixed-width splitting of a line of
 * text into fields, the way Excel's Convert Text to Columns Wizard does.
 */

export interface DelimitedSplitOptions {
	tab: boolean;
	semicolon: boolean;
	comma: boolean;
	space: boolean;
	/** One extra delimiter character ("Other"); only the first character counts. */
	other?: string;
	/** Runs of delimiters separate a single pair of fields. */
	treatConsecutiveAsOne: boolean;
	/** Fields starting with this character are quoted: delimiters inside them are text. */
	textQualifier: '"' | "'" | null;
}

/** How to split each line: by delimiters or at fixed character positions. */
export interface TextSplitOptions extends Partial<DelimitedSplitOptions> {
	mode: 'delimited' | 'fixed';
	/** Fixed width: column break positions (characters from the start of the line). */
	breaks?: number[];
}

function delimiterSet(opts: DelimitedSplitOptions): Set<string> {
	const set = new Set<string>();
	if (opts.tab) {
		set.add('\t');
	}
	if (opts.semicolon) {
		set.add(';');
	}
	if (opts.comma) {
		set.add(',');
	}
	if (opts.space) {
		set.add(' ');
	}
	if (opts.other) {
		set.add(opts.other.charAt(0));
	}
	return set;
}

/**
 * Split `text` at the chosen delimiters. A field that starts with the text
 * qualifier runs to the closing qualifier (a doubled qualifier stands for
 * one), so delimiters inside it are kept. With `treatConsecutiveAsOne`,
 * a run of delimiters ends a field only once.
 */
export function splitDelimitedText(text: string, opts: DelimitedSplitOptions): string[] {
	const delimiters = delimiterSet(opts);
	const q = opts.textQualifier;
	const out: string[] = [];
	let field = '';
	let fieldStart = true;
	let afterDelimiter = false;
	let i = 0;
	while (i < text.length) {
		const ch = text[i];
		if (fieldStart && q !== null && ch === q) {
			i++;
			while (i < text.length) {
				if (text[i] === q) {
					if (text[i + 1] !== q) {
						i++;
						break;
					}
					i++;
				}
				field += text[i++];
			}
			fieldStart = false;
			afterDelimiter = false;
			continue;
		}
		i++;
		if (delimiters.has(ch)) {
			if (!(opts.treatConsecutiveAsOne && afterDelimiter)) {
				out.push(field);
				field = '';
			}
			fieldStart = true;
			afterDelimiter = true;
			continue;
		}
		field += ch;
		fieldStart = false;
		afterDelimiter = false;
	}
	out.push(field);
	return out;
}

/** Sorted, distinct, positive integer break positions. */
export function normalizeBreaks(breaks: readonly number[]): number[] {
	const valid = breaks.filter((b) => Number.isInteger(b) && b > 0);
	return [...new Set(valid)].sort((a, b) => a - b);
}

/**
 * Split `text` at fixed character positions. Fields lose their trailing
 * spaces; fields starting past the end of the text are not returned.
 */
export function splitFixedWidthText(text: string, breaks: readonly number[]): string[] {
	const out: string[] = [];
	let start = 0;
	for (const end of [...normalizeBreaks(breaks), text.length]) {
		if (end <= start && out.length > 0) {
			continue;
		}
		out.push(text.slice(start, Math.max(start, end)).replace(/ +$/, ''));
		start = end;
		if (start >= text.length) {
			break;
		}
	}
	return out;
}

/**
 * Guess fixed-width column breaks: a break goes where a column of spaces
 * shared by every non-empty line ends and text starts again.
 */
export function guessFixedWidthBreaks(lines: readonly string[]): number[] {
	const used = lines.filter((l) => l.trim() !== '');
	if (used.length === 0) {
		return [];
	}
	const width = Math.max(...used.map((l) => l.length));
	const blank: boolean[] = [];
	for (let p = 0; p < width; p++) {
		blank.push(used.every((l) => p >= l.length || l[p] === ' '));
	}
	const breaks: number[] = [];
	let seenText = !blank[0];
	for (let p = 1; p < width; p++) {
		if (blank[p - 1] && !blank[p] && seenText) {
			breaks.push(p);
		}
		seenText ||= !blank[p];
	}
	return breaks;
}

function delimitedOptions(o: TextSplitOptions): DelimitedSplitOptions {
	return {
		tab: o.tab ?? false,
		semicolon: o.semicolon ?? false,
		comma: o.comma ?? false,
		space: o.space ?? false,
		other: o.other,
		treatConsecutiveAsOne: o.treatConsecutiveAsOne ?? false,
		textQualifier: o.textQualifier === undefined ? '"' : o.textQualifier,
	};
}

/** Split one line with the chosen mode. */
export function splitTextLine(text: string, options: TextSplitOptions): string[] {
	return options.mode === 'fixed'
		? splitFixedWidthText(text, options.breaks ?? [])
		: splitDelimitedText(text, delimitedOptions(options));
}

/** Split every line (the wizard's data preview); empty lines give no fields. */
export function parseTextToColumns(
	lines: readonly string[],
	options: TextSplitOptions,
): string[][] {
	return lines.map((line) => (line === '' ? [] : splitTextLine(line, options)));
}
