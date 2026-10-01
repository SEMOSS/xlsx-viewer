/**
 * Types of the number format engine: the public result and option shapes,
 * the token stream produced by the tokenizer and the compiled (cached)
 * representation of a format code.
 */

/** Options for {@link formatValue}. */
export interface FormatOptions {
	/** Workbook uses the 1904 date system. */
	date1904?: boolean;
	/**
	 * Character budget of the General format (digits, decimal point and
	 * exponent; the minus sign of a negative number is not counted). Drives
	 * Excel's width-dependent rounding and switch to scientific notation.
	 * Defaults to 11, the classic General width (what `TEXT(x, "General")`
	 * produces and what a standard-width column shows for typed numbers).
	 */
	generalMaxChars?: number;
	/**
	 * Workbook palette (64 `RRGGBB` or `AARRGGBB` entries) used by `[ColorN]`
	 * tags. Defaults to Excel's legacy indexed palette.
	 */
	indexedColors?: readonly string[];
}

export interface FormattedValue {
	/** Display text (without the repeat-fill padding, see `fill`). */
	text: string;
	/** Color from a `[Red]` / `[Color10]` section, as `#RRGGBB`. */
	color?: string;
	/**
	 * Repeat-fill from a `*x` token (accounting formats): the cell shows
	 * `before`, then `char` repeated to fill the width, then `after`.
	 * `text` is always `before + after`.
	 */
	fill?: { before: string; char: string; after: string };
	/** The value cannot be shown in this format (e.g. a negative date): Excel shows `#####`. */
	invalid?: boolean;
}

export type FormatCategory =
	| 'general'
	| 'number'
	| 'currency'
	| 'accounting'
	| 'date'
	| 'time'
	| 'percentage'
	| 'fraction'
	| 'scientific'
	| 'text'
	| 'special'
	| 'custom';

/** Digit placeholder characters. */
export type Placeholder = '0' | '#' | '?';

export type CompareOp = '<' | '<=' | '>' | '>=' | '=' | '<>';

/** A section condition such as `[>=100]`. */
export interface NumberFormatCondition {
	readonly op: CompareOp;
	readonly value: number;
}

/** Letters that start date/time codes (`e` year, `b` Buddhist year, `g` era). */
export type DateLetter = 'y' | 'm' | 'd' | 'h' | 's' | 'e' | 'b' | 'g';

export type TimeUnit = 'h' | 'm' | 's';

/**
 * One lexical element of a format code. `raw` is the exact source text, so
 * joining the raws of a section reproduces it (used by `adjustDecimals`).
 */
export type NumberFormatToken =
	/** Quoted text, an escaped `\x` character or a character shown as-is. */
	| { readonly type: 'literal'; readonly text: string; readonly raw: string }
	/** Unquoted digit 1-9 (literal, or a fixed fraction denominator). */
	| { readonly type: 'numeral'; readonly text: string; readonly raw: string }
	/** `_x`: a space as wide as `x`. */
	| { readonly type: 'space'; readonly char: string; readonly raw: string }
	/** `*x`: repeat `x` to fill the cell. */
	| { readonly type: 'fill'; readonly char: string; readonly raw: string }
	| { readonly type: 'digit'; readonly char: Placeholder; readonly raw: string }
	| { readonly type: 'point'; readonly raw: string }
	| { readonly type: 'comma'; readonly raw: string }
	| { readonly type: 'percent'; readonly raw: string }
	| { readonly type: 'slash'; readonly raw: string }
	| {
			readonly type: 'exponent';
			readonly upper: boolean;
			readonly plus: boolean;
			readonly raw: string;
	  }
	/** `@` text placeholder. */
	| { readonly type: 'text'; readonly raw: string }
	| { readonly type: 'general'; readonly raw: string }
	| {
			readonly type: 'date';
			readonly part: DateLetter;
			readonly length: number;
			readonly raw: string;
	  }
	/** `AM/PM` or `A/P` with the strings to print (case as written). */
	| { readonly type: 'ampm'; readonly am: string; readonly pm: string; readonly raw: string }
	/** `[h]`, `[mm]`, `[ss]` elapsed time. */
	| {
			readonly type: 'elapsed';
			readonly unit: TimeUnit;
			readonly length: number;
			readonly raw: string;
	  }
	| { readonly type: 'color'; readonly hex?: string; readonly index?: number; readonly raw: string }
	| {
			readonly type: 'condition';
			readonly op: CompareOp;
			readonly value: number;
			readonly raw: string;
	  }
	/** `[$sym-lcid]`: currency symbol and / or locale (`system` for F800 / F400 / x-sysdate). */
	| {
			readonly type: 'locale';
			readonly symbol: string;
			readonly lcid?: number;
			readonly system?: 'date' | 'time';
			readonly raw: string;
	  }
	/** Recognised but irrelevant for en-US output (`[DBNum1]`, `B2`, dangling escapes). */
	| { readonly type: 'ignored'; readonly raw: string };

export type DateField =
	| 'year2'
	| 'year4'
	| 'buddhist2'
	| 'buddhist4'
	| 'era'
	| 'month'
	| 'monthPad'
	| 'monthShort'
	| 'monthLong'
	| 'monthLetter'
	| 'day'
	| 'dayPad'
	| 'weekdayShort'
	| 'weekdayLong'
	| 'hour'
	| 'hourPad'
	| 'minute'
	| 'minutePad'
	| 'second'
	| 'secondPad';

/**
 * Compiled render program of a section. Integer-like slots count `pos` from
 * the right (0 = units); decimal-like slots count `idx` from the left.
 */
export type RenderItem =
	| { readonly k: 'lit'; readonly text: string }
	| { readonly k: 'fill'; readonly char: string }
	| { readonly k: 'int'; readonly pos: number; readonly ph: Placeholder }
	/** Integer digits of a format without integer placeholders (`.00`). */
	| { readonly k: 'intAll' }
	| { readonly k: 'point' }
	| { readonly k: 'dec'; readonly idx: number; readonly ph: Placeholder }
	| { readonly k: 'expMark'; readonly upper: boolean; readonly plus: boolean }
	| { readonly k: 'exp'; readonly pos: number; readonly ph: Placeholder }
	| { readonly k: 'whole'; readonly pos: number; readonly ph: Placeholder }
	| { readonly k: 'num'; readonly pos: number; readonly ph: Placeholder }
	| { readonly k: 'slash' }
	| { readonly k: 'den'; readonly idx: number; readonly ph: Placeholder }
	| { readonly k: 'denFixed'; readonly text: string }
	| { readonly k: 'general' }
	| { readonly k: 'text' }
	| { readonly k: 'date'; readonly field: DateField }
	| { readonly k: 'ampm'; readonly am: string; readonly pm: string }
	| { readonly k: 'elapsed'; readonly unit: TimeUnit; readonly len: number }
	| { readonly k: 'subsec'; readonly len: number };

export type SectionKind =
	| 'empty'
	| 'general'
	| 'number'
	| 'scientific'
	| 'fraction'
	| 'date'
	| 'text';

/** One compiled `;`-separated section. Layout fields not used by its kind are 0 / false. */
export interface NumberFormatSection {
	/** Tokens as written, brackets included. */
	readonly tokens: readonly NumberFormatToken[];
	readonly kind: SectionKind;
	readonly items: readonly RenderItem[];
	/** `#RRGGBB` of a named color tag (`[Red]`). */
	readonly color?: string;
	/** N of a `[ColorN]` tag (1-56, resolved against the palette when formatting). */
	readonly colorIndex?: number;
	readonly condition?: NumberFormatCondition;
	/** Locale id of a `[$-409]` style tag. */
	readonly lcid?: number;
	/** Currency symbol of a `[$€-407]` style tag. */
	readonly currencySymbol?: string;
	readonly hasFill: boolean;
	/** Integer placeholders (number / scientific mantissa) or whole-part placeholders (fraction). */
	readonly intSlots: number;
	readonly decSlots: number;
	readonly expSlots: number;
	readonly grouping: boolean;
	/** Decimal exponent applied before rounding: +2 per `%`, -3 per scaling comma. */
	readonly scale: number;
	readonly percent: boolean;
	readonly numSlots: number;
	readonly denSlots: number;
	/** Fixed denominator (`# ?/8`), 0 when the denominator is computed. */
	readonly denFixed: number;
	readonly hasDate: boolean;
	readonly hasTime: boolean;
	readonly hour12: boolean;
	readonly subsecDigits: number;
	/** Largest elapsed unit (`[h]` > `[m]` > `[s]`), shown as a running total. */
	readonly elapsedUnit: TimeUnit | null;
}

/** How a numeric section is chosen and whether it prints the minus sign itself. */
export interface SectionRule {
	/** Effective condition; absent means "any remaining value". */
	readonly condition?: NumberFormatCondition;
	/** The section only receives negative values, so the minus sign is left to it. */
	readonly dropSign: boolean;
}

/** Structured, cached representation of a format code (see `parseFormatCode`). */
export interface ParsedNumberFormat {
	readonly code: string;
	/** All sections as written (at most four). */
	readonly sections: readonly NumberFormatSection[];
	/** Sections that format numbers, in order (positive, negative, zero). */
	readonly numeric: readonly NumberFormatSection[];
	/** The text section: the 4th section, or a trailing section containing `@`. */
	readonly text: NumberFormatSection | null;
	readonly rules: readonly SectionRule[];
	/** The first two numeric sections carry explicit conditions. */
	readonly conditional: boolean;
	/** The whole code is `General` (or empty). */
	readonly isGeneral: boolean;
	/** The first numeric section is a date / time section. */
	readonly isDate: boolean;
	/** The first numeric section shows only time fields. */
	readonly isTimeOnly: boolean;
}
