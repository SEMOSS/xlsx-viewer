/**
 * Format code parser: compiles a code into a {@link ParsedNumberFormat} (sections,
 * render programs, section-selection rules) and caches the result, so each
 * distinct code is parsed once and formatting a value stays cheap.
 */
import { compileDate } from './compile-date';
import { compileFraction, compileNumber, compileSimple, findFractionSlash } from './compile-number';
import { SYSTEM_LONG_DATE_FORMAT, SYSTEM_TIME_FORMAT } from './locale-en-us';
import { tokenizeFormat } from './tokenizer';
import type {
	NumberFormatCondition,
	NumberFormatSection,
	NumberFormatToken,
	ParsedNumberFormat,
	SectionKind,
	SectionRule,
} from './types';

const GENERAL_RE = /^general$/i;

const LAYOUT_DEFAULTS = {
	hasFill: false,
	intSlots: 0,
	decSlots: 0,
	expSlots: 0,
	grouping: false,
	scale: 0,
	percent: false,
	numSlots: 0,
	denSlots: 0,
	denFixed: 0,
	hasDate: false,
	hasTime: false,
	hour12: false,
	subsecDigits: 0,
	elapsedUnit: null,
} as const;

function detectKind(body: readonly NumberFormatToken[], isText: boolean): SectionKind {
	if (isText) {
		return 'text';
	}
	if (body.length === 0) {
		return 'empty';
	}
	let general = false;
	let strongDate = false;
	let weakDate = false;
	let digits = false;
	let exponent = false;
	for (const t of body) {
		switch (t.type) {
			case 'general':
				general = true;
				break;
			case 'digit':
				digits = true;
				break;
			case 'exponent':
				exponent ||= digits;
				break;
			case 'ampm':
			case 'elapsed':
				strongDate = true;
				break;
			case 'date':
				if (t.part === 'e' || t.part === 'b' || t.part === 'g') {
					weakDate = true;
				} else {
					strongDate = true;
				}
				break;
			default:
				break;
		}
	}
	if (general) {
		return 'general';
	}
	// `e`, `b` and `g` are date codes only in sections without digit placeholders.
	if (strongDate || (weakDate && !digits)) {
		return 'date';
	}
	if (findFractionSlash(body) >= 0) {
		return 'fraction';
	}
	return exponent ? 'scientific' : 'number';
}

/** Compile one section's tokens. `isText` marks the text (4th or trailing `@`) section. */
export function compileSection(
	tokens: readonly NumberFormatToken[],
	isText: boolean,
): NumberFormatSection {
	let color: string | undefined;
	let colorIndex: number | undefined;
	let condition: NumberFormatCondition | undefined;
	let lcid: number | undefined;
	let currencySymbol: string | undefined;
	let system: 'date' | 'time' | undefined;
	const body: NumberFormatToken[] = [];
	for (const t of tokens) {
		switch (t.type) {
			case 'color':
				if (color === undefined && colorIndex === undefined) {
					color = t.hex;
					colorIndex = t.index;
				}
				break;
			case 'condition':
				condition ??= { op: t.op, value: t.value };
				break;
			case 'locale':
				lcid ??= t.lcid;
				system ??= t.system;
				if (t.symbol !== '') {
					currencySymbol ??= t.symbol;
					body.push({ type: 'literal', text: t.symbol, raw: t.raw });
				}
				break;
			case 'ignored':
				break;
			default:
				body.push(t);
		}
	}
	// System long date / time tags replace the pattern written after them.
	const source =
		system !== undefined && !isText
			? tokenizeFormat(system === 'date' ? SYSTEM_LONG_DATE_FORMAT : SYSTEM_TIME_FORMAT)[0]
			: body;
	// A section holding only tags (`[Red][<=100]`) shows the value as General (text as-is).
	const implicit =
		source.length === 0 &&
		tokens.some((t) => t.type === 'color' || t.type === 'condition' || t.type === 'locale');
	const kind = implicit && !isText ? 'general' : detectKind(source, isText);
	const base = {
		...LAYOUT_DEFAULTS,
		tokens,
		kind,
		color,
		colorIndex,
		condition,
		lcid,
		currencySymbol,
		hasFill: source.some((t) => t.type === 'fill'),
	};
	switch (kind) {
		case 'number':
		case 'scientific':
			return { ...base, ...compileNumber(source, kind === 'scientific') };
		case 'fraction':
			return { ...base, ...compileFraction(source, findFractionSlash(source)) };
		case 'date':
			return { ...base, ...compileDate(source) };
		case 'general':
		case 'text':
			if (implicit) {
				return { ...base, items: [isText ? { k: 'text' } : { k: 'general' }] };
			}
			return { ...base, items: compileSimple(source, kind === 'text') };
		default:
			return { ...base, items: [] };
	}
}

/** The section only ever receives negative numbers (so it prints its own sign). */
function negativeOnly(c: NumberFormatCondition): boolean {
	switch (c.op) {
		case '<':
		case '<=':
			return c.value <= 0;
		case '=':
			return c.value < 0;
		default:
			return false;
	}
}

/**
 * Selection rules. Without conditions: one section serves every number (with
 * a minus sign); two are `>= 0` / `< 0`; three are `> 0` / `< 0` / `= 0`, and
 * the `< 0` section omits the minus sign. With conditions in the first two
 * sections, missing ones default like Excel (`> 0` / `>= 0` first, `< 0`
 * second when a third exists) and the third section takes the rest.
 */
function buildRules(numeric: readonly NumberFormatSection[], conditional: boolean): SectionRule[] {
	const n = numeric.length;
	if (!conditional) {
		return numeric.map((_, i) => ({ dropSign: i === 1 }));
	}
	return numeric.map((s, i) => {
		let condition = i < 2 ? s.condition : undefined;
		if (condition === undefined && i === 0 && n >= 2) {
			condition = { op: n >= 3 ? '>' : '>=', value: 0 };
		} else if (condition === undefined && i === 1 && n >= 3) {
			condition = { op: '<', value: 0 };
		}
		return condition ? { condition, dropSign: negativeOnly(condition) } : { dropSign: false };
	});
}

function buildFormat(code: string): ParsedNumberFormat {
	const isGeneral = code === '' || GENERAL_RE.test(code);
	const tokenSections = tokenizeFormat(isGeneral ? 'General' : code).slice(0, 4);
	const count = tokenSections.length;
	let textIndex = -1;
	if (count === 4) {
		textIndex = 3;
	} else if (tokenSections[count - 1].some((t) => t.type === 'text')) {
		textIndex = count - 1;
	}
	const sections = tokenSections.map((tokens, i) => compileSection(tokens, i === textIndex));
	const numeric = textIndex < 0 ? sections : sections.slice(0, textIndex);
	const conditional = numeric.slice(0, 2).some((s) => s.condition !== undefined);
	const first = numeric[0];
	const isDate = first !== undefined && first.kind === 'date';
	return {
		code,
		sections,
		numeric,
		text: textIndex < 0 ? null : sections[textIndex],
		rules: buildRules(numeric, conditional),
		conditional,
		isGeneral,
		isDate,
		isTimeOnly: isDate && first.hasTime && !first.hasDate,
	};
}

const CACHE_LIMIT = 4096;
const cache = new Map<string, ParsedNumberFormat>();

/**
 * Parse a format code into its structured representation. Results are
 * cached per code (bounded; oldest entries are evicted first).
 */
export function parseFormatCode(formatCode: string): ParsedNumberFormat {
	const code = typeof formatCode === 'string' ? formatCode : '';
	let parsed = cache.get(code);
	if (parsed === undefined) {
		parsed = buildFormat(code);
		if (cache.size >= CACHE_LIMIT) {
			const oldest = cache.keys().next();
			if (!oldest.done) {
				cache.delete(oldest.value);
			}
		}
		cache.set(code, parsed);
	}
	return parsed;
}

/** Forget all cached parses. */
export function clearFormatCache(): void {
	cache.clear();
}
