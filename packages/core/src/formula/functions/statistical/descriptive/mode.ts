import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { columnOf, define, numbersOf } from '../args';

/**
 * Values occurring most often (at least twice), in order of first
 * occurrence; empty when every value is unique.
 */
export function modes(xs: readonly number[]): number[] {
	const counts = new Map<number, number>();
	let best = 1;
	for (const x of xs) {
		const c = (counts.get(x) ?? 0) + 1;
		counts.set(x, c);
		if (c > best) {
			best = c;
		}
	}
	if (best < 2) {
		return [];
	}
	// Map iteration follows insertion order, which is first-occurrence order.
	const out: number[] = [];
	for (const [value, c] of counts) {
		if (c === best) {
			out.push(value);
		}
	}
	return out;
}

function modeSingle(name: string): FunctionDefinition {
	return define({
		name,
		minArgs: 1,
		maxArgs: Infinity,
		params: ['any'],
		impl: (args, ctx) => {
			const xs = numbersOf(args, ctx);
			if (isCellError(xs)) {
				return xs;
			}
			const m = modes(xs);
			return m.length === 0 ? ERRORS.NA : m[0];
		},
	});
}

export const MODE_FUNCTIONS: FunctionDefinition[] = [
	modeSingle('MODE'),
	modeSingle('MODE.SNGL'),
	define({
		name: 'MODE.MULT',
		minArgs: 1,
		maxArgs: Infinity,
		params: ['any'],
		returnsArray: true,
		impl: (args, ctx) => {
			const xs = numbersOf(args, ctx);
			if (isCellError(xs)) {
				return xs;
			}
			const m = modes(xs);
			return m.length === 0 ? ERRORS.NA : columnOf(m);
		},
	}),
];
