import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { integerArg, textArg } from '../math/args';
import { alias, limitText } from './text-utils';

/**
 * Positions and lengths count UTF-16 code units, like Excel. The `B`
 * variants (LEFTB, MIDB, ...) count bytes only when a DBCS language is the
 * default editing language; in other locales they equal the plain
 * functions, which is the behaviour implemented here.
 */

/** Optional character count (default 1); negative counts are #VALUE!. */
function countArg(
	args: readonly FunctionArg[],
	index: number,
	ctx: FunctionContext,
): number | CellError {
	if (index >= args.length) {
		return 1;
	}
	const n = integerArg(args[index], ctx);
	if (typeof n !== 'number') {
		return n;
	}
	return n < 0 ? ERRORS.VALUE : n;
}

function edge(name: string, fromRight: boolean): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 2,
		params: ['value', 'value'],
		impl: (args, ctx) => {
			const text = textArg(args[0]);
			if (isCellError(text)) {
				return text;
			}
			const n = countArg(args, 1, ctx);
			if (typeof n !== 'number') {
				return n;
			}
			if (fromRight) {
				return n === 0 ? '' : text.slice(Math.max(0, text.length - n));
			}
			return text.slice(0, n);
		},
	};
}

const MID: FunctionDefinition = {
	name: 'MID',
	minArgs: 3,
	maxArgs: 3,
	params: ['value', 'value', 'value'],
	impl: (args, ctx) => {
		const text = textArg(args[0]);
		if (isCellError(text)) {
			return text;
		}
		const start = integerArg(args[1], ctx);
		if (typeof start !== 'number') {
			return start;
		}
		const n = integerArg(args[2], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		if (start < 1 || n < 0) {
			return ERRORS.VALUE;
		}
		return text.substr(start - 1, n);
	},
};

const REPLACE: FunctionDefinition = {
	name: 'REPLACE',
	minArgs: 4,
	maxArgs: 4,
	params: ['value', 'value', 'value', 'value'],
	impl: (args, ctx) => {
		const text = textArg(args[0]);
		if (isCellError(text)) {
			return text;
		}
		const start = integerArg(args[1], ctx);
		if (typeof start !== 'number') {
			return start;
		}
		const n = integerArg(args[2], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		const replacement = textArg(args[3]);
		if (isCellError(replacement)) {
			return replacement;
		}
		if (start < 1 || n < 0) {
			return ERRORS.VALUE;
		}
		const at = Math.min(start - 1, text.length);
		return limitText(text.slice(0, at) + replacement + text.slice(at + n));
	},
};

/** Replace the `instance`-th (1-based) occurrence of `find`, or every occurrence when omitted. */
export function substitute(
	text: string,
	find: string,
	replacement: string,
	instance?: number,
): string {
	if (find === '') {
		return text;
	}
	if (instance === undefined) {
		return text.split(find).join(replacement);
	}
	let from = 0;
	for (let i = 1; ; i++) {
		const at = text.indexOf(find, from);
		if (at < 0) {
			return text;
		}
		if (i === instance) {
			return text.slice(0, at) + replacement + text.slice(at + find.length);
		}
		from = at + find.length;
	}
}

const SUBSTITUTE: FunctionDefinition = {
	name: 'SUBSTITUTE',
	minArgs: 3,
	maxArgs: 4,
	params: ['value', 'value', 'value', 'value'],
	impl: (args, ctx) => {
		const parts: string[] = [];
		for (let i = 0; i < 3; i++) {
			const t = textArg(args[i]);
			if (isCellError(t)) {
				return t;
			}
			parts.push(t);
		}
		let instance: number | undefined;
		if (args.length > 3) {
			const n = integerArg(args[3], ctx);
			if (typeof n !== 'number') {
				return n;
			}
			if (n < 1) {
				return ERRORS.VALUE;
			}
			instance = n;
		}
		return limitText(substitute(parts[0], parts[1], parts[2], instance));
	},
};

const LEFT = edge('LEFT', false);
const RIGHT = edge('RIGHT', true);

export const SUBSTRING_FUNCTIONS: FunctionDefinition[] = [
	LEFT,
	alias(LEFT, 'LEFTB'),
	RIGHT,
	alias(RIGHT, 'RIGHTB'),
	MID,
	alias(MID, 'MIDB'),
	REPLACE,
	alias(REPLACE, 'REPLACEB'),
	SUBSTITUTE,
];
