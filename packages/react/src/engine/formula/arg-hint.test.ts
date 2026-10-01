import type { FunctionInfo } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { buildArgHint } from './arg-hint';

const fn = (name: string, args: Array<[string, ('opt' | 'rep')[]?]>): FunctionInfo => ({
	name,
	category: 'Math & Trig',
	description: '',
	args: args.map(([n, flags = []]) => ({
		name: n,
		description: '',
		optional: flags.includes('opt'),
		repeating: flags.includes('rep'),
	})),
});

const SUM = fn('SUM', [['number1'], ['number2', ['opt', 'rep']]]);
const SUMIFS = fn('SUMIFS', [
	['sum_range'],
	['criteria_range1'],
	['criteria1'],
	['criteria_range2', ['opt', 'rep']],
	['criteria2', ['opt', 'rep']],
]);
const LAMBDA = fn('LAMBDA', [['parameter1', ['opt', 'rep']], ['calculation']]);
const ABS = fn('ABS', [['number']]);

const text = (h: ReturnType<typeof buildArgHint>) =>
	h.parts
		.map((p) => (p.current ? `*${p.label}*` : p.label) + (p.ellipsis ? ', ...' : ''))
		.join(', ');

describe('buildArgHint', () => {
	it('marks the current argument', () => {
		expect(text(buildArgHint(SUM, 0))).toBe('*number1*, [number2], ...');
		expect(text(buildArgHint(SUM, 1))).toBe('number1, *[number2]*, ...');
	});

	it('extends a repeating argument past the listed ones', () => {
		expect(text(buildArgHint(SUM, 3))).toBe('number1, [number2], [number3], *[number4]*, ...');
	});

	it('extends repeating pairs as whole groups', () => {
		expect(text(buildArgHint(SUMIFS, 5))).toBe(
			'sum_range, criteria_range1, criteria1, [criteria_range2], [criteria2], *[criteria_range3]*, [criteria3], ...',
		);
	});

	it('keeps arguments that follow a repeating group last', () => {
		expect(text(buildArgHint(LAMBDA, 0))).toBe('*[parameter1]*, ..., calculation');
		expect(text(buildArgHint(LAMBDA, 2))).toBe(
			'[parameter1], [parameter2], *[parameter3]*, ..., calculation',
		);
	});

	it('has no current argument past the end of a fixed signature', () => {
		const hint = buildArgHint(ABS, 2);
		expect(hint.parts.some((p) => p.current)).toBe(false);
		expect(hint.parts.some((p) => p.ellipsis)).toBe(false);
	});
});
