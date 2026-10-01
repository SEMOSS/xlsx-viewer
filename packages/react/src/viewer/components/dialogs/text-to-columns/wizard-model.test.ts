import { describe, expect, it } from 'vitest';
import {
	createWorkbook,
	EditTransaction,
	MAX_ROW_INDEX,
	parseRange,
	setCellInput,
} from 'xlsx-viewer-core';

import { createTranslator } from '../../../../i18n';
import { ViewerStore } from '../../../state/viewer-store';
import {
	addBreak,
	columnCountOf,
	DEFAULT_DELIMITERS,
	formatBreaks,
	formatDestination,
	formatLabel,
	guessSplitMode,
	parseBreaksText,
	parseDestination,
	splitOptionsOf,
	textToColumnsSetup,
} from './wizard-model';

function storeWith(entries: Array<[number, string]>) {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	for (const [row, text] of entries) {
		setCellInput(tx, wb, 0, row, 0, text);
	}
	tx.commit();
	return new ViewerStore(wb);
}

const select = (store: ViewerStore, ...refs: string[]) => {
	const ranges = refs.map((r) => parseRange(r)!);
	const active = { row: ranges[0].top, col: ranges[0].left };
	store.setSelection({ active, anchor: active, ranges });
};

describe('textToColumnsSetup', () => {
	it('uses the selected column, limited to the used rows', () => {
		const store = storeWith([
			[0, 'a,b'],
			[2, 'c'],
		]);
		select(store, 'A1:A5');
		expect(textToColumnsSetup(store)).toEqual({
			source: parseRange('A1:A5'),
			lines: ['a,b', '', 'c', '', ''],
			error: null,
		});
		store.setSelection({
			active: { row: 0, col: 0 },
			anchor: { row: 0, col: 0 },
			ranges: [{ top: 0, left: 0, bottom: MAX_ROW_INDEX, right: 0 }],
		});
		expect(textToColumnsSetup(store).source).toEqual(parseRange('A1:A3'));
	});

	it('refuses several columns, several areas and empty selections', () => {
		const store = storeWith([[0, 'x']]);
		select(store, 'A1:B2');
		expect(textToColumnsSetup(store).error).toContain('can convert only one column at a time');
		select(store, 'A1', 'A3');
		expect(textToColumnsSetup(store).error).toBe(
			'This command cannot be used on multiple selections.',
		);
		select(store, 'C1:C4');
		expect(textToColumnsSetup(store).error).toBe('No data was selected to parse.');
	});
});

describe('wizard helpers', () => {
	it('guesses the data type', () => {
		expect(guessSplitMode(['a,b', 'c d'])).toBe('delimited');
		expect(guessSplitMode(['Ann   42', 'Bob   7', ''])).toBe('fixed');
		expect(guessSplitMode(['one line'])).toBe('delimited');
	});

	it('builds split options', () => {
		expect(splitOptionsOf('fixed', DEFAULT_DELIMITERS, [4])).toEqual({
			mode: 'fixed',
			breaks: [4],
		});
		expect(
			splitOptionsOf(
				'delimited',
				{ ...DEFAULT_DELIMITERS, other: true, otherChar: '|', qualifier: 'none' },
				[],
			),
		).toEqual({
			mode: 'delimited',
			tab: true,
			semicolon: false,
			comma: false,
			space: false,
			other: '|',
			treatConsecutiveAsOne: false,
			textQualifier: null,
		});
		expect(
			splitOptionsOf('delimited', { ...DEFAULT_DELIMITERS, other: true }, []).other,
		).toBeUndefined();
	});

	it('reads and writes break lists', () => {
		expect(parseBreaksText('12, 5 x 5;0 -3 7')).toEqual([5, 7, 12]);
		expect(formatBreaks([5, 12])).toBe('5, 12');
		expect(addBreak([5], 3, 10)).toEqual([3, 5]);
		expect(addBreak([5], 5, 10)).toEqual([5]);
		expect(addBreak([5], 0, 10)).toEqual([5]);
		expect(addBreak([5], 10, 10)).toEqual([5]);
	});

	it('parses the destination on the active sheet', () => {
		expect(formatDestination(0, 0)).toBe('$A$1');
		expect(parseDestination('$C$4', 'Sheet1')).toEqual({ row: 3, col: 2 });
		expect(parseDestination('=B2:D9', 'Sheet1')).toEqual({ row: 1, col: 1 });
		expect(parseDestination("'sheet1'!B2", 'Sheet1')).toEqual({ row: 1, col: 1 });
		expect(parseDestination('Other!B2', 'Sheet1')).toBeNull();
		expect(parseDestination('nope!', 'Sheet1')).toBeNull();
		expect(parseDestination('', 'Sheet1')).toBeNull();
	});

	it('counts columns and labels formats', () => {
		expect(columnCountOf([])).toBe(1);
		expect(columnCountOf([['a'], ['b', 'c', 'd'], []])).toBe(3);
		const t = createTranslator();
		expect(['general', 'text', 'skip'].map((f) => formatLabel(f as 'general', t))).toEqual([
			'General',
			'Text',
			'Skip column',
		]);
	});
});
