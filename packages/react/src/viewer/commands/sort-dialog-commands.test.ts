import { describe, expect, it } from 'vitest';
import {
	cellKey,
	createWorkbook,
	EditTransaction,
	applyStylePatch,
	parseRange,
} from 'xlsx-viewer-core';

import { selectRange, typeCells } from '../components/dialogs/__tests__/render-dialog';
import { currentRange } from '../../engine/selection/selection';
import { createTranslator } from '../../i18n';
import { ViewerStore } from '../state/viewer-store';
import {
	applySortLevels,
	checkSortLevels,
	selectSortBlock,
	sortDataRange,
	sortDialogSetup,
	sortKeyOptions,
	sortOrderLabels,
} from './sort-dialog-commands';

const r = (a1: string) => parseRange(a1)!;
const t = createTranslator();

const SORT_MULTIPLE_SELECTION_MESSAGE = "This action won't work on multiple selections.";
const SORT_EMPTY_MESSAGE =
	"This can't be applied to the selected range. Select a single cell in a range and try again.";
const SORT_COLUMN_REQUIRED_MESSAGE =
	'All sort criteria must have a column specified. Check the selected sort criteria and try again.';

function peopleStore() {
	const wb = createWorkbook();
	typeCells(wb, {
		A1: 'Name',
		B1: 'Age',
		C1: 'Joined',
		A2: 'carol',
		B2: '41',
		C2: '3/1/2021',
		A3: 'Alice',
		B3: '29',
		C3: '1/15/2020',
		A4: 'bob',
		B4: '35',
		C4: '7/4/2019',
		A5: 'Alice',
		B5: '52',
		C5: '2/2/2022',
	});
	return new ViewerStore(wb);
}

const column = (store: ViewerStore, col: number, from = 1, to = 4) => {
	const out: unknown[] = [];
	for (let row = from; row <= to; row++) {
		out.push(store.activeSheet.cells.get(cellKey(row, col))?.value ?? null);
	}
	return out;
};

describe('sortDialogSetup', () => {
	it('expands a single cell to its current region and guesses the header', () => {
		const store = peopleStore();
		selectRange(store, 'B3');
		const setup = sortDialogSetup(store);
		expect(setup.block).toEqual(r('A1:C5'));
		expect(setup.hasHeader).toBe(true);
		expect(setup.active).toEqual({ row: 2, col: 1 });
		expect(setup.error).toBeUndefined();
	});

	it('clamps whole columns to the used range', () => {
		const store = peopleStore();
		selectRange(store, 'A:B');
		expect(sortDialogSetup(store).block).toEqual(r('A1:B5'));
	});

	it('refuses multiple selections and empty cells', () => {
		const store = peopleStore();
		store.setSelection({
			active: { row: 0, col: 0 },
			anchor: { row: 0, col: 0 },
			ranges: [r('A1:A2'), r('C1:C2')],
		});
		expect(sortDialogSetup(store)).toMatchObject({
			block: null,
			error: SORT_MULTIPLE_SELECTION_MESSAGE,
		});
		selectRange(store, 'H20');
		expect(sortDialogSetup(store)).toMatchObject({ block: null, error: SORT_EMPTY_MESSAGE });
		selectRange(store, 'H20:J30');
		expect(sortDialogSetup(store).error).toBe(SORT_EMPTY_MESSAGE);
		expect(sortDialogSetup(new ViewerStore(createWorkbook())).error).toBe(SORT_EMPTY_MESSAGE);
	});
});

describe('sortKeyOptions', () => {
	it('labels columns by header text and detects text, numbers and dates', () => {
		const store = peopleStore();
		expect(sortKeyOptions(store, r('A1:C5'), { hasHeader: true, byColumns: false })).toEqual([
			{ index: 0, label: 'Name', kind: 'text' },
			{ index: 1, label: 'Age', kind: 'number' },
			{ index: 2, label: 'Joined', kind: 'date' },
		]);
	});

	it('uses column letters without headers and row numbers left to right', () => {
		const store = peopleStore();
		const plain = sortKeyOptions(store, r('A1:C5'), { hasHeader: false, byColumns: false });
		expect(plain.map((o) => o.label)).toEqual(['Column A', 'Column B', 'Column C']);
		// The header text is data now, so the first value decides the kind.
		expect(plain[1].kind).toBe('text');
		const rows = sortKeyOptions(store, r('A2:C3'), { hasHeader: false, byColumns: true });
		expect(rows).toEqual([
			{ index: 1, label: 'Row 2', kind: 'text' },
			{ index: 2, label: 'Row 3', kind: 'text' },
		]);
	});

	it('falls back to the column letter for blank headers', () => {
		const store = peopleStore();
		typeCells(store.workbook, { B1: '' });
		const labels = sortKeyOptions(store, r('A1:C5'), { hasHeader: true, byColumns: false }).map(
			(o) => o.label,
		);
		expect(labels).toEqual(['Name', 'Column B', 'Joined']);
	});
});

describe('sort dialog helpers', () => {
	it('words the order by value kind', () => {
		expect(sortOrderLabels('text', t)).toEqual(['A to Z', 'Z to A']);
		expect(sortOrderLabels('number', t)).toEqual(['Smallest to Largest', 'Largest to Smallest']);
		expect(sortOrderLabels('date', t)).toEqual(['Oldest to Newest', 'Newest to Oldest']);
	});

	it('finds the data part of the block', () => {
		expect(sortDataRange(r('A1:C5'), { hasHeader: true, byColumns: false })).toEqual(r('A2:C5'));
		expect(sortDataRange(r('A1:C5'), { hasHeader: true, byColumns: true })).toEqual(r('B1:C5'));
		expect(sortDataRange(r('A1:C1'), { hasHeader: true, byColumns: false })).toEqual(r('A1:C1'));
		expect(sortDataRange(r('A1:C5'), { hasHeader: false, byColumns: false })).toEqual(r('A1:C5'));
	});

	it('rejects levels without a column and duplicate columns', () => {
		const options = sortKeyOptions(peopleStore(), r('A1:C5'), {
			hasHeader: true,
			byColumns: false,
		});
		expect(
			checkSortLevels(
				[
					{ key: 0, descending: false },
					{ key: null, descending: false },
				],
				options,
				t,
			),
		).toBe(SORT_COLUMN_REQUIRED_MESSAGE);
		expect(
			checkSortLevels(
				[
					{ key: 1, descending: false },
					{ key: 1, descending: true },
				],
				options,
				t,
			),
		).toBe(
			'Age is being sorted by values more than once. Delete the duplicate sort condition and try again.',
		);
		expect(
			checkSortLevels(
				[
					{ key: 1, descending: false },
					{ key: 0, descending: true },
				],
				options,
				t,
			),
		).toBeNull();
		expect(checkSortLevels([], options, t)).toBeNull();
	});

	it('selects the data rows of the block', () => {
		const store = peopleStore();
		selectRange(store, 'B1');
		selectSortBlock(store, r('A1:C5'), { hasHeader: true, byColumns: false });
		const sel = store.selectionOf();
		expect(currentRange(sel)).toEqual(r('A2:C5'));
		expect(sel.active).toEqual({ row: 1, col: 1 });
		selectSortBlock(store, r('A1:C5'), { hasHeader: false, byColumns: false });
		expect(currentRange(store.selectionOf())).toEqual(r('A1:C5'));
	});
});

describe('applySortLevels', () => {
	it('sorts by several levels in one undoable step', () => {
		const store = peopleStore();
		const options = sortKeyOptions(store, r('A1:C5'), { hasHeader: true, byColumns: false });
		const levels = [
			{ key: 0, descending: false },
			{ key: 1, descending: true },
		];
		expect(
			applySortLevels(
				store,
				r('A1:C5'),
				levels,
				{ hasHeader: true, byColumns: false, caseSensitive: false },
				options,
			),
		).toBeNull();
		expect(column(store, 0)).toEqual(['Alice', 'Alice', 'bob', 'carol']);
		expect(column(store, 1)).toEqual([52, 29, 35, 41]);
		expect(store.activeSheet.cells.get(cellKey(0, 0))?.value).toBe('Name');
		expect(store.getSnapshot().history.undo).toBe('Sort');
		store.undo();
		expect(column(store, 0)).toEqual(['carol', 'Alice', 'bob', 'Alice']);
	});

	it('sorts left to right', () => {
		const wb = createWorkbook();
		typeCells(wb, { A1: 'c', B1: 'a', C1: 'b', A2: 1, B2: 2, C2: 3 });
		const store = new ViewerStore(wb);
		const opts = { hasHeader: false, byColumns: true, caseSensitive: false };
		expect(
			applySortLevels(store, r('A1:C2'), [{ key: 0, descending: false }], opts, []),
		).toBeNull();
		expect([0, 1, 2].map((c) => column(store, c, 0, 1))).toEqual([
			['a', 2],
			['b', 3],
			['c', 1],
		]);
	});

	it('passes the case sensitivity on', () => {
		const sortRow = (a: string, b: string, caseSensitive: boolean) => {
			const wb = createWorkbook();
			typeCells(wb, { A1: a, B1: b });
			const store = new ViewerStore(wb);
			applySortLevels(
				store,
				r('A1:B1'),
				[{ key: 0, descending: false }],
				{ hasHeader: false, byColumns: true, caseSensitive },
				[],
			);
			return column(store, 0, 0, 0).concat(column(store, 1, 0, 0));
		};
		// Ignoring case, "b" and "B" tie and keep their places.
		expect(sortRow('b', 'B', false)).toEqual(['b', 'B']);
		expect(sortRow('B', 'b', false)).toEqual(['B', 'b']);
		// Matching case, they have a fixed order whatever the input order.
		expect(sortRow('b', 'B', true)).toEqual(sortRow('B', 'b', true));
	});

	it('reports merged cells and bad levels without sorting', () => {
		const store = peopleStore();
		store.workbook.sheets[0].merges.push(r('A2:A3'));
		const opts = { hasHeader: true, byColumns: false, caseSensitive: false };
		expect(applySortLevels(store, r('A1:C5'), [{ key: 0, descending: false }], opts, [])).toBe(
			'To do this, all the merged cells need to be the same size.',
		);
		expect(applySortLevels(store, r('A1:C5'), [{ key: null, descending: false }], opts, [])).toBe(
			SORT_COLUMN_REQUIRED_MESSAGE,
		);
		expect(store.isDirty).toBe(false);
	});

	it('does nothing without levels', () => {
		const store = peopleStore();
		expect(
			applySortLevels(
				store,
				r('A1:C5'),
				[],
				{ hasHeader: true, byColumns: false, caseSensitive: false },
				[],
			),
		).toBeNull();
		expect(store.isDirty).toBe(false);
	});

	it('keeps date detection on formatted numbers', () => {
		const wb = createWorkbook();
		typeCells(wb, { A1: 'When', A2: 45000, A3: 44000 });
		const tx = new EditTransaction(wb, 'fmt');
		applyStylePatch(tx, wb, 0, [r('A2:A3')], { numFmt: 'yyyy-mm-dd' });
		tx.commit();
		const store = new ViewerStore(wb);
		expect(sortKeyOptions(store, r('A1:A3'), { hasHeader: true, byColumns: false })[0].kind).toBe(
			'date',
		);
	});
});
