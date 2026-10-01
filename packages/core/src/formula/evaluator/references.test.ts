import { describe, expect, it } from 'vitest';

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import { ERRORS } from '../../values/errors';
import type { CellRefNode, RangeRefNode } from '../ast';
import { parseFormula } from '../language/parser';
import type { RefValue } from '../types';
import type { EvaluationHost } from './host';
import {
	cellNodeRange,
	intersectionOperator,
	intersectRef,
	rangeNodeRange,
	rangeOperator,
	resolveSheetSpan,
	unionOperator,
} from './references';

const ref = (top: number, left: number, bottom = top, right = left, sheet = 0): RefValue => ({
	kind: 'ref',
	sheet,
	range: { top, left, bottom, right },
});

const host = {
	sheetIndex: (name: string) => ['sheet1', 'sheet2', 'sheet3'].indexOf(name.toLowerCase()),
} as unknown as EvaluationHost;

describe('reference nodes', () => {
	it('normalises ranges and spans whole rows / columns', () => {
		const r = parseFormula('C5:A1') as RangeRefNode;
		expect(rangeNodeRange(r, false, 0, 0)).toEqual({ top: 0, left: 0, bottom: 4, right: 2 });
		const cols = parseFormula('B:D') as RangeRefNode;
		expect(rangeNodeRange(cols, false, 0, 0)).toEqual({
			top: 0,
			left: 1,
			bottom: MAX_ROW_INDEX,
			right: 3,
		});
		const rows = parseFormula('3:2') as RangeRefNode;
		expect(rangeNodeRange(rows, false, 0, 0)).toEqual({
			top: 1,
			left: 0,
			bottom: 2,
			right: MAX_COL_INDEX,
		});
	});

	it('shifts relative references of defined names by the formula cell, wrapping the grid', () => {
		const rel = parseFormula('B2') as CellRefNode;
		expect(cellNodeRange(rel, true, 3, 4)).toEqual({ top: 4, left: 5, bottom: 4, right: 5 });
		const abs = parseFormula('$B$2') as CellRefNode;
		expect(cellNodeRange(abs, true, 3, 4)).toEqual({ top: 1, left: 1, bottom: 1, right: 1 });
		const up = parseFormula('A1048576') as CellRefNode;
		expect(cellNodeRange(up, true, 2, 0)).toEqual({ top: 1, left: 0, bottom: 1, right: 0 });
		const mixed = parseFormula('$A1:A$1') as RangeRefNode;
		expect(rangeNodeRange(mixed, true, 5, 2)).toEqual({ top: 0, left: 0, bottom: 5, right: 2 });
	});

	it('resolves sheet qualifiers, 3-D spans and unknown sheets', () => {
		expect(resolveSheetSpan(undefined, 1, host)).toEqual({ start: 1, end: 1 });
		expect(resolveSheetSpan({ sheet: 'SHEET2' }, 0, host)).toEqual({ start: 1, end: 1 });
		expect(resolveSheetSpan({ sheet: 'Sheet3', sheetEnd: 'Sheet1' }, 0, host)).toEqual({
			start: 0,
			end: 2,
		});
		expect(resolveSheetSpan({ sheet: 'Nope' }, 0, host)).toBe(ERRORS.REF);
		expect(resolveSheetSpan({ sheet: 'Sheet1', workbook: '1' }, 0, host)).toBe(ERRORS.REF);
	});
});

describe('reference operators', () => {
	it('builds the bounding range with `:`', () => {
		expect(rangeOperator(ref(0, 0), ref(4, 2))).toEqual(ref(0, 0, 4, 2));
		expect(rangeOperator(ref(4, 2), ref(0, 0, 1, 1))).toEqual(ref(0, 0, 4, 2));
		expect(rangeOperator(ref(0, 0), ref(0, 0, 0, 0, 1))).toBe(ERRORS.VALUE);
		expect(rangeOperator(ref(0, 0), 5)).toBe(ERRORS.VALUE);
		expect(rangeOperator(ERRORS.NA, ref(0, 0))).toBe(ERRORS.NA);
	});

	it('unions areas on one sheet', () => {
		const u = unionOperator(ref(0, 0), ref(2, 2, 3, 3));
		expect(u).toEqual({ kind: 'refs', refs: [ref(0, 0), ref(2, 2, 3, 3)] });
		const u3 = unionOperator(u, ref(9, 9));
		expect(u3).toMatchObject({ kind: 'refs', refs: { length: 3 } });
		expect(unionOperator(ref(0, 0), ref(0, 0, 0, 0, 1))).toBe(ERRORS.VALUE);
		expect(unionOperator(ref(0, 0), 'x')).toBe(ERRORS.VALUE);
	});

	it('intersects references, #NULL! when disjoint', () => {
		expect(intersectionOperator(ref(0, 0, 4, 4), ref(2, 2, 9, 9))).toEqual(ref(2, 2, 4, 4));
		expect(intersectionOperator(ref(0, 0), ref(5, 5))).toBe(ERRORS.NULL);
		const multi = unionOperator(ref(0, 0, 0, 9), ref(5, 0, 5, 9));
		expect(intersectionOperator(multi, ref(0, 3, 9, 3))).toEqual({
			kind: 'refs',
			refs: [ref(0, 3), ref(5, 3)],
		});
		expect(intersectionOperator(ref(0, 0), ref(0, 0, 0, 0, 1))).toBe(ERRORS.NULL);
	});

	it('implicitly intersects keeping a reference', () => {
		expect(intersectRef(ref(0, 0, 9, 0), 4, 7)).toEqual(ref(4, 0));
		expect(intersectRef(ref(2, 0, 2, 9), 0, 3)).toEqual(ref(2, 3));
		expect(intersectRef(ref(0, 0, 3, 3), 1, 1)).toBe(ERRORS.VALUE);
		expect(intersectRef(ref(0, 0, 9, 0), 20, 0)).toBe(ERRORS.VALUE);
		expect(intersectRef(ref(5, 5), 0, 0)).toEqual(ref(5, 5));
	});
});
