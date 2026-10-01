import { describe, expect, it } from 'vitest';
import { createCellIsRule, createDataBarRule, createWorkbook, parseRange } from 'xlsx-viewer-core';

import { selectRange } from '../components/dialogs/__tests__/render-dialog';
import { ViewerStore } from '../state/viewer-store';
import { addConditionalRule, clearConditionalRules } from './cf-commands';

function store() {
	return new ViewerStore(createWorkbook());
}

describe('addConditionalRule', () => {
	it('adds the rule to the selection, interning its format, as one undo step', () => {
		const s = store();
		selectRange(s, 'B2:B9');
		const dxfs = s.workbook.styles.dxfs.length;
		expect(addConditionalRule(s, createCellIsRule('greaterThan', ['5'], 'redText'))).toBe(true);
		const formats = s.activeSheet.conditionalFormats;
		expect(formats).toHaveLength(1);
		expect(formats[0].ranges).toEqual([parseRange('B2:B9')]);
		expect(formats[0].rules[0]).toMatchObject({ type: 'cellIs', priority: 1, dxfId: dxfs });
		expect(s.workbook.styles.dxfs).toHaveLength(dxfs + 1);
		expect(s.getSnapshot().history.undo).toBe('Conditional Formatting');
		s.undo();
		expect(s.activeSheet.conditionalFormats).toHaveLength(0);
		s.redo();
		expect(s.activeSheet.conditionalFormats).toHaveLength(1);
	});

	it('gives the newest rule the top priority', () => {
		const s = store();
		selectRange(s, 'A1:A5');
		addConditionalRule(s, createDataBarRule('FF638EC6'));
		addConditionalRule(s, createCellIsRule('lessThan', ['0'], 'redText'));
		const rules = s.activeSheet.conditionalFormats.flatMap((f) => f.rules);
		expect(rules.find((r) => r.type === 'cellIs')?.priority).toBe(1);
		expect(rules.find((r) => r.type === 'dataBar')?.priority).toBe(2);
	});

	it('is refused when read-only', () => {
		const s = new ViewerStore(createWorkbook(), { readOnly: true });
		expect(addConditionalRule(s, createDataBarRule('FF638EC6'))).toBe(false);
		expect(s.activeSheet.conditionalFormats).toHaveLength(0);
	});
});

describe('clearConditionalRules', () => {
	it('clears rules of the selection or of the whole sheet', () => {
		const s = store();
		selectRange(s, 'A1:A5');
		addConditionalRule(s, createDataBarRule('FF638EC6'));
		selectRange(s, 'C1:C5');
		addConditionalRule(s, createDataBarRule('FF63C384'));
		selectRange(s, 'A3');
		expect(clearConditionalRules(s, 'selection')).toBe(true);
		const ranges = s.activeSheet.conditionalFormats.map((f) => f.ranges);
		expect(ranges).toContainEqual([parseRange('C1:C5')]);
		expect(ranges).toContainEqual([parseRange('A1:A2'), parseRange('A4:A5')]);
		expect(clearConditionalRules(s, 'sheet')).toBe(true);
		expect(s.activeSheet.conditionalFormats).toHaveLength(0);
	});

	it('adds no undo step when there is nothing to clear', () => {
		const s = store();
		addConditionalRule(s, createDataBarRule('FF638EC6'));
		selectRange(s, 'Z99');
		expect(clearConditionalRules(s, 'selection')).toBe(false);
		expect(s.getSnapshot().history.undo).toBe('Conditional Formatting');
		const empty = store();
		expect(clearConditionalRules(empty, 'sheet')).toBe(false);
		expect(empty.getSnapshot().history.undo).toBeNull();
	});
});
