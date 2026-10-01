import { parseCellAddress } from '../../address/cell-address';
import { parseRange, parseSqref } from '../../address/range';
import { createWorkbook } from '../../sheet/workbook';
import { putCell } from '../../sheet/worksheet';
import type { RangeAddress } from '../../types/address';
import type { Cell, CellValue } from '../../types/cell';
import type { CfRule } from '../../types/conditional-format';
import type { Worksheet } from '../../types/sheet';
import type { Workbook } from '../../types/workbook';
import { ConditionalFormatEngine } from '../engine';
import type { ConditionalFormatEngineOptions } from '../types';
import { createFakeEvaluator, type FakeEvaluator } from './fake-evaluator';

/** Range from A1 notation (throws on invalid input). */
export function rng(ref: string): RangeAddress {
	const r = parseRange(ref);
	if (!r) {
		throw new Error(`bad range ${ref}`);
	}
	return r;
}

/** Set cell values by A1 address. */
export function setCells(sheet: Worksheet, cells: Record<string, CellValue | Cell>): void {
	for (const [addr, value] of Object.entries(cells)) {
		const a = parseCellAddress(addr);
		if (!a) {
			throw new Error(`bad address ${addr}`);
		}
		const cell: Cell =
			value !== null && typeof value === 'object' && 'value' in value
				? value
				: { value: value as CellValue };
		putCell(sheet, a.row, a.col, cell);
	}
}

/** Fill consecutive cells downwards starting at `start`. */
export function setColumn(sheet: Worksheet, start: string, values: CellValue[]): void {
	const a = parseCellAddress(start)!;
	values.forEach((value, i) => putCell(sheet, a.row + i, a.col, { value }));
}

let autoPriority = 0;

/** Append a conditional format block with one rule (priority defaults to an increasing counter). */
export function addRule(
	sheet: Worksheet,
	sqref: string,
	rule: Omit<CfRule, 'priority'> & { priority?: number },
): CfRule {
	const full = { ...rule, priority: rule.priority ?? ++autoPriority } as CfRule;
	sheet.conditionalFormats.push({ ranges: parseSqref(sqref), rules: [full] });
	return full;
}

export interface CfFixture {
	workbook: Workbook;
	sheet: Worksheet;
	evaluator: FakeEvaluator;
	engine: ConditionalFormatEngine;
	/** Result at an A1 address of the first sheet. */
	at: (addr: string) => ReturnType<ConditionalFormatEngine['getCellResult']>;
}

/** A one-sheet workbook (plus optional extra sheets) with a fake evaluator and an engine. */
export function cfFixture(
	cells: Record<string, CellValue | Cell> = {},
	options: ConditionalFormatEngineOptions & { sheets?: string[] } = {},
): CfFixture {
	const workbook = createWorkbook({ sheetNames: options.sheets ?? ['Sheet1'] });
	const sheet = workbook.sheets[0];
	setCells(sheet, cells);
	const evaluator = createFakeEvaluator(workbook);
	const engine = new ConditionalFormatEngine(workbook, evaluator, options);
	const at = (addr: string) => {
		const a = parseCellAddress(addr)!;
		return engine.getCellResult(0, a.row, a.col);
	};
	return { workbook, sheet, evaluator, engine, at };
}
