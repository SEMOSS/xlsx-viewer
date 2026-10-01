import { isValidSheetName, quoteSheetName } from '../address/sheet-name';
import { createWorksheet, generateId } from '../sheet/worksheet';
import { findSheet, nextSheetId, nextSheetName } from '../sheet/workbook';
import type { Worksheet } from '../types/sheet';
import type { ColorSpec } from '../types/style';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

/** Rewrite sheet-qualified references from `oldName` to `newName` (outside string literals). */
export function renameSheetReferences(formula: string, oldName: string, newName: string): string {
	const lowerOld = oldName.toLowerCase();
	let out = '';
	let i = 0;
	while (i < formula.length) {
		const ch = formula[i];
		if (ch === '"') {
			let j = i + 1;
			while (j < formula.length) {
				if (formula[j] === '"') {
					if (formula[j + 1] === '"') {
						j += 2;
						continue;
					}
					break;
				}
				j++;
			}
			out += formula.slice(i, j + 1);
			i = j + 1;
			continue;
		}
		if (ch === "'") {
			let j = i + 1;
			while (j < formula.length) {
				if (formula[j] === "'") {
					if (formula[j + 1] === "'") {
						j += 2;
						continue;
					}
					break;
				}
				j++;
			}
			const name = formula.slice(i + 1, j).replace(/''/g, "'");
			if (formula[j + 1] === '!' && name.toLowerCase() === lowerOld) {
				out += quoteSheetName(newName);
			} else {
				out += formula.slice(i, j + 1);
			}
			i = j + 1;
			continue;
		}
		const prev = i > 0 ? formula[i - 1] : '';
		if (!/[A-Za-z0-9_.\u00C0-\uFFFF]/.test(prev)) {
			const candidate = formula.slice(i, i + oldName.length);
			if (candidate.toLowerCase() === lowerOld && formula[i + oldName.length] === '!') {
				out += quoteSheetName(newName);
				i += oldName.length;
				continue;
			}
		}
		out += ch;
		i++;
	}
	return out;
}

function rewriteAllFormulas(workbook: Workbook, rewrite: (f: string) => string): void {
	for (const sheet of workbook.sheets) {
		let touched = false;
		for (const [key, cell] of sheet.cells) {
			if (cell.formula) {
				const next = rewrite(cell.formula);
				if (next !== cell.formula) {
					sheet.cells.set(key, { ...cell, formula: next });
					touched = true;
				}
			}
		}
		if (touched) {
			sheet.dirty = true;
		}
	}
	for (const name of workbook.definedNames) {
		name.formula = rewrite(name.formula);
	}
}

export class SheetNameError extends Error {
	constructor(
		message: string,
		readonly reason: 'invalid' | 'duplicate',
	) {
		super(message);
		this.name = 'SheetNameError';
	}
}

function assertName(workbook: Workbook, name: string, except?: Worksheet): void {
	if (!isValidSheetName(name)) {
		throw new SheetNameError(`Invalid sheet name: ${name}`, 'invalid');
	}
	const existing = findSheet(workbook, name);
	if (existing && existing !== except) {
		throw new SheetNameError(`A sheet named ${name} already exists`, 'duplicate');
	}
}

/** Insert a new empty sheet at `index` (default: after the active sheet). Returns its index. */
export function addSheet(
	tx: EditTransaction,
	workbook: Workbook,
	index?: number,
	name?: string,
): number {
	tx.touchWorkbook('sheets');
	const sheetName = name ?? nextSheetName(workbook);
	assertName(workbook, sheetName);
	const sheet = createWorksheet(sheetName, nextSheetId(workbook));
	const at = Math.max(0, Math.min(workbook.sheets.length, index ?? workbook.view.activeTab + 1));
	workbook.sheets.splice(at, 0, sheet);
	shiftLocalNames(workbook, at, 1);
	workbook.dirty = true;
	return at;
}

export function renameSheet(
	tx: EditTransaction,
	workbook: Workbook,
	index: number,
	newName: string,
): void {
	const sheet = workbook.sheets[index];
	if (sheet.name === newName) {
		return;
	}
	assertName(workbook, newName, sheet);
	tx.touchWorkbook('sheets');
	for (let i = 0; i < workbook.sheets.length; i++) {
		tx.touchSheet(i, false, 'references');
	}
	const old = sheet.name;
	rewriteAllFormulas(workbook, (f) => renameSheetReferences(f, old, newName));
	sheet.name = newName;
	sheet.dirty = true;
	workbook.dirty = true;
}

/** Delete a sheet. References to it become `#REF!` (simplified: formulas are left to the calc engine). */
export function deleteSheet(tx: EditTransaction, workbook: Workbook, index: number): void {
	if (
		workbook.sheets.filter((s) => s.state === 'visible').length <= 1 &&
		workbook.sheets[index].state === 'visible'
	) {
		throw new Error('A workbook must contain at least one visible worksheet');
	}
	tx.touchWorkbook('sheets');
	workbook.sheets.splice(index, 1);
	workbook.definedNames = workbook.definedNames.filter((n) => n.localSheet !== index);
	shiftLocalNames(workbook, index + 1, -1);
	if (workbook.view.activeTab >= workbook.sheets.length) {
		workbook.view.activeTab = workbook.sheets.length - 1;
	}
	workbook.dirty = true;
}

/** Move a sheet to a new position. */
export function moveSheet(tx: EditTransaction, workbook: Workbook, from: number, to: number): void {
	if (from === to) {
		return;
	}
	tx.touchWorkbook('sheets');
	const [sheet] = workbook.sheets.splice(from, 1);
	workbook.sheets.splice(to, 0, sheet);
	for (const n of workbook.definedNames) {
		if (n.localSheet === undefined) {
			continue;
		}
		if (n.localSheet === from) {
			n.localSheet = to;
		} else if (from < to && n.localSheet > from && n.localSheet <= to) {
			n.localSheet -= 1;
		} else if (from > to && n.localSheet >= to && n.localSheet < from) {
			n.localSheet += 1;
		}
	}
	workbook.dirty = true;
}

/** Copy a sheet (Move or Copy > Create a copy). Returns the new index. */
export function duplicateSheet(tx: EditTransaction, workbook: Workbook, index: number): number {
	tx.touchWorkbook('sheets');
	const src = workbook.sheets[index];
	let n = 2;
	let name = `${src.name} (${n})`;
	while (findSheet(workbook, name)) {
		n++;
		name = `${src.name.slice(0, 25)} (${n})`;
	}
	const copy: Worksheet = {
		...structuredClone({
			...src,
			cells: undefined,
			rows: undefined,
			source: undefined,
			drawings: [],
		}),
		id: generateId('sheet'),
		name,
		sheetId: nextSheetId(workbook),
		cells: new Map(src.cells),
		rows: new Map(src.rows),
		drawings: src.drawings.map((d) => ({
			...d,
			id: generateId('obj'),
			raw: undefined,
			dirty: true,
		})),
		source: undefined,
		dirty: true,
		drawingsDirty: true,
		commentsDirty: true,
		tables: [],
	} as Worksheet;
	copy.view = { ...src.view, tabSelected: false };
	workbook.sheets.splice(index + 1, 0, copy);
	shiftLocalNames(workbook, index + 1, 1);
	workbook.dirty = true;
	return index + 1;
}

export function setSheetTabColor(
	tx: EditTransaction,
	workbook: Workbook,
	index: number,
	color: ColorSpec | undefined,
): void {
	tx.checkWorkbook('sheets');
	tx.touchSheet(index, false, 'sheets');
	workbook.sheets[index].tabColor = color;
	workbook.sheets[index].dirty = true;
}

export function setSheetVisibility(
	tx: EditTransaction,
	workbook: Workbook,
	index: number,
	state: Worksheet['state'],
): void {
	tx.touchWorkbook('sheets');
	tx.touchSheet(index, false, 'sheets');
	if (state !== 'visible') {
		const visible = workbook.sheets.filter((s, i) => i !== index && s.state === 'visible');
		if (visible.length === 0) {
			throw new Error('A workbook must contain at least one visible worksheet');
		}
	}
	workbook.sheets[index].state = state;
	workbook.sheets[index].dirty = true;
	workbook.dirty = true;
}

function shiftLocalNames(workbook: Workbook, from: number, delta: number): void {
	for (const n of workbook.definedNames) {
		if (n.localSheet !== undefined && n.localSheet >= from) {
			n.localSheet += delta;
		}
	}
}
