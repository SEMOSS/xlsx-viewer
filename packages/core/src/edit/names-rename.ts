import { getFunctionInfo } from '../formula/catalog';
import {
	analyzeCalls,
	applyEdits,
	soleToken,
	type TextEdit,
} from '../formula/language/token-utils';
import { tokenizeFormula } from '../formula/language/tokenizer';
import type { FormulaToken } from '../formula/language/tokens';
import { findSheetIndex } from '../sheet/workbook';
import type { Cell } from '../types/cell';
import type { CfRule, Cfvo, ConditionalFormat } from '../types/conditional-format';
import type { DataValidation } from '../types/data-validation';
import type { Table, TableColumn } from '../types/table';
import type { DefinedName, Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

/**
 * Renaming a defined name: every formula that uses it (cells, other names,
 * conditional formatting, data validation and table columns) follows the
 * new name, the way Excel's Name Manager does. Scoping is respected: a
 * sheet-level name shadows a workbook-level name of the same spelling on
 * its own sheet.
 */

/**
 * Decides whether a use of the name points at the renamed definition.
 * `qualifier` is the unquoted sheet of a `Sheet1!Rate` style use, or
 * `undefined` for a plain `Rate`.
 */
export type NameUseMatcher = (qualifier: string | undefined) => boolean;

/** Token ranges `[start, end)` of LET / LAMBDA calls that declare a parameter spelled `lower`. */
function shadowingCalls(tokens: readonly FormulaToken[], lower: string): Array<[number, number]> {
	const out: Array<[number, number]> = [];
	for (const call of analyzeCalls(tokens)) {
		if (call.name !== 'LET' && call.name !== 'LAMBDA') {
			continue;
		}
		const params = call.args.slice(0, -1).filter((_, k) => call.name === 'LAMBDA' || k % 2 === 0);
		const declares = params.some(([s, e]) => {
			const k = soleToken(tokens, s, e);
			const t = tokens[k];
			return k >= 0 && t.type === 'name' && !t.sheet && t.text.toLowerCase() === lower;
		});
		if (declares) {
			out.push([call.open, call.close < 0 ? tokens.length : call.close + 1]);
		}
	}
	return out;
}

/**
 * Replace uses of the defined name `oldName` with `newName` in a formula.
 * Only name tokens are touched (never strings, functions of the same
 * spelling or LET / LAMBDA parameters); calls of a LAMBDA name such as
 * `MyFn(1)` are renamed too. `matches` filters uses by their qualifier.
 */
export function renameDefinedNameInFormula(
	formula: string,
	oldName: string,
	newName: string,
	matches: NameUseMatcher = () => true,
): string {
	const lower = oldName.toLowerCase();
	if (lower === '' || !formula.toLowerCase().includes(lower)) {
		return formula;
	}
	const tokens = tokenizeFormula(formula);
	let shadows: Array<[number, number]> | undefined;
	const edits: TextEdit[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		const isCall = t.type === 'function' && !getFunctionInfo(t.text);
		if (t.type !== 'name' && !isCall) {
			continue;
		}
		const q = t.qualifierLength ?? 0;
		if (t.text.slice(q).toLowerCase() !== lower) {
			continue;
		}
		const sheet = t.sheet;
		if (sheet !== undefined && (sheet.workbook !== undefined || sheet.sheetEnd !== undefined)) {
			continue;
		}
		if (sheet === undefined) {
			shadows ??= shadowingCalls(tokens, lower);
			if (shadows.some(([s, e]) => i >= s && i < e)) {
				continue;
			}
		}
		if (matches(sheet?.sheet)) {
			edits.push({ start: t.start + q, end: t.end, text: newName });
		}
	}
	return applyEdits(formula, edits);
}

type Rewrite = (formula: string) => string;

function rewriteCfvos(cfvos: Cfvo[], rewrite: Rewrite): Cfvo[] {
	let changed = false;
	const out = cfvos.map((c) => {
		if (c.type !== 'formula' || c.value === undefined) {
			return c;
		}
		const value = rewrite(c.value);
		if (value === c.value) {
			return c;
		}
		changed = true;
		return { ...c, value };
	});
	return changed ? out : cfvos;
}

function rewriteRule(rule: CfRule, rewrite: Rewrite): CfRule {
	const next: CfRule = { ...rule };
	if (rule.formulas) {
		const formulas = rule.formulas.map(rewrite);
		if (formulas.some((f, i) => f !== rule.formulas![i])) {
			next.formulas = formulas;
		}
	}
	if (rule.colorScale) {
		const cfvos = rewriteCfvos(rule.colorScale.cfvos, rewrite);
		if (cfvos !== rule.colorScale.cfvos) {
			next.colorScale = { ...rule.colorScale, cfvos };
		}
	}
	if (rule.dataBar) {
		const cfvos = rewriteCfvos(rule.dataBar.cfvos, rewrite);
		if (cfvos !== rule.dataBar.cfvos) {
			next.dataBar = { ...rule.dataBar, cfvos: [cfvos[0], cfvos[1]] };
		}
	}
	if (rule.iconSet) {
		const cfvos = rewriteCfvos(rule.iconSet.cfvos, rewrite);
		if (cfvos !== rule.iconSet.cfvos) {
			next.iconSet = { ...rule.iconSet, cfvos };
		}
	}
	const keys = ['formulas', 'colorScale', 'dataBar', 'iconSet'] as const;
	return keys.some((k) => next[k] !== rule[k]) ? next : rule;
}

/** Conditional formats with renamed formulas; the input array when nothing changed. */
function rewriteFormats(formats: ConditionalFormat[], rewrite: Rewrite): ConditionalFormat[] {
	let changed = false;
	const out = formats.map((cf) => {
		const rules = cf.rules.map((r) => rewriteRule(r, rewrite));
		if (rules.every((r, i) => r === cf.rules[i])) {
			return cf;
		}
		changed = true;
		return { ...cf, rules };
	});
	return changed ? out : formats;
}

function rewriteValidations(list: DataValidation[], rewrite: Rewrite): DataValidation[] {
	let changed = false;
	const out = list.map((dv) => {
		const formula1 = dv.formula1 === undefined ? undefined : rewrite(dv.formula1);
		const formula2 = dv.formula2 === undefined ? undefined : rewrite(dv.formula2);
		if (formula1 === dv.formula1 && formula2 === dv.formula2) {
			return dv;
		}
		changed = true;
		return { ...dv, formula1, formula2 };
	});
	return changed ? out : list;
}

function rewriteColumn(column: TableColumn, rewrite: Rewrite): TableColumn {
	const next: TableColumn = { ...column };
	if (column.calculatedColumnFormula !== undefined) {
		next.calculatedColumnFormula = rewrite(column.calculatedColumnFormula);
	}
	if (column.totalsRowFormula !== undefined) {
		next.totalsRowFormula = rewrite(column.totalsRowFormula);
	}
	const same =
		next.calculatedColumnFormula === column.calculatedColumnFormula &&
		next.totalsRowFormula === column.totalsRowFormula;
	return same ? column : next;
}

/** Tables whose calculated-column or totals formulas changed are replaced and marked dirty. */
function rewriteTables(tables: Table[], rewrite: Rewrite): Table[] {
	let changed = false;
	const out = tables.map((table) => {
		const columns = table.columns.map((c) => rewriteColumn(c, rewrite));
		if (columns.every((c, i) => c === table.columns[i])) {
			return table;
		}
		changed = true;
		return { ...table, columns, dirty: true };
	});
	return changed ? out : tables;
}

/**
 * Rewrite every use of `def` (still holding its old name) to `newName`
 * across the workbook. The caller must already have called
 * `tx.touchWorkbook()`; sheets whose formulas change are touched here.
 * Uses that resolve to a different definition are left alone: on a sheet
 * with its own name of the same spelling, plain uses mean that one.
 */
export function rewriteNameReferences(
	tx: EditTransaction,
	workbook: Workbook,
	def: DefinedName,
	newName: string,
): void {
	const oldName = def.name;
	const lower = oldName.toLowerCase();
	const scope = def.localSheet;
	const hasLocal = (sheet: number): boolean =>
		workbook.definedNames.some(
			(n) => n !== def && n.localSheet === sheet && n.name.toLowerCase() === lower,
		);
	/** Uses inside formulas evaluated on sheet `context` (undefined: a workbook-level name). */
	const matcherFor =
		(context: number | undefined): NameUseMatcher =>
		(qualifier) => {
			if (qualifier !== undefined) {
				const sheet = findSheetIndex(workbook, qualifier);
				return sheet >= 0 && (scope === undefined ? !hasLocal(sheet) : sheet === scope);
			}
			if (scope === undefined) {
				return context === undefined || !hasLocal(context);
			}
			return context === scope;
		};

	workbook.sheets.forEach((sheet, index) => {
		const match = matcherFor(index);
		const rewrite: Rewrite = (f) => renameDefinedNameInFormula(f, oldName, newName, match);
		const cells: Array<[number, Cell]> = [];
		for (const [key, cell] of sheet.cells) {
			if (cell.formula) {
				const formula = rewrite(cell.formula);
				if (formula !== cell.formula) {
					cells.push([key, { ...cell, formula }]);
				}
			}
		}
		const formats = rewriteFormats(sheet.conditionalFormats, rewrite);
		const validations = rewriteValidations(sheet.dataValidations, rewrite);
		const tables = rewriteTables(sheet.tables, rewrite);
		if (
			cells.length === 0 &&
			formats === sheet.conditionalFormats &&
			validations === sheet.dataValidations &&
			tables === sheet.tables
		) {
			return;
		}
		tx.touchSheet(index);
		for (const [key, cell] of cells) {
			sheet.cells.set(key, cell);
		}
		sheet.conditionalFormats = formats;
		sheet.dataValidations = validations;
		sheet.tables = tables;
		sheet.dirty = true;
	});

	for (const n of workbook.definedNames) {
		const formula = renameDefinedNameInFormula(
			n.formula,
			oldName,
			newName,
			matcherFor(n.localSheet),
		);
		if (formula !== n.formula) {
			n.formula = formula;
		}
	}
}
