import type { HostTable, NameDefinition } from '../formula/evaluator/host';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { FormulaCache, ParsedFormula } from './formula-cache';

/** A defined name with its parsed formula. */
export interface WorkbookName extends NameDefinition {
	readonly parsed: ParsedFormula;
}

const XLNM = '_xlnm.';

/**
 * Defined names by upper-case name: sheet-scoped names shadow workbook
 * names on their sheet. Built-in names (`_xlnm.Print_Area`) are also found
 * without the prefix.
 */
export class NameTable {
	private readonly global = new Map<string, WorkbookName>();
	private readonly local = new Map<string, Map<number, WorkbookName>>();
	private snapshot: Array<[string, string, number | undefined]> = [];

	build(workbook: Workbook, cache: FormulaCache): void {
		this.global.clear();
		this.local.clear();
		this.snapshot = workbook.definedNames.map((d) => [d.name, d.formula, d.localSheet]);
		for (const d of workbook.definedNames) {
			const parsed = cache.get(d.formula.startsWith('=') ? d.formula.slice(1) : d.formula);
			const def: WorkbookName = { name: d.name, ast: parsed.ast, localSheet: d.localSheet, parsed };
			this.register(d.name.toUpperCase(), def);
			if (d.name.slice(0, XLNM.length).toLowerCase() === XLNM) {
				this.register(d.name.slice(XLNM.length).toUpperCase(), def, true);
			}
		}
	}

	private register(key: string, def: WorkbookName, weak = false): void {
		if (def.localSheet === undefined) {
			if (!weak || !this.global.has(key)) {
				this.global.set(key, def);
			}
			return;
		}
		let bySheet = this.local.get(key);
		if (bySheet === undefined) {
			bySheet = new Map();
			this.local.set(key, bySheet);
		}
		if (!weak || !bySheet.has(def.localSheet)) {
			bySheet.set(def.localSheet, def);
		}
	}

	/** The name visible from `sheet`. */
	find(name: string, sheet: number): WorkbookName | undefined {
		const key = name.toUpperCase();
		return this.local.get(key)?.get(sheet) ?? this.global.get(key);
	}

	/** A name with this spelling exists in any scope. */
	has(name: string): boolean {
		const key = name.toUpperCase();
		return this.global.has(key) || this.local.has(key);
	}

	/** The workbook's names differ from the ones this table was built from. */
	changed(workbook: Workbook): boolean {
		const names = workbook.definedNames;
		if (names.length !== this.snapshot.length) {
			return true;
		}
		return names.some((d, i) => {
			const s = this.snapshot[i];
			return d.name !== s[0] || d.formula !== s[1] || d.localSheet !== s[2];
		});
	}
}

function tableSignature(sheets: readonly Worksheet[]): string {
	const parts: string[] = [];
	sheets.forEach((sheet, i) => {
		for (const t of sheet.tables) {
			const r = t.ref;
			parts.push(
				`${i}|${t.name}|${t.displayName}|${r.top},${r.left},${r.bottom},${r.right}|` +
					`${t.headerRowCount},${t.totalsRowCount}|${t.columns.map((c) => c.name).join('\u0001')}`,
			);
		}
	});
	return parts.join('\u0002');
}

/** Tables by (display) name and by position. */
export class TableTable {
	private readonly byName = new Map<string, HostTable>();
	private bySheet: HostTable[][] = [];
	private signature = '';

	build(workbook: Workbook): void {
		this.byName.clear();
		this.bySheet = workbook.sheets.map(() => []);
		workbook.sheets.forEach((sheet, index) => {
			for (const table of sheet.tables) {
				const entry: HostTable = { sheet: index, table };
				this.bySheet[index].push(entry);
				for (const n of [table.displayName, table.name]) {
					if (n && !this.byName.has(n.toUpperCase())) {
						this.byName.set(n.toUpperCase(), entry);
					}
				}
			}
		});
		this.signature = tableSignature(workbook.sheets);
	}

	find(name: string): HostTable | undefined {
		return this.byName.get(name.toUpperCase());
	}

	at(sheet: number, row: number, col: number): HostTable | undefined {
		for (const t of this.bySheet[sheet] ?? []) {
			const r = t.table.ref;
			if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) {
				return t;
			}
		}
		return undefined;
	}

	/** Any table on the sheet overlaps the range. */
	overlaps(sheet: number, top: number, left: number, bottom: number, right: number): boolean {
		for (const t of this.bySheet[sheet] ?? []) {
			const r = t.table.ref;
			if (r.left <= right && left <= r.right && r.top <= bottom && top <= r.bottom) {
				return true;
			}
		}
		return false;
	}

	changed(workbook: Workbook): boolean {
		return tableSignature(workbook.sheets) !== this.signature;
	}
}

/** Identity and names of the sheets, to detect sheet insertions, deletions, moves and renames. */
export class SheetList {
	private sheets: Worksheet[] = [];
	private names: string[] = [];

	build(workbook: Workbook): void {
		this.sheets = [...workbook.sheets];
		this.names = workbook.sheets.map((s) => s.name);
	}

	changed(workbook: Workbook): boolean {
		const s = workbook.sheets;
		return (
			s.length !== this.sheets.length ||
			s.some((sheet, i) => sheet !== this.sheets[i] || sheet.name !== this.names[i])
		);
	}
}
