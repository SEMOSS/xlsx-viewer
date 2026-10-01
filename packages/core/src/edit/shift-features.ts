import type { RangeAddress } from '../types/address';
import type { ChartModel } from '../types/chart';
import type { ConditionalFormat } from '../types/conditional-format';
import type { DataValidation } from '../types/data-validation';
import type { DrawingObject } from '../types/drawing';
import type { SparklineGroup, Worksheet } from '../types/sheet';
import type { Table } from '../types/table';
import type { Workbook } from '../types/workbook';
import {
	adjustFormulaForBand,
	adjustRangeForBand,
	bandFateOf,
	type BandZones,
	type CellBandShift,
	movePointForBand,
} from './shift-band';

/**
 * Insert / delete cells with shifting: everything except the cells
 * themselves (formulas on every sheet, names, rules, validations, links,
 * notes, tables, charts and sparklines).
 */

type Fix = (formula: string) => string;

/** What happens to a table: kept, removed with the deleted cells, moved / resized, or refused. */
export type TableFate =
	| { kind: 'keep' }
	| { kind: 'remove' }
	| { kind: 'resize'; ref: RangeAddress }
	| { kind: 'split' };

/**
 * Tables move with the shifted block or disappear with the deleted cells.
 * A vertical shift spanning the whole table width inside it inserts or
 * deletes table rows (never the header or totals row, and at least one data
 * row stays); any other overlap would split the table and is refused.
 */
export function tableFateForBand(table: Table, band: CellBandShift, zones: BandZones): TableFate {
	const fate = bandFateOf(table.ref, zones);
	if (fate === 'outside') {
		return { kind: 'keep' };
	}
	if (fate === 'removed') {
		return { kind: 'remove' };
	}
	const r = table.ref;
	if (fate === 'split') {
		if (!band.vertical || r.left < band.from || r.right > band.to) {
			return { kind: 'split' };
		}
		if (!band.insert) {
			const firstData = r.top + table.headerRowCount;
			const lastData = r.bottom - table.totalsRowCount;
			const last = band.at + band.count - 1;
			if (band.at < firstData || last > lastData || band.count > lastData - firstData) {
				return { kind: 'split' };
			}
		}
	}
	const ref = adjustRangeForBand(r, band);
	return ref ? { kind: 'resize', ref } : { kind: 'split' };
}

function fixRef<T extends { ref?: string }>(
	o: T | undefined,
	fix: Fix,
	hit: () => void,
): T | undefined {
	if (!o?.ref) {
		return o;
	}
	const ref = fix(o.ref);
	if (ref === o.ref) {
		return o;
	}
	hit();
	return { ...o, ref };
}

function adjustChart(chart: ChartModel, fix: Fix): ChartModel {
	let changed = false;
	const hit = () => {
		changed = true;
	};
	const plots = chart.plots.map((p) => ({
		...p,
		series: p.series.map((s) => ({
			...s,
			name: fixRef(s.name, fix, hit),
			categories: fixRef(s.categories, fix, hit),
			values: fixRef(s.values, fix, hit),
			bubbleSizes: fixRef(s.bubbleSizes, fix, hit),
		})),
	}));
	const title = fixRef(chart.title, fix, hit);
	return changed ? { ...chart, title, plots, dirty: true } : chart;
}

/** Charts (also inside groups) with adjusted series references; `null` when nothing changed. */
function adjustDrawings(objects: DrawingObject[], fix: Fix): DrawingObject[] | null {
	let changed = false;
	const next = objects.map((d): DrawingObject => {
		if (d.kind === 'chart') {
			const chart = adjustChart(d.chart, fix);
			if (chart !== d.chart) {
				changed = true;
				return { ...d, chart };
			}
		} else if (d.kind === 'group') {
			const children = adjustDrawings(d.children, fix);
			if (children) {
				changed = true;
				return { ...d, children };
			}
		}
		return d;
	});
	return changed ? next : null;
}

function adjustRuleFormulas(cf: ConditionalFormat, fix: Fix): ConditionalFormat {
	let changed = false;
	const rules = cf.rules.map((rule) => {
		const formulas = rule.formulas?.map(fix);
		if (!formulas || formulas.every((f, i) => f === rule.formulas?.[i])) {
			return rule;
		}
		changed = true;
		return { ...rule, formulas };
	});
	return changed ? { ...cf, rules } : cf;
}

function adjustValidationFormulas(dv: DataValidation, fix: Fix): DataValidation {
	const formula1 = dv.formula1 !== undefined ? fix(dv.formula1) : undefined;
	const formula2 = dv.formula2 !== undefined ? fix(dv.formula2) : undefined;
	return formula1 === dv.formula1 && formula2 === dv.formula2 ? dv : { ...dv, formula1, formula2 };
}

/** Sparkline formulas; on the shifted sheet the locations move with their cells too. */
function adjustSparklines(
	groups: SparklineGroup[],
	fix: Fix,
	band: CellBandShift | null,
): SparklineGroup[] {
	return groups.map((g) => {
		let changed = false;
		const sparklines = g.sparklines.flatMap((s) => {
			const location = band ? movePointForBand(s.location, band) : s.location;
			if (!location) {
				changed = true;
				return [];
			}
			const formula = fix(s.formula);
			if (location === s.location && formula === s.formula) {
				return [s];
			}
			changed = true;
			return [{ ...s, location, formula }];
		});
		return changed ? { ...g, sparklines } : g;
	});
}

function adjustSheetFormulas(sheet: Worksheet, band: CellBandShift, home: boolean): void {
	const fix: Fix = (f) => adjustFormulaForBand(f, band, sheet.name);
	for (const [key, cell] of sheet.cells) {
		// Formulas of other sheets need a sheet qualifier to reach the shifted cells.
		if (cell.formula === undefined || (!home && !cell.formula.includes('!'))) {
			continue;
		}
		const formula = fix(cell.formula);
		if (formula !== cell.formula) {
			sheet.cells.set(key, { ...cell, formula });
		}
	}
	sheet.conditionalFormats = sheet.conditionalFormats.map((cf) => adjustRuleFormulas(cf, fix));
	sheet.dataValidations = sheet.dataValidations.map((dv) => adjustValidationFormulas(dv, fix));
	sheet.sparklineGroups = adjustSparklines(sheet.sparklineGroups, fix, home ? band : null);
	const drawings = adjustDrawings(sheet.drawings, fix);
	if (drawings) {
		sheet.drawings = drawings;
		sheet.drawingsDirty = true;
	}
}

/** Rewrite every formula of the workbook (cells, names, rules, validations, charts, sparklines). */
export function adjustWorkbookForBand(
	workbook: Workbook,
	band: CellBandShift,
	home: Worksheet,
): void {
	for (const sheet of workbook.sheets) {
		adjustSheetFormulas(sheet, band, sheet === home);
	}
	const names = workbook.definedNames;
	for (let i = 0; i < names.length; i++) {
		const n = names[i];
		const scope =
			n.localSheet !== undefined ? (workbook.sheets[n.localSheet]?.name ?? home.name) : home.name;
		const formula = adjustFormulaForBand(n.formula, band, scope);
		if (formula !== n.formula) {
			names[i] = { ...n, formula };
		}
	}
}

function adjustRanges(ranges: RangeAddress[], band: CellBandShift): RangeAddress[] {
	return ranges.flatMap((r) => {
		const next = adjustRangeForBand(r, band);
		return next ? [next] : [];
	});
}

/** Shifted sheet: merges, links, rule / validation ranges, filters, tables and notes. */
export function adjustHomeSheetForBand(
	sheet: Worksheet,
	band: CellBandShift,
	zones: BandZones,
): void {
	sheet.merges = adjustRanges(sheet.merges, band).filter(
		(m) => m.top !== m.bottom || m.left !== m.right,
	);
	sheet.hyperlinks = sheet.hyperlinks.flatMap((h) => {
		const ref = adjustRangeForBand(h.ref, band);
		return ref ? [ref === h.ref ? h : { ...h, ref }] : [];
	});
	sheet.conditionalFormats = sheet.conditionalFormats.flatMap((cf) => {
		const ranges = adjustRanges(cf.ranges, band);
		return ranges.length === 0 ? [] : [{ ...cf, ranges }];
	});
	sheet.dataValidations = sheet.dataValidations.flatMap((dv) => {
		const ranges = adjustRanges(dv.ranges, band);
		return ranges.length === 0 ? [] : [{ ...dv, ranges }];
	});
	if (sheet.autoFilter) {
		const ref = adjustRangeForBand(sheet.autoFilter.ref, band);
		sheet.autoFilter = ref ? { ...sheet.autoFilter, ref } : undefined;
	}
	sheet.tables = sheet.tables.flatMap((t) => {
		const fate = tableFateForBand(t, band, zones);
		if (fate.kind === 'remove') {
			return [];
		}
		if (fate.kind !== 'resize' || fate.ref === t.ref) {
			return [t];
		}
		const ref = fate.ref;
		const autoFilter = t.autoFilter
			? { ...t.autoFilter, ref: { ...ref, bottom: ref.bottom - (t.totalsRowCount ? 1 : 0) } }
			: undefined;
		return [{ ...t, ref, autoFilter, dirty: true }];
	});
	let comments = false;
	sheet.notes = sheet.notes.flatMap((n) => {
		const ref = movePointForBand(n.ref, band);
		comments ||= ref !== n.ref;
		return ref ? [ref === n.ref ? n : { ...n, ref }] : [];
	});
	sheet.threadedComments = sheet.threadedComments.flatMap((c) => {
		const ref = movePointForBand(c.ref, band);
		comments ||= ref !== c.ref;
		return ref ? [ref === c.ref ? c : { ...c, ref }] : [];
	});
	if (comments) {
		sheet.commentsDirty = true;
	}
}
