/**
 * One worksheet substream -> `Worksheet`. Records are routed by type to the
 * cell, formula, layout, view, link, note, rule and settings readers; a
 * record that fails to parse is reported and skipped.
 */
import { createWorksheet } from '../../sheet/worksheet';
import type { Worksheet } from '../../types/sheet';
import type { CalcProperties } from '../../types/workbook';
import { BOF_TYPE, RT } from './biff-opcodes';
import type { BiffRecord, Substream } from './biff-records';
import type { FormulaEnv } from './formula';
import { readValueCell, VALUE_CELL_RECORDS } from './sheet-cells';
import { type SheetContext, sheetWarn } from './sheet-context';
import { FORMULA_RECORDS, FormulaCells } from './sheet-formulas';
import { finishLayout, LAYOUT_RECORDS, readLayout } from './sheet-layout';
import { LINK_RECORDS, readLink } from './sheet-links';
import { NOTE_RECORDS, NoteReader } from './sheet-notes';
import { SETTINGS_RECORDS, SettingsReader } from './sheet-print';
import { RULE_RECORDS, RuleReader } from './sheet-rules';
import { VIEW_RECORDS, ViewReader } from './sheet-view';
import type { XlsSharedStrings } from './sst';
import type { XlsWarnings } from './warnings';
import type { XlsStyles } from './xls-styles';

export interface SheetInput {
	name: string;
	sheetId: number;
	state: Worksheet['state'];
}

export interface SheetShared {
	styles: XlsStyles;
	sst: XlsSharedStrings;
	env: FormulaEnv;
	warnings: XlsWarnings;
}

export interface SheetResult {
	sheet: Worksheet;
	calc: CalcProperties;
	/** The sheet has an AutoFilter (range from its `_FilterDatabase` name). */
	autoFilter: boolean;
	/** WSBOOL marks the substream as an Excel 5 dialog sheet. */
	dialog: boolean;
}

/** Records allowed between a FORMULA and the STRING record holding its result. */
const BEFORE_STRING = new Set<number>([RT.STRING, RT.SHRFMLA, RT.ARRAY, RT.TABLE]);

function reportSkipped(ctx: SheetContext, nested: readonly number[], notes: NoteReader): void {
	const name = ctx.sheet.name;
	const charts = notes.skipped.charts + nested.filter((t) => t === BOF_TYPE.chart).length;
	if (charts > 0) {
		sheetWarn(ctx, 'chart-skipped', `Charts on "${name}" are not imported from .xls files`);
	}
	if (notes.skipped.pictures > 0) {
		sheetWarn(ctx, 'image-skipped', `Pictures on "${name}" are not imported from .xls files`);
	}
	if (notes.skipped.other > 0) {
		sheetWarn(
			ctx,
			'drawing-skipped',
			`Shapes, text boxes and controls on "${name}" are not imported from .xls files`,
		);
	}
}

/** Read a worksheet substream (as returned by `readSubstream`). */
export function readWorksheet(
	stream: Substream,
	input: SheetInput,
	shared: SheetShared,
): SheetResult {
	const sheet = createWorksheet(input.name, input.sheetId);
	sheet.state = input.state;
	const ctx: SheetContext = { sheet, ...shared };
	const formulas = new FormulaCells(ctx);
	const view = new ViewReader();
	const notes = new NoteReader(ctx);
	const rules = new RuleReader(ctx);
	const settings = new SettingsReader(ctx);
	const route = new Map<number, (r: BiffRecord) => unknown>();
	const add = (types: readonly number[], fn: (r: BiffRecord) => unknown): void => {
		for (const t of types) {
			route.set(t, fn);
		}
	};
	add(VALUE_CELL_RECORDS, (r) => readValueCell(ctx, r));
	add(FORMULA_RECORDS, (r) => formulas.read(r));
	add(LAYOUT_RECORDS, (r) => readLayout(ctx, r));
	add(VIEW_RECORDS, (r) => view.read(r));
	add(LINK_RECORDS, (r) => readLink(sheet.hyperlinks, r));
	add(NOTE_RECORDS, (r) => notes.read(r));
	add(RULE_RECORDS, (r) => rules.read(r));
	add(SETTINGS_RECORDS, (r) => settings.read(r));

	let dialog = false;
	for (let i = 1; i < stream.records.length; i++) {
		const record = stream.records[i];
		if (!BEFORE_STRING.has(record.type)) {
			formulas.endString();
		}
		if (record.type === RT.WSBOOL && record.data.length >= 1 && record.data[0] & 0x10) {
			dialog = true;
		}
		const handler = route.get(record.type);
		if (!handler) {
			continue;
		}
		try {
			handler(record);
		} catch (e) {
			shared.warnings.corrupt(
				`record 0x${record.type.toString(16)} on "${sheet.name}"`,
				e,
				sheet.name,
			);
		}
	}
	formulas.finish();
	finishLayout(ctx);
	sheet.view = view.view();
	sheet.notes = notes.finish();
	rules.finish();
	settings.finish();
	reportSkipped(ctx, stream.nested, notes);
	if (stream.truncated) {
		sheetWarn(ctx, 'truncated-record', `The sheet "${sheet.name}" is incomplete in the file`);
	}
	return { sheet, calc: settings.calc, autoFilter: settings.autoFilter, dialog };
}
