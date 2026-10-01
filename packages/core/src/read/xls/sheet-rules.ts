/**
 * Data validation (DVAL + DV) and conditional formatting (CONDFMT + CF).
 * Rule formulas are relative to the top-left cell of the first range, as
 * in SpreadsheetML. Excel 2007 rule types stored in CF12 / CONDFMT12
 * records (data bars, color scales, icon sets) are reported, not imported.
 */
import { formatFormulaString } from '../../formula/language/serializer';
import { FIRST_CUSTOM_NUMFMT_ID } from '../../styles/builtin-number-formats';
import type { RangeAddress } from '../../types/address';
import type { CfOperator, CfRule, ConditionalFormat } from '../../types/conditional-format';
import type {
	DataValidation,
	DataValidationOperator,
	DataValidationType,
} from '../../types/data-validation';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import { readDxfn } from './dxf';
import { decodeRgce, renderFormula, XlsExternalReferenceError } from './formula';
import { makeRange, type SheetContext, sheetWarn } from './sheet-context';

const DV_TYPES: readonly DataValidationType[] = [
	'none',
	'whole',
	'decimal',
	'list',
	'date',
	'time',
	'textLength',
	'custom',
];
const DV_OPERATORS: readonly DataValidationOperator[] = [
	'between',
	'notBetween',
	'equal',
	'notEqual',
	'greaterThan',
	'lessThan',
	'greaterThanOrEqual',
	'lessThanOrEqual',
];
const DV_ERROR_STYLES = ['stop', 'warning', 'information'] as const;
const CF_OPERATORS: ReadonlyArray<CfOperator | undefined> = [
	undefined,
	'between',
	'notBetween',
	'equal',
	'notEqual',
	'greaterThan',
	'lessThan',
	'greaterThanOrEqual',
	'lessThanOrEqual',
];

/** `SqRefU`: count and 8-byte ranges. */
function readSqref(cur: BiffCursor): RangeAddress[] {
	const count = Math.min(cur.u16(), Math.floor(cur.remaining / 8));
	const out: RangeAddress[] = [];
	for (let i = 0; i < count; i++) {
		const r = makeRange(cur.u16(), cur.u16(), cur.u16(), cur.u16());
		if (r) {
			out.push(r);
		}
	}
	return out;
}

/** Validation strings: an unset string is stored as a single NUL character. */
function dvString(cur: BiffCursor): string | undefined {
	const s = cur.unicodeString();
	return s === '\u0000' || s === '' ? undefined : s;
}

/** An explicit list (`fStrLookup`): one string token with NUL-separated items. */
function explicitList(rgce: Uint8Array): string | undefined {
	try {
		const t = decodeRgce(rgce)[0];
		return t?.k === 'str' ? formatFormulaString(t.v.split('\u0000').join(',')) : undefined;
	} catch {
		return undefined;
	}
}

/** Record types handled by {@link RuleReader.read}. */
export const RULE_RECORDS: readonly number[] = [
	RT.DVAL,
	RT.DV,
	RT.CONDFMT,
	RT.CF,
	RT.CONDFMT12,
	RT.CF12,
];

export class RuleReader {
	private format: ConditionalFormat | undefined;
	private cfLeft = 0;
	private priority = 0;
	private skippedNew = false;

	constructor(private readonly ctx: SheetContext) {}

	private formula(rgce: Uint8Array, base: RangeAddress): string | undefined {
		if (rgce.length === 0) {
			return undefined;
		}
		try {
			const ptgs = decodeRgce(rgce);
			return renderFormula(ptgs, this.ctx.env, { row: base.top, col: base.left, mode: 'rule' });
		} catch (e) {
			const what =
				e instanceof XlsExternalReferenceError ? `refers to ${e.source}` : 'could not be read';
			sheetWarn(
				this.ctx,
				'rule-formula-skipped',
				`A rule formula on "${this.ctx.sheet.name}" ${what}`,
			);
			return undefined;
		}
	}

	private validation(record: BiffRecord): void {
		const cur = new BiffCursor(record.data, record.breaks);
		const flags = cur.u32();
		const promptTitle = dvString(cur);
		const errorTitle = dvString(cur);
		const prompt = dvString(cur);
		const error = dvString(cur);
		const f1 = cur.bytes(cur.u16() + 2).subarray(2);
		const f2 = cur.bytes(cur.u16() + 2).subarray(2);
		const ranges = readSqref(cur);
		if (ranges.length === 0) {
			return;
		}
		const type = DV_TYPES[flags & 0x0f] ?? 'none';
		const dv: DataValidation = { ranges, type };
		if (type !== 'none' && type !== 'list' && type !== 'custom') {
			dv.operator = DV_OPERATORS[(flags >>> 20) & 0x0f] ?? 'between';
		}
		if (type === 'list' && flags & 0x80) {
			const list = explicitList(f1);
			if (list !== undefined) {
				dv.formula1 = list;
			}
		} else {
			const formula1 = this.formula(f1, ranges[0]);
			const formula2 = this.formula(f2, ranges[0]);
			if (formula1 !== undefined) {
				dv.formula1 = formula1;
			}
			if (formula2 !== undefined) {
				dv.formula2 = formula2;
			}
		}
		if (flags & 0x100) {
			dv.allowBlank = true;
		}
		if (type === 'list') {
			dv.inCellDropdown = !(flags & 0x200);
		}
		if (flags & 0x40000) {
			dv.showInputMessage = true;
		}
		if (flags & 0x80000) {
			dv.showErrorMessage = true;
		}
		const errorStyle = DV_ERROR_STYLES[(flags >>> 4) & 0x07];
		if (errorStyle && errorStyle !== 'stop') {
			dv.errorStyle = errorStyle;
		}
		if (promptTitle !== undefined) {
			dv.promptTitle = promptTitle;
		}
		if (prompt !== undefined) {
			dv.prompt = prompt;
		}
		if (errorTitle !== undefined) {
			dv.errorTitle = errorTitle;
		}
		if (error !== undefined) {
			dv.error = error;
		}
		this.ctx.sheet.dataValidations.push(dv);
	}

	private customFormatId(code: string): number {
		const numFmts = this.ctx.styles.sheet.numFmts;
		for (const [id, c] of Object.entries(numFmts)) {
			if (c === code) {
				return Number(id);
			}
		}
		const ids = Object.keys(numFmts).map(Number);
		const id = Math.max(FIRST_CUSTOM_NUMFMT_ID - 1, ...ids) + 1;
		numFmts[id] = code;
		return id;
	}

	private conditionalFormat(record: BiffRecord): void {
		const cur = new BiffCursor(record.data, record.breaks);
		this.cfLeft = cur.u16();
		cur.skip(2 + 8);
		const ranges = readSqref(cur);
		this.format = ranges.length ? { ranges, rules: [] } : undefined;
		if (this.format) {
			this.ctx.sheet.conditionalFormats.push(this.format);
		}
	}

	private rule(record: BiffRecord): void {
		const format = this.format;
		if (!format || this.cfLeft <= 0) {
			return;
		}
		this.cfLeft--;
		const cur = new BiffCursor(record.data, record.breaks);
		const kind = cur.u8();
		const op = cur.u8();
		const cce1 = cur.u16();
		const cce2 = cur.u16();
		const styles = this.ctx.styles.sheet;
		const dxf = readDxfn(cur, styles.numFmts, (code) => this.customFormatId(code));
		const f1 = this.formula(cur.bytes(cce1), format.ranges[0]);
		const f2 = this.formula(cur.bytes(cce2), format.ranges[0]);
		const formulas = [f1, f2].filter((f): f is string => f !== undefined);
		const operator = kind === 1 ? CF_OPERATORS[op] : undefined;
		if ((kind !== 1 && kind !== 2) || (kind === 1 && !operator) || formulas.length === 0) {
			sheetWarn(
				this.ctx,
				'conditional-format-skipped',
				`A conditional format rule on "${this.ctx.sheet.name}" could not be read`,
			);
			return;
		}
		const rule: CfRule = {
			type: kind === 1 ? 'cellIs' : 'expression',
			priority: ++this.priority,
			stopIfTrue: true,
			dxfId: styles.dxfs.length,
			formulas: kind === 1 ? formulas : formulas.slice(0, 1),
		};
		styles.dxfs.push(dxf);
		if (operator) {
			rule.operator = operator;
		}
		format.rules.push(rule);
	}

	/** Handle a rule record; returns `false` for other record types. */
	read(record: BiffRecord): boolean {
		switch (record.type) {
			case RT.DVAL:
				return true;
			case RT.DV:
				this.validation(record);
				return true;
			case RT.CONDFMT:
				this.conditionalFormat(record);
				return true;
			case RT.CF:
				this.rule(record);
				return true;
			case RT.CONDFMT12:
			case RT.CF12:
				this.skippedNew = true;
				return true;
			default:
				return false;
		}
	}

	finish(): void {
		const sheet = this.ctx.sheet;
		sheet.conditionalFormats = sheet.conditionalFormats.filter((f) => f.rules.length > 0);
		if (this.skippedNew) {
			sheetWarn(
				this.ctx,
				'conditional-format-unsupported',
				`Excel 2007 conditional formats on "${sheet.name}" ` +
					'(data bars, color scales, icon sets) were not imported',
			);
		}
	}
}
