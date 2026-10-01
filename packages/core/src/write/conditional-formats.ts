/**
 * Conditional formatting: `<conditionalFormatting>` blocks plus the Excel
 * 2010 `x14:conditionalFormattings` extension for rules that need it
 * (data bar extensions, x14-only icon sets, extension-only rules).
 */
import { a1 } from '../address/cell-address';
import { formatSqref } from '../address/range';
import type {
	CfDataBar,
	CfIconSet,
	CfRule,
	Cfvo,
	ConditionalFormat,
} from '../types/conditional-format';
import type { DifferentialFormat } from '../types/style';
import { el, textEl } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { defaultCfFormula, fileFormula } from './formulas';
import { dxfInnerXml } from './style-entries';
import { colorAttrs, flag, newGuid, unflag } from './xml-utils';

export const CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';
export const DATABAR_EXT_URI = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';

/** Icon sets that only exist in the Excel 2010 schema. */
const X14_ICON_SETS = new Set(['3Stars', '3Triangles', '5Boxes', 'NoIcons']);
const OPERATOR_TYPES = new Set([
	'cellIs',
	'containsText',
	'notContainsText',
	'beginsWith',
	'endsWith',
]);
const DEFAULT_TEXT_OPERATOR: Record<string, string> = {
	containsText: 'containsText',
	notContainsText: 'notContains',
	beginsWith: 'beginsWith',
	endsWith: 'endsWith',
};

/** Differential formats while writing rules. */
export interface CfStyles {
	/** The current `dxfs` (it grows when inline formats are interned). */
	dxfs(): readonly DifferentialFormat[];
	/** Index of an inline format in `dxfs`, appended when new. */
	internDxf(dxf: DifferentialFormat): number;
}

export interface CfOutput {
	elements: XmlElement[];
	/** `<ext>` carrying the x14 rules, or `null`. */
	ext: XmlElement | null;
	/** Rules skipped because they were incomplete. */
	skipped: number;
}

function isExtOnly(rule: CfRule): boolean {
	if (rule.fromExtension) {
		return true;
	}
	const is = rule.iconSet;
	return (
		rule.type === 'iconSet' && !!is && (X14_ICON_SETS.has(is.iconSet) || !!is.customIcons?.length)
	);
}

function hasDataBarExt(db: CfDataBar): boolean {
	return (
		db.gradient !== undefined ||
		db.border !== undefined ||
		!!db.borderColor ||
		!!db.negativeFillColor ||
		!!db.negativeBorderColor ||
		db.negativeBarColorSameAsPositive !== undefined ||
		!!db.axisPosition ||
		!!db.axisColor ||
		!!db.direction ||
		db.cfvos.some((c) => c.type === 'autoMin' || c.type === 'autoMax')
	);
}

function needsValue(type: Cfvo['type']): boolean {
	return type === 'num' || type === 'percent' || type === 'percentile' || type === 'formula';
}

function cfvoValue(c: Cfvo): string {
	const v = c.value ?? '0';
	return c.type === 'formula' ? fileFormula(v) : v;
}

function stdCfvo(c: Cfvo): XmlElement {
	const type = c.type === 'autoMin' ? 'min' : c.type === 'autoMax' ? 'max' : c.type;
	return el('cfvo', { type, val: needsValue(type) ? cfvoValue(c) : undefined, gte: unflag(c.gte) });
}

function x14Cfvo(c: Cfvo): XmlElement {
	return el(
		'x14:cfvo',
		{ type: c.type, gte: unflag(c.gte) },
		needsValue(c.type) ? [textEl('xm:f', cfvoValue(c))] : [],
	);
}

function iconSetAttrs(is: CfIconSet): Record<string, string | undefined> {
	return {
		iconSet: is.iconSet,
		showValue: unflag(is.showValue),
		percent: unflag(is.percent),
		reverse: flag(is.reverse),
	};
}

/** Rule formulas, falling back to the formula Excel stores for text / period rules. */
function ruleFormulas(rule: CfRule, cell: string): string[] {
	if (rule.formulas?.length) {
		return rule.formulas.slice(0, 3);
	}
	const f = defaultCfFormula(rule, cell);
	return f ? [f] : [];
}

function commonAttrs(rule: CfRule): Record<string, string | number | boolean | undefined> {
	const operator = OPERATOR_TYPES.has(rule.type)
		? (rule.operator ?? DEFAULT_TEXT_OPERATOR[rule.type])
		: undefined;
	return {
		type: rule.type,
		priority: Math.max(1, Math.trunc(rule.priority || 1)),
		stopIfTrue: flag(rule.stopIfTrue),
		aboveAverage: rule.type === 'aboveAverage' ? unflag(rule.aboveAverage) : undefined,
		percent: rule.type === 'top10' ? flag(rule.percent) : undefined,
		bottom: rule.type === 'top10' ? flag(rule.bottom) : undefined,
		operator,
		text: rule.text,
		timePeriod: rule.type === 'timePeriod' ? rule.timePeriod : undefined,
		rank: rule.type === 'top10' ? (rule.rank ?? 10) : undefined,
		stdDev: rule.type === 'aboveAverage' ? rule.stdDev : undefined,
		equalAverage: rule.type === 'aboveAverage' ? flag(rule.equalAverage) : undefined,
	};
}

/** The rule has what its type needs (formulas, a color scale, ...). */
function isComplete(rule: CfRule, formulas: string[]): boolean {
	switch (rule.type) {
		case 'expression':
		case 'cellIs':
			return formulas.length > 0;
		case 'colorScale':
			return !!rule.colorScale && rule.colorScale.cfvos.length >= 2;
		case 'dataBar':
			return !!rule.dataBar && rule.dataBar.cfvos.length === 2;
		case 'iconSet':
			return !!rule.iconSet && rule.iconSet.cfvos.length >= 2;
		case 'timePeriod':
			return formulas.length > 0 && !!rule.timePeriod;
		default:
			return true;
	}
}

/**
 * `dxfId` of a 2007 rule. An inline format wins (as in the evaluation
 * engine) and is interned into the stylesheet; a dangling id is dropped.
 */
function stdDxfId(rule: CfRule, styles: CfStyles): number | undefined {
	if (rule.dxf) {
		return styles.internDxf(rule.dxf);
	}
	const id = rule.dxfId;
	return id !== undefined && Number.isInteger(id) && id >= 0 && id < styles.dxfs().length
		? id
		: undefined;
}

function stdRule(rule: CfRule, formulas: string[], styles: CfStyles, extId?: string): XmlElement {
	const kids: XmlElement[] = formulas.map((f) => textEl('formula', fileFormula(f)));
	if (rule.type === 'colorScale' && rule.colorScale) {
		const cs = rule.colorScale;
		kids.push(
			el('colorScale', {}, [
				...cs.cfvos.map(stdCfvo),
				...cs.colors.map((c) => el('color', colorAttrs(c))),
			]),
		);
	} else if (rule.type === 'dataBar' && rule.dataBar) {
		const db = rule.dataBar;
		kids.push(
			el(
				'dataBar',
				{ minLength: db.minLength, maxLength: db.maxLength, showValue: unflag(db.showValue) },
				[stdCfvo(db.cfvos[0]), stdCfvo(db.cfvos[1]), el('color', colorAttrs(db.color))],
			),
		);
		if (extId) {
			kids.push(
				el('extLst', {}, [
					el('ext', { uri: DATABAR_EXT_URI, 'xmlns:x14': NS.x14 }, [textEl('x14:id', extId)]),
				]),
			);
		}
	} else if (rule.type === 'iconSet' && rule.iconSet) {
		kids.push(el('iconSet', iconSetAttrs(rule.iconSet), rule.iconSet.cfvos.map(stdCfvo)));
	}
	const attrs = commonAttrs(rule);
	return el('cfRule', { type: attrs.type, dxfId: stdDxfId(rule, styles), ...attrs }, kids);
}

function x14DataBar(db: CfDataBar, withFill: boolean): XmlElement {
	return el(
		'x14:dataBar',
		{
			minLength: db.minLength,
			maxLength: db.maxLength,
			showValue: unflag(db.showValue),
			border: flag(db.border),
			gradient: unflag(db.gradient),
			direction: db.direction && db.direction !== 'context' ? db.direction : undefined,
			negativeBarColorSameAsPositive: flag(db.negativeBarColorSameAsPositive),
			axisPosition:
				db.axisPosition && db.axisPosition !== 'automatic' ? db.axisPosition : undefined,
		},
		[
			x14Cfvo(db.cfvos[0]),
			x14Cfvo(db.cfvos[1]),
			withFill ? el('x14:fillColor', colorAttrs(db.color)) : undefined,
			db.borderColor ? el('x14:borderColor', colorAttrs(db.borderColor)) : undefined,
			db.negativeFillColor
				? el('x14:negativeFillColor', colorAttrs(db.negativeFillColor))
				: undefined,
			db.negativeBorderColor
				? el('x14:negativeBorderColor', colorAttrs(db.negativeBorderColor))
				: undefined,
			db.axisColor ? el('x14:axisColor', colorAttrs(db.axisColor)) : undefined,
		],
	);
}

function x14Rule(
	rule: CfRule,
	formulas: string[],
	id: string,
	styles: CfStyles,
	counterpart: boolean,
): XmlElement {
	const kids: XmlElement[] = [];
	if (!counterpart) {
		kids.push(...formulas.map((f) => textEl('xm:f', fileFormula(f))));
	}
	if (rule.type === 'colorScale' && rule.colorScale) {
		const cs = rule.colorScale;
		kids.push(
			el('x14:colorScale', {}, [
				...cs.cfvos.map(x14Cfvo),
				...cs.colors.map((c) => el('x14:color', colorAttrs(c))),
			]),
		);
	} else if (rule.type === 'dataBar' && rule.dataBar) {
		kids.push(x14DataBar(rule.dataBar, !counterpart));
	} else if (rule.type === 'iconSet' && rule.iconSet) {
		const is = rule.iconSet;
		kids.push(
			el('x14:iconSet', { ...iconSetAttrs(is), custom: is.customIcons?.length ? 1 : undefined }, [
				...is.cfvos.map(x14Cfvo),
				...(is.customIcons ?? []).map((ci) =>
					el('x14:cfIcon', { iconSet: ci.iconSet, iconId: ci.iconId }),
				),
			]),
		);
	}
	const dxf = rule.dxf ?? (rule.dxfId !== undefined ? styles.dxfs()[rule.dxfId] : undefined);
	if (!counterpart && dxf) {
		const inner = dxfInnerXml(dxf);
		kids.push({ name: 'x14:dxf', attrs: {}, children: inner ? [{ raw: inner }] : [] });
	}
	const attrs = counterpart ? { type: rule.type } : commonAttrs(rule);
	return el('x14:cfRule', { ...attrs, id }, kids);
}

/** Serialize a sheet's conditional formats. */
export function buildConditionalFormats(
	formats: readonly ConditionalFormat[],
	styles: CfStyles,
): CfOutput {
	const elements: XmlElement[] = [];
	const extBlocks: XmlElement[] = [];
	let skipped = 0;
	for (const cf of formats) {
		if (!cf.ranges.length) {
			skipped += cf.rules.length;
			continue;
		}
		const first = cf.ranges[0];
		const cell = a1(first.top, first.left);
		const sqref = formatSqref(cf.ranges);
		const std: XmlElement[] = [];
		const ext: XmlElement[] = [];
		for (const rule of cf.rules) {
			const formulas = ruleFormulas(rule, cell);
			if (!isComplete(rule, formulas)) {
				skipped++;
				continue;
			}
			if (isExtOnly(rule)) {
				ext.push(x14Rule(rule, formulas, rule.extId ?? newGuid(), styles, false));
			} else if (
				rule.type === 'dataBar' &&
				rule.dataBar &&
				(rule.extId || hasDataBarExt(rule.dataBar))
			) {
				const id = rule.extId ?? newGuid();
				std.push(stdRule(rule, formulas, styles, id));
				ext.push(x14Rule(rule, formulas, id, styles, true));
			} else {
				std.push(stdRule(rule, formulas, styles));
			}
		}
		if (std.length) {
			elements.push(el('conditionalFormatting', { sqref, pivot: flag(cf.pivot) }, std));
		}
		if (ext.length) {
			extBlocks.push(
				el('x14:conditionalFormatting', { 'xmlns:xm': NS.xm, pivot: flag(cf.pivot) }, [
					...ext,
					textEl('xm:sqref', sqref),
				]),
			);
		}
	}
	const ext = extBlocks.length
		? el('ext', { uri: CF_EXT_URI, 'xmlns:x14': NS.x14 }, [
				el('x14:conditionalFormattings', {}, extBlocks),
			])
		: null;
	return { elements, ext, skipped };
}
