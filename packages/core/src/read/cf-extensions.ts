import { parseSqref } from '../address/range';
import type { CfDataBar, CfRule, ConditionalFormat } from '../types/conditional-format';
import { isElement, textContent } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { parseCfvo, parseDataBar, parseIconSet, parseRuleCommon } from './conditional-formatting';
import { parseColor, parseDxf } from './style-parts';
import { bool, int, kid, kids, lname, oneOf, str } from './xml-utils';

/**
 * Excel 2010 conditional formatting (`extLst` / `x14:conditionalFormattings`).
 *
 * A 2010 rule that extends a 2007 rule carries the same GUID (`x14:cfRule/@id`
 * = `cfRule/extLst/.../x14:id`); its data-bar and icon-set properties are
 * merged into the 2007 rule. Rules that exist only in the extension (for
 * example formulas referring to other sheets, or the 2010-only icon sets)
 * become new rules with `fromExtension: true` and an inline `dxf`.
 */

/** Children of the worksheet `extLst` we understand, located by local name. */
export interface WorksheetExtensions {
	conditionalFormattings?: XmlElement;
	dataValidations?: XmlElement;
	sparklineGroups?: XmlElement;
	/** Local names of every extension payload, for feature warnings (slicers, timelines...). */
	names: Set<string>;
}

export function findWorksheetExtensions(root: XmlElement): WorksheetExtensions {
	const result: WorksheetExtensions = { names: new Set() };
	for (const ext of kids(kid(root, 'extLst'), 'ext')) {
		for (const payload of ext.children) {
			if (!isElement(payload)) {
				continue;
			}
			const name = lname(payload);
			result.names.add(name);
			if (name === 'conditionalFormattings') {
				result.conditionalFormattings ??= payload;
			} else if (name === 'dataValidations') {
				result.dataValidations ??= payload;
			} else if (name === 'sparklineGroups') {
				result.sparklineGroups ??= payload;
			}
		}
	}
	return result;
}

function guidKey(id: string): string {
	return id
		.trim()
		.replace(/^\{|\}$/g, '')
		.toUpperCase();
}

/** Merge `x14:dataBar` properties into a data bar. */
export function applyDataBarExtension(bar: CfDataBar, ext: XmlElement): void {
	const minLength = int(ext, 'minLength');
	if (minLength !== undefined) {
		bar.minLength = minLength;
	}
	const maxLength = int(ext, 'maxLength');
	if (maxLength !== undefined) {
		bar.maxLength = maxLength;
	}
	bar.gradient = bool(ext, 'gradient') ?? true;
	const border = bool(ext, 'border');
	if (border !== undefined) {
		bar.border = border;
	}
	const direction = oneOf(str(ext, 'direction'), [
		'context',
		'leftToRight',
		'rightToLeft',
	] as const);
	if (direction) {
		bar.direction = direction;
	}
	const same = bool(ext, 'negativeBarColorSameAsPositive');
	if (same !== undefined) {
		bar.negativeBarColorSameAsPositive = same;
	}
	const axis = oneOf(str(ext, 'axisPosition'), ['automatic', 'middle', 'none'] as const);
	if (axis) {
		bar.axisPosition = axis;
	}
	const cfvos = kids(ext, 'cfvo').map(parseCfvo);
	if (cfvos.length >= 2) {
		bar.cfvos = [cfvos[0], cfvos[1]];
	}
	const fill = parseColor(kid(ext, 'fillColor'));
	if (fill) {
		bar.color = fill;
	}
	const colors = [
		['borderColor', 'borderColor'],
		['negativeFillColor', 'negativeFillColor'],
		['negativeBorderColor', 'negativeBorderColor'],
		['axisColor', 'axisColor'],
	] as const;
	for (const [local, key] of colors) {
		const c = parseColor(kid(ext, local));
		if (c) {
			bar[key] = c;
		}
	}
}

function mergeIntoRule(target: CfRule, ext: XmlElement): void {
	const bar = kid(ext, 'dataBar');
	if (bar) {
		target.dataBar ??= parseDataBar(bar);
		applyDataBarExtension(target.dataBar, bar);
	}
	const icons = kid(ext, 'iconSet');
	if (icons) {
		const parsed = parseIconSet(icons);
		if (!target.iconSet) {
			target.iconSet = parsed;
		} else {
			const set = target.iconSet;
			if (str(icons, 'iconSet')) {
				set.iconSet = parsed.iconSet;
			}
			if (parsed.cfvos.length > 0) {
				set.cfvos = parsed.cfvos;
			}
			if (parsed.customIcons) {
				set.customIcons = parsed.customIcons;
			}
			for (const key of ['showValue', 'reverse', 'percent'] as const) {
				if (parsed[key] !== undefined) {
					set[key] = parsed[key];
				}
			}
		}
	}
}

/** Merge the `x14:conditionalFormattings` payload into the 2007 formats (mutates `formats`). */
export function mergeCfExtensions(formats: ConditionalFormat[], payload: XmlElement): void {
	const byId = new Map<string, CfRule>();
	for (const f of formats) {
		for (const r of f.rules) {
			if (r.extId) {
				byId.set(guidKey(r.extId), r);
			}
		}
	}
	for (const block of kids(payload, 'conditionalFormatting')) {
		const ranges = parseSqref(textContent(kid(block, 'sqref')));
		const extOnly: CfRule[] = [];
		kids(block, 'cfRule').forEach((el, i) => {
			const id = str(el, 'id');
			const linked = id ? byId.get(guidKey(id)) : undefined;
			if (linked) {
				mergeIntoRule(linked, el);
				return;
			}
			const rule = parseRuleCommon(el, i);
			if (!rule) {
				return;
			}
			rule.fromExtension = true;
			if (id) {
				rule.extId = id;
			}
			const dxf = kid(el, 'dxf');
			if (dxf) {
				rule.dxf = parseDxf(dxf);
			}
			const bar = kid(el, 'dataBar');
			if (bar && rule.dataBar) {
				applyDataBarExtension(rule.dataBar, bar);
			}
			extOnly.push(rule);
		});
		if (extOnly.length > 0 && ranges.length > 0) {
			const format: ConditionalFormat = { ranges, rules: extOnly };
			if (bool(block, 'pivot')) {
				format.pivot = true;
			}
			formats.push(format);
		}
	}
}
