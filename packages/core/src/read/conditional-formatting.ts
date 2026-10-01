import { parseSqref } from '../address/range';
import { stripFormulaPrefixes } from '../formula/prefixes';
import type {
	CfColorScale,
	CfDataBar,
	CfIconSet,
	CfOperator,
	CfRule,
	CfRuleType,
	CfTimePeriod,
	Cfvo,
	ConditionalFormat,
	IconSetName,
} from '../types/conditional-format';
import type { ColorSpec } from '../types/style';
import { findDescendantLocal, textContent } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { parseColor } from './style-parts';
import { bool, compact, int, kid, kids, oneOf, str } from './xml-utils';

/** Excel 2007 `<conditionalFormatting>` blocks. The 2010 extensions are merged by `cf-extensions.ts`. */

/** Rule types of ST_CfType (kept private: `conditional/` owns the public icon-set and rule tables). */
const RULE_TYPES: readonly CfRuleType[] = [
	'expression',
	'cellIs',
	'colorScale',
	'dataBar',
	'iconSet',
	'top10',
	'uniqueValues',
	'duplicateValues',
	'containsText',
	'notContainsText',
	'beginsWith',
	'endsWith',
	'containsBlanks',
	'notContainsBlanks',
	'containsErrors',
	'notContainsErrors',
	'timePeriod',
	'aboveAverage',
];

const RULE_OPERATORS: readonly CfOperator[] = [
	'lessThan',
	'lessThanOrEqual',
	'equal',
	'notEqual',
	'greaterThanOrEqual',
	'greaterThan',
	'between',
	'notBetween',
	'containsText',
	'notContains',
	'beginsWith',
	'endsWith',
];

const TIME_PERIODS: readonly CfTimePeriod[] = [
	'today',
	'yesterday',
	'tomorrow',
	'last7Days',
	'thisMonth',
	'lastMonth',
	'nextMonth',
	'thisWeek',
	'lastWeek',
	'nextWeek',
];

const ICON_SET_NAMES: readonly IconSetName[] = [
	'3Arrows',
	'3ArrowsGray',
	'3Flags',
	'3TrafficLights1',
	'3TrafficLights2',
	'3Signs',
	'3Symbols',
	'3Symbols2',
	'3Stars',
	'3Triangles',
	'4Arrows',
	'4ArrowsGray',
	'4RedToBlack',
	'4Rating',
	'4TrafficLights',
	'5Arrows',
	'5ArrowsGray',
	'5Rating',
	'5Quarters',
	'5Boxes',
	'NoIcons',
];

const CFVO_TYPES: ReadonlyArray<Cfvo['type']> = [
	'num',
	'percent',
	'max',
	'min',
	'formula',
	'percentile',
	'autoMin',
	'autoMax',
];

/** A `cfvo` threshold: the value is the `val` attribute or (2010) an `xm:f` child. */
export function parseCfvo(el: XmlElement): Cfvo {
	const cfvo: Cfvo = { type: oneOf(str(el, 'type'), CFVO_TYPES) ?? 'num' };
	const f = kid(el, 'f');
	const value = f ? textContent(f) : str(el, 'val');
	if (value !== undefined) {
		cfvo.value = stripFormulaPrefixes(value);
	}
	const gte = bool(el, 'gte');
	if (gte !== undefined) {
		cfvo.gte = gte;
	}
	return cfvo;
}

export function parseColorScale(el: XmlElement): CfColorScale {
	return {
		cfvos: kids(el, 'cfvo').map(parseCfvo),
		colors: kids(el, 'color').map((c) => parseColor(c) ?? {}),
	};
}

function pairCfvos(list: Cfvo[]): [Cfvo, Cfvo] {
	return [list[0] ?? { type: 'min' }, list[1] ?? { type: 'max' }];
}

export function parseDataBar(el: XmlElement): CfDataBar {
	const color: ColorSpec = parseColor(kid(el, 'color')) ?? parseColor(kid(el, 'fillColor')) ?? {};
	return compact<CfDataBar>({
		cfvos: pairCfvos(kids(el, 'cfvo').map(parseCfvo)),
		color,
		showValue: bool(el, 'showValue'),
		minLength: int(el, 'minLength'),
		maxLength: int(el, 'maxLength'),
	});
}

export function parseIconSet(el: XmlElement): CfIconSet {
	const set: CfIconSet = compact({
		iconSet: oneOf(str(el, 'iconSet'), ICON_SET_NAMES) ?? '3TrafficLights1',
		cfvos: kids(el, 'cfvo').map(parseCfvo),
		showValue: bool(el, 'showValue'),
		reverse: bool(el, 'reverse'),
		percent: bool(el, 'percent'),
	});
	const icons = kids(el, 'cfIcon').map((i) => ({
		iconSet: oneOf(str(i, 'iconSet'), ICON_SET_NAMES) ?? 'NoIcons',
		iconId: int(i, 'iconId') ?? 0,
	}));
	if (icons.length > 0) {
		set.customIcons = icons;
	}
	return set;
}

/**
 * Attributes and children shared by 2007 `cfRule` and 2010 `x14:cfRule`.
 * Formulas are `formula` (2007) or `xm:f` (2010) children.
 */
export function parseRuleCommon(el: XmlElement, index: number): CfRule | undefined {
	const type = oneOf(str(el, 'type'), RULE_TYPES);
	if (!type) {
		return undefined;
	}
	const rule: CfRule = compact({
		type,
		priority: int(el, 'priority') ?? index + 1,
		stopIfTrue: bool(el, 'stopIfTrue'),
		dxfId: int(el, 'dxfId'),
		operator: oneOf(str(el, 'operator'), RULE_OPERATORS),
		text: str(el, 'text'),
		timePeriod: oneOf(str(el, 'timePeriod'), TIME_PERIODS),
		rank: int(el, 'rank'),
		percent: bool(el, 'percent'),
		bottom: bool(el, 'bottom'),
		equalAverage: bool(el, 'equalAverage'),
		stdDev: int(el, 'stdDev'),
	});
	if (type === 'aboveAverage') {
		rule.aboveAverage = bool(el, 'aboveAverage') ?? true;
	}
	const formulas = [...kids(el, 'formula'), ...kids(el, 'f')].map((f) =>
		stripFormulaPrefixes(textContent(f)),
	);
	if (formulas.length > 0) {
		rule.formulas = formulas;
	}
	const scale = kid(el, 'colorScale');
	if (scale) {
		rule.colorScale = parseColorScale(scale);
	}
	const bar = kid(el, 'dataBar');
	if (bar) {
		rule.dataBar = parseDataBar(bar);
	}
	const icons = kid(el, 'iconSet');
	if (icons) {
		rule.iconSet = parseIconSet(icons);
	}
	return rule;
}

/** The `x14:id` GUID linking a 2007 rule to its 2010 extension. */
function extensionId(rule: XmlElement): string | undefined {
	const ext = kid(rule, 'extLst');
	const id = ext ? findDescendantLocal(ext, 'id') : undefined;
	const text = id ? textContent(id).trim() : '';
	return text || undefined;
}

/** All 2007 `<conditionalFormatting>` children of the worksheet root. */
export function parseConditionalFormats(root: XmlElement): ConditionalFormat[] {
	const out: ConditionalFormat[] = [];
	for (const cf of kids(root, 'conditionalFormatting')) {
		const ranges = parseSqref(str(cf, 'sqref') ?? '');
		if (ranges.length === 0) {
			continue;
		}
		const rules: CfRule[] = [];
		kids(cf, 'cfRule').forEach((r, i) => {
			const rule = parseRuleCommon(r, i);
			if (!rule) {
				return;
			}
			const id = extensionId(r);
			if (id) {
				rule.extId = id;
			}
			rules.push(rule);
		});
		const format: ConditionalFormat = { ranges, rules };
		if (bool(cf, 'pivot')) {
			format.pivot = true;
		}
		out.push(format);
	}
	return out;
}
