import { parseHex, type Rgb } from '../color/color-math';
import { resolveColorHex } from '../color/resolve-color';
import type { CellValue } from '../types/cell';
import type { Cfvo } from '../types/conditional-format';
import { INVALID_RULE, type CompiledRule, type RuleEnv, type RuleState } from './compiled';
import { interpolateColorScale, rgbToCss } from './color-scale';
import { buildDataBar, prepareDataBarStyle } from './data-bar';
import { formulaResultToScalar, scalarToNumber } from './formula-eval';
import { defaultIconCfvos, iconForThreshold, iconSetSize, selectIconThreshold } from './icon-set';
import { resolveCfvo, type ThresholdEvaluator } from './thresholds';
import type { DataBarResult, IconResult } from './types';

const WHITE: Rgb = { r: 1, g: 1, b: 1 };

/** Threshold formulas are evaluated once, at the anchor cell of the rule's block. */
function thresholdEvaluator(cr: CompiledRule, env: RuleEnv): ThresholdEvaluator {
	return (formula) => {
		let result;
		try {
			result = env.evaluate(formula, env.sheetIndex, cr.block.anchorRow, cr.block.anchorCol);
		} catch {
			return undefined;
		}
		return scalarToNumber(formulaResultToScalar(result, env.read), env.workbook.date1904);
	};
}

function resolveThresholds(
	cr: CompiledRule,
	env: RuleEnv,
	cfvos: readonly Cfvo[],
): number[] | undefined {
	const stats = cr.block.statistics;
	const evaluate = thresholdEvaluator(cr, env);
	const out: number[] = [];
	for (const cfvo of cfvos) {
		const t = resolveCfvo(cfvo, stats, evaluate);
		if (t === undefined) {
			return undefined;
		}
		out.push(t);
	}
	return out;
}

function isNumber(value: CellValue): value is number {
	return typeof value === 'number' && Number.isFinite(value);
}

function prepareColorScale(cr: CompiledRule, env: RuleEnv): RuleState {
	const scale = cr.rule.colorScale;
	if (!scale) {
		return INVALID_RULE;
	}
	const n = Math.min(scale.cfvos.length, scale.colors.length, 3);
	if (n < 2) {
		return INVALID_RULE;
	}
	const thresholds = resolveThresholds(cr, env, scale.cfvos.slice(0, n));
	if (!thresholds) {
		return INVALID_RULE;
	}
	const colors = scale.colors
		.slice(0, n)
		.map((c) => parseHex(resolveColorHex(c, env.colorCtx) ?? 'FFFFFF') ?? WHITE);
	return { kind: 'colorScale', thresholds, colors };
}

/** Color-scale fill (CSS) for a numeric cell, `undefined` for other values or an invalid rule. */
export function colorScaleColor(
	cr: CompiledRule,
	value: CellValue,
	env: RuleEnv,
): string | undefined {
	if (!isNumber(value)) {
		return undefined;
	}
	cr.state ??= prepareColorScale(cr, env);
	if (cr.state.kind !== 'colorScale') {
		return undefined;
	}
	return rgbToCss(interpolateColorScale(value, cr.state.thresholds, cr.state.colors));
}

function prepareDataBar(cr: CompiledRule, env: RuleEnv): RuleState {
	const db = cr.rule.dataBar;
	if (!db || !db.cfvos || db.cfvos.length < 2) {
		return INVALID_RULE;
	}
	const thresholds = resolveThresholds(cr, env, db.cfvos.slice(0, 2));
	if (!thresholds) {
		return INVALID_RULE;
	}
	const style = prepareDataBarStyle(db, env.colorCtx, env.sheet.view?.rightToLeft === true);
	return { kind: 'dataBar', lo: thresholds[0], hi: thresholds[1], style };
}

/** Data bar for a numeric cell, `undefined` for other values or an invalid rule. */
export function dataBarFor(
	cr: CompiledRule,
	value: CellValue,
	env: RuleEnv,
): DataBarResult | undefined {
	if (!isNumber(value)) {
		return undefined;
	}
	cr.state ??= prepareDataBar(cr, env);
	if (cr.state.kind !== 'dataBar') {
		return undefined;
	}
	return buildDataBar(cr.state.style, value, cr.state.lo, cr.state.hi);
}

function prepareIconSet(cr: CompiledRule, env: RuleEnv): RuleState {
	const set = cr.rule.iconSet;
	if (!set) {
		return INVALID_RULE;
	}
	const count = iconSetSize(set);
	if (count < 2) {
		return INVALID_RULE;
	}
	const cfvos = set.cfvos?.length === count ? set.cfvos : defaultIconCfvos(count);
	const stats = cr.block.statistics;
	const evaluate = thresholdEvaluator(cr, env);
	const thresholds: number[] = [];
	for (let i = 0; i < count; i++) {
		const t = resolveCfvo(cfvos[i], stats, evaluate);
		if (t === undefined && i > 0) {
			return INVALID_RULE;
		}
		thresholds.push(t ?? -Infinity);
	}
	return { kind: 'iconSet', thresholds, gte: cfvos.map((c) => c.gte !== false), count };
}

export interface IconOutcome {
	/** The icon to draw, or `undefined` for "No Cell Icon". */
	icon: IconResult | undefined;
	showValue: boolean;
}

/** Icon-set outcome for a numeric cell, `undefined` for other values or an invalid rule. */
export function iconFor(cr: CompiledRule, value: CellValue, env: RuleEnv): IconOutcome | undefined {
	const set = cr.rule.iconSet;
	if (!isNumber(value) || !set) {
		return undefined;
	}
	cr.state ??= prepareIconSet(cr, env);
	if (cr.state.kind !== 'iconSet') {
		return undefined;
	}
	const index = selectIconThreshold(value, cr.state.thresholds, cr.state.gte);
	return {
		icon: iconForThreshold(set, index, cr.state.count) ?? undefined,
		showValue: set.showValue !== false,
	};
}
