import { parseCellAddress } from '../address/cell-address';
import type { Sparkline, SparklineGroup } from '../types/sheet';
import { textContent } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { parseColor } from './style-parts';
import { bool, kid, kids, num, oneOf, str } from './xml-utils';

const COLORS = [
	['colorSeries', 'seriesColor'],
	['colorNegative', 'negativeColor'],
	['colorMarkers', 'markersColor'],
	['colorHigh', 'highColor'],
	['colorLow', 'lowColor'],
	['colorFirst', 'firstColor'],
	['colorLast', 'lastColor'],
] as const;

const FLAGS = ['markers', 'high', 'low', 'first', 'last', 'negative', 'rightToLeft'] as const;

const AXIS_TYPES = ['individual', 'group', 'custom'] as const;

function parseGroup(el: XmlElement): SparklineGroup {
	const group: SparklineGroup = {
		type: oneOf(str(el, 'type'), ['line', 'column', 'stacked'] as const) ?? 'line',
		sparklines: [],
	};
	for (const [local, key] of COLORS) {
		const c = parseColor(kid(el, local));
		if (c) {
			group[key] = c;
		}
	}
	for (const key of FLAGS) {
		const v = bool(el, key);
		if (v !== undefined) {
			group[key] = v;
		}
	}
	const weight = num(el, 'lineWeight');
	if (weight !== undefined) {
		group.lineWeight = weight;
	}
	const empty = oneOf(str(el, 'displayEmptyCellsAs'), ['gap', 'zero', 'span'] as const);
	if (empty) {
		group.displayEmptyCellsAs = empty;
	}
	const minAxis = oneOf(str(el, 'minAxisType'), AXIS_TYPES);
	if (minAxis) {
		group.minAxisType = minAxis;
	}
	const maxAxis = oneOf(str(el, 'maxAxisType'), AXIS_TYPES);
	if (maxAxis) {
		group.maxAxisType = maxAxis;
	}
	const manualMin = num(el, 'manualMin');
	if (manualMin !== undefined) {
		group.manualMin = manualMin;
	}
	const manualMax = num(el, 'manualMax');
	if (manualMax !== undefined) {
		group.manualMax = manualMax;
	}
	for (const s of kids(kid(el, 'sparklines'), 'sparkline')) {
		const location = parseCellAddress(textContent(kid(s, 'sqref')));
		if (!location) {
			continue;
		}
		const sparkline: Sparkline = { formula: textContent(kid(s, 'f')).trim(), location };
		group.sparklines.push(sparkline);
	}
	return group;
}

/** `x14:sparklineGroups` extension payload of a worksheet. */
export function parseSparklineGroups(payload: XmlElement | undefined): SparklineGroup[] {
	return kids(payload, 'sparklineGroup').map(parseGroup);
}
