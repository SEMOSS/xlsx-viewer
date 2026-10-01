/** Sparkline groups (`x14:sparklineGroups` worksheet extension). */
import { a1 } from '../address/cell-address';
import type { SparklineGroup } from '../types/sheet';
import type { ColorSpec } from '../types/style';
import { el, textEl } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { fileFormula } from './formulas';
import { colorAttrs, flag } from './xml-utils';

export const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';

/** Colors Excel assigns to a new sparkline group (Office theme). */
const DEFAULT_COLORS: Record<string, ColorSpec> = {
	colorSeries: { theme: 4, tint: -0.499984740745262 },
	colorNegative: { theme: 5 },
	colorAxis: { rgb: 'FF000000' },
	colorMarkers: { theme: 4, tint: -0.499984740745262 },
	colorFirst: { theme: 4, tint: 0.39997558519241921 },
	colorLast: { theme: 4, tint: 0.39997558519241921 },
	colorHigh: { theme: 4 },
	colorLow: { theme: 4 },
};

function group(g: SparklineGroup): XmlElement | null {
	const lines = g.sparklines.filter((s) => s.formula);
	if (!lines.length) {
		return null;
	}
	const colors: Array<[string, ColorSpec | undefined]> = [
		['colorSeries', g.seriesColor],
		['colorNegative', g.negativeColor],
		['colorAxis', undefined],
		['colorMarkers', g.markersColor],
		['colorFirst', g.firstColor],
		['colorLast', g.lastColor],
		['colorHigh', g.highColor],
		['colorLow', g.lowColor],
	];
	return el(
		'x14:sparklineGroup',
		{
			manualMax: g.maxAxisType === 'custom' ? g.manualMax : undefined,
			manualMin: g.minAxisType === 'custom' ? g.manualMin : undefined,
			lineWeight: g.lineWeight,
			type: g.type !== 'line' ? g.type : undefined,
			displayEmptyCellsAs: g.displayEmptyCellsAs,
			markers: flag(g.markers),
			high: flag(g.high),
			low: flag(g.low),
			first: flag(g.first),
			last: flag(g.last),
			negative: flag(g.negative),
			minAxisType: g.minAxisType && g.minAxisType !== 'individual' ? g.minAxisType : undefined,
			maxAxisType: g.maxAxisType && g.maxAxisType !== 'individual' ? g.maxAxisType : undefined,
			rightToLeft: flag(g.rightToLeft),
		},
		[
			...colors.map(([name, c]) => el(`x14:${name}`, colorAttrs(c ?? DEFAULT_COLORS[name]))),
			el(
				'x14:sparklines',
				{},
				lines.map((s) =>
					el('x14:sparkline', {}, [
						textEl('xm:f', fileFormula(s.formula)),
						textEl('xm:sqref', a1(s.location.row, s.location.col)),
					]),
				),
			),
		],
	);
}

/** The sparkline `<ext>`, or `null` when the sheet has none. */
export function buildSparklineExt(groups: readonly SparklineGroup[]): XmlElement | null {
	const items = groups.map(group).filter((g): g is XmlElement => g !== null);
	if (!items.length) {
		return null;
	}
	return el('ext', { uri: SPARKLINE_EXT_URI, 'xmlns:x14': NS.x14 }, [
		el('x14:sparklineGroups', { 'xmlns:xm': NS.xm }, items),
	]);
}
