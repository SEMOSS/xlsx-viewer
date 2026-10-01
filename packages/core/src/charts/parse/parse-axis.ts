/** Readers for axes, titles and the legend. */
import type { ChartAxis, ChartLegend, ChartText } from '../../types/chart';
import type { XmlElement } from '../../xml/types';
import { boolOf, kid, kids, numOf, readLine, readShapeProps, valOf } from '../xml/drawingml-read';
import { readManualLayout, readRichText, readTxPr } from '../xml/text-read';
import { readSeriesText } from './parse-series';

/** Read a `c:title` (chart or axis title). A title without `c:tx` is an automatic title. */
export function readTitle(el: XmlElement | undefined): ChartText | undefined {
	if (!el) {
		return undefined;
	}
	const out: ChartText = {};
	const tx = kid(el, 'tx');
	const rich = kid(tx, 'rich');
	if (rich) {
		const r = readRichText(rich);
		out.text = r.text;
		out.paragraphs = r.paragraphs;
		if (r.style) {
			out.style = r.style;
		}
		if (r.rotation !== undefined) {
			out.rotation = r.rotation;
		}
	} else if (tx) {
		Object.assign(out, readSeriesText(tx));
	}
	const txPr = readTxPr(kid(el, 'txPr'));
	if (txPr.style && !out.style) {
		out.style = txPr.style;
	}
	if (txPr.rotation !== undefined && out.rotation === undefined) {
		out.rotation = txPr.rotation;
	}
	const overlay = boolOf(el, 'overlay');
	if (overlay !== undefined) {
		out.overlay = overlay;
	}
	const layout = readManualLayout(kid(el, 'layout'));
	if (layout) {
		out.layout = layout;
	}
	return out;
}

const AXIS_KINDS: Record<string, ChartAxis['kind']> = {
	catAx: 'cat',
	valAx: 'val',
	dateAx: 'date',
	serAx: 'ser',
};

const TICK_MARKS = new Set(['cross', 'in', 'none', 'out']);
const TICK_LABEL_POSITIONS = new Set(['high', 'low', 'nextTo', 'none']);
const TIME_UNITS = new Set(['days', 'months', 'years']);

export function isAxisElement(local: string): boolean {
	return local in AXIS_KINDS;
}

function gridlines(el: XmlElement | undefined): ChartAxis['majorGridlines'] {
	if (!el) {
		return undefined;
	}
	return readLine(kid(kid(el, 'spPr'), 'ln')) ?? true;
}

/** Read a `c:catAx`, `c:valAx`, `c:dateAx` or `c:serAx`. */
export function readAxis(el: XmlElement, local: string): ChartAxis {
	const kind = AXIS_KINDS[local] ?? 'val';
	const pos = valOf(el, 'axPos');
	const axis: ChartAxis = {
		id: numOf(el, 'axId') ?? 0,
		kind,
		position:
			pos === 'b' || pos === 'l' || pos === 'r' || pos === 't' ? pos : kind === 'val' ? 'l' : 'b',
		crossAxisId: numOf(el, 'crossAx') ?? 0,
	};
	const deleted = boolOf(el, 'delete');
	if (deleted !== undefined) {
		axis.deleted = deleted;
	}
	const title = readTitle(kid(el, 'title'));
	if (title) {
		axis.title = title;
	}
	const numFmt = kid(el, 'numFmt');
	if (numFmt?.attrs.formatCode !== undefined) {
		axis.numFmt = { code: numFmt.attrs.formatCode };
		if (numFmt.attrs.sourceLinked !== undefined) {
			axis.numFmt.sourceLinked =
				numFmt.attrs.sourceLinked === '1' || numFmt.attrs.sourceLinked === 'true';
		}
	}
	const scaling = kid(el, 'scaling');
	const min = numOf(scaling, 'min');
	if (min !== undefined) {
		axis.min = min;
	}
	const max = numOf(scaling, 'max');
	if (max !== undefined) {
		axis.max = max;
	}
	const logBase = numOf(scaling, 'logBase');
	if (logBase !== undefined) {
		axis.logBase = logBase;
	}
	const orientation = valOf(scaling, 'orientation');
	if (orientation === 'minMax' || orientation === 'maxMin') {
		axis.orientation = orientation;
	}
	const majorUnit = numOf(el, 'majorUnit');
	if (majorUnit !== undefined && majorUnit > 0) {
		axis.majorUnit = majorUnit;
	}
	const minorUnit = numOf(el, 'minorUnit');
	if (minorUnit !== undefined && minorUnit > 0) {
		axis.minorUnit = minorUnit;
	}
	const crosses = valOf(el, 'crosses');
	if (crosses === 'autoZero' || crosses === 'min' || crosses === 'max') {
		axis.crosses = crosses;
	}
	const crossesAt = numOf(el, 'crossesAt');
	if (crossesAt !== undefined) {
		axis.crossesAt = crossesAt;
	}
	const major = gridlines(kid(el, 'majorGridlines'));
	if (major) {
		axis.majorGridlines = major;
	}
	const minor = gridlines(kid(el, 'minorGridlines'));
	if (minor) {
		axis.minorGridlines = minor;
	}
	for (const [tag, key] of [
		['majorTickMark', 'majorTickMark'],
		['minorTickMark', 'minorTickMark'],
	] as const) {
		const v = valOf(el, tag);
		if (v && TICK_MARKS.has(v)) {
			axis[key] = v as ChartAxis['majorTickMark'];
		}
	}
	const lblPos = valOf(el, 'tickLblPos');
	if (lblPos && TICK_LABEL_POSITIONS.has(lblPos)) {
		axis.tickLabelPosition = lblPos as ChartAxis['tickLabelPosition'];
	}
	const { line } = readShapeProps(kid(el, 'spPr'));
	if (line) {
		axis.line = line;
	}
	const txPr = readTxPr(kid(el, 'txPr'));
	if (txPr.style) {
		axis.textStyle = txPr.style;
	}
	if (txPr.rotation !== undefined) {
		axis.labelRotation = txPr.rotation;
	}
	const crossBetween = valOf(el, 'crossBetween');
	if (crossBetween === 'between' || crossBetween === 'midCat') {
		axis.crossBetween = crossBetween;
	}
	const skip = numOf(el, 'tickLblSkip');
	if (skip !== undefined && skip >= 1) {
		axis.tickLabelSkip = skip;
	}
	const noMulti = boolOf(el, 'noMultiLvlLbl');
	if (noMulti) {
		axis.noMultiLevelLabels = true;
	}
	for (const key of ['baseTimeUnit', 'majorTimeUnit'] as const) {
		const v = valOf(el, key);
		if (v && TIME_UNITS.has(v)) {
			axis[key] = v as ChartAxis['baseTimeUnit'];
		}
	}
	const dispUnits = kid(el, 'dispUnits');
	if (dispUnits) {
		const builtIn = valOf(dispUnits, 'builtInUnit');
		const custom = numOf(dispUnits, 'custUnit');
		const unit = builtIn ?? custom;
		if (unit !== undefined) {
			axis.displayUnits = { unit };
			if (kid(dispUnits, 'dispUnitsLbl')) {
				axis.displayUnits.showLabel = true;
			}
		}
	}
	return axis;
}

const LEGEND_POSITIONS = new Set(['r', 'l', 't', 'b', 'tr']);

export function readLegend(el: XmlElement | undefined): ChartLegend | undefined {
	if (!el) {
		return undefined;
	}
	const pos = valOf(el, 'legendPos');
	const legend: ChartLegend = {
		position: pos && LEGEND_POSITIONS.has(pos) ? (pos as ChartLegend['position']) : 'r',
	};
	const overlay = boolOf(el, 'overlay');
	if (overlay !== undefined) {
		legend.overlay = overlay;
	}
	const deleted: number[] = [];
	for (const entry of kids(el, 'legendEntry')) {
		const idx = numOf(entry, 'idx');
		if (idx !== undefined && boolOf(entry, 'delete')) {
			deleted.push(idx);
		}
	}
	if (deleted.length) {
		legend.deletedEntries = deleted;
	}
	const { style } = readTxPr(kid(el, 'txPr'));
	if (style) {
		legend.textStyle = style;
	}
	Object.assign(legend, readShapeProps(kid(el, 'spPr')));
	const layout = readManualLayout(kid(el, 'layout'));
	if (layout) {
		legend.layout = layout;
	}
	return legend;
}
