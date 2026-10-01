/**
 * Series-level readers: data sources (`c:cat`, `c:val`, ...), markers, data
 * points, data labels and trendlines.
 */
import type {
	ChartDataLabels,
	ChartDataPoint,
	ChartDataRef,
	ChartMarker,
	ChartSeries,
	ChartText,
	ChartTrendline,
} from '../../types/chart';
import { textContent } from '../../xml/dom';
import type { XmlElement } from '../../xml/types';
import { boolOf, kid, kids, numAttr, numOf, readShapeProps, valOf } from '../xml/drawingml-read';
import { readManualLayout, readTxPr } from '../xml/text-read';

/** Cached points of a `c:numCache` / `c:strCache` / literal / `c:lvl`, indexed by `idx`. */
function readPoints(data: XmlElement | undefined, numeric: boolean): Array<string | number | null> {
	if (!data) {
		return [];
	}
	const pts = kids(data, 'pt');
	let count = numOf(data, 'ptCount') ?? 0;
	for (const pt of pts) {
		count = Math.max(count, (numAttr(pt, 'idx') ?? 0) + 1);
	}
	const out: Array<string | number | null> = new Array(count).fill(null);
	pts.forEach((pt, i) => {
		const idx = numAttr(pt, 'idx') ?? i;
		const raw = textContent(kid(pt, 'v'));
		if (numeric) {
			const n = raw.trim() === '' ? NaN : Number(raw);
			out[idx] = Number.isFinite(n) ? n : null;
		} else {
			out[idx] = raw;
		}
	});
	return out;
}

/** Formula of a reference; filtered references keep it in `c15:formulaRef/c15:sqref`. */
function readFormula(refEl: XmlElement): string | undefined {
	const f = kid(refEl, 'f');
	if (f) {
		return textContent(f).trim() || undefined;
	}
	for (const ext of kids(kid(refEl, 'extLst'), 'ext')) {
		const sqref = kid(kid(ext, 'formulaRef'), 'sqref');
		if (sqref) {
			return textContent(sqref).trim() || undefined;
		}
	}
	return undefined;
}

/** Read a data source container (`c:cat`, `c:val`, `c:xVal`, `c:yVal`, `c:bubbleSize`). */
export function readDataRef(container: XmlElement | undefined): ChartDataRef | undefined {
	if (!container) {
		return undefined;
	}
	const numRef = kid(container, 'numRef');
	if (numRef) {
		const cacheEl = kid(numRef, 'numCache');
		const out: ChartDataRef = { kind: 'num', cache: readPoints(cacheEl, true) };
		const ref = readFormula(numRef);
		if (ref) {
			out.ref = ref;
		}
		const fmt = cacheEl ? textContent(kid(cacheEl, 'formatCode')) : '';
		if (fmt) {
			out.formatCode = fmt;
		}
		return out;
	}
	const strRef = kid(container, 'strRef');
	if (strRef) {
		const out: ChartDataRef = { kind: 'str', cache: readPoints(kid(strRef, 'strCache'), false) };
		const ref = readFormula(strRef);
		if (ref) {
			out.ref = ref;
		}
		return out;
	}
	const multi = kid(container, 'multiLvlStrRef');
	if (multi) {
		const cacheEl = kid(multi, 'multiLvlStrCache');
		// File order is innermost level first; the model stores outer levels first.
		const fileLevels = kids(cacheEl, 'lvl').map(
			(lvl) => readPoints(lvl, false) as Array<string | null>,
		);
		const count = numOf(cacheEl, 'ptCount') ?? fileLevels[0]?.length ?? 0;
		const levels = fileLevels.map((l) => padTo(l, count)).reverse();
		const out: ChartDataRef = {
			kind: 'multiLvlStr',
			cache: levels.length ? [...levels[levels.length - 1]] : [],
			levels,
		};
		const ref = readFormula(multi);
		if (ref) {
			out.ref = ref;
		}
		return out;
	}
	const numLit = kid(container, 'numLit');
	if (numLit) {
		const out: ChartDataRef = { kind: 'literal', cache: readPoints(numLit, true) };
		const fmt = textContent(kid(numLit, 'formatCode'));
		out.formatCode = fmt || 'General';
		return out;
	}
	const strLit = kid(container, 'strLit');
	if (strLit) {
		return { kind: 'literal', cache: readPoints(strLit, false) };
	}
	return undefined;
}

function padTo<T>(arr: Array<T | null>, count: number): Array<T | null> {
	const out = arr.slice(0, Math.max(count, arr.length));
	while (out.length < count) {
		out.push(null);
	}
	return out;
}

/** Series name (`c:tx`): a cell reference with cached text, or a literal. */
export function readSeriesText(tx: XmlElement | undefined): ChartText | undefined {
	if (!tx) {
		return undefined;
	}
	const strRef = kid(tx, 'strRef');
	if (strRef) {
		const out: ChartText = {};
		const ref = readFormula(strRef);
		if (ref) {
			out.ref = ref;
		}
		const cache = readPoints(kid(strRef, 'strCache'), false).filter((v) => v !== null);
		if (cache.length) {
			out.text = cache.join(' ');
		}
		return out;
	}
	const v = kid(tx, 'v');
	return v ? { text: textContent(v) } : undefined;
}

const MARKER_SYMBOLS = new Set([
	'none',
	'auto',
	'circle',
	'dash',
	'diamond',
	'dot',
	'picture',
	'plus',
	'square',
	'star',
	'triangle',
	'x',
]);

export function readMarker(el: XmlElement | undefined): ChartMarker | undefined {
	if (!el) {
		return undefined;
	}
	const marker: ChartMarker = {};
	const symbol = valOf(el, 'symbol');
	if (symbol && MARKER_SYMBOLS.has(symbol)) {
		marker.symbol = symbol as ChartMarker['symbol'];
	}
	const size = numOf(el, 'size');
	if (size !== undefined) {
		marker.size = size;
	}
	Object.assign(marker, readShapeProps(kid(el, 'spPr')));
	return marker;
}

const LABEL_POSITIONS = new Set([
	'bestFit',
	'b',
	'ctr',
	'inBase',
	'inEnd',
	'l',
	'outEnd',
	'r',
	't',
]);

/** The label settings shared by `c:dLbls` and `c:dLbl`. */
function readLabelFields(el: XmlElement, out: ChartDataLabels): void {
	if (boolOf(el, 'delete')) {
		out.deleted = true;
	}
	const numFmt = kid(el, 'numFmt');
	if (numFmt?.attrs.formatCode !== undefined && numFmt.attrs.sourceLinked !== '1') {
		out.numFmt = numFmt.attrs.formatCode;
	}
	const sp = readShapeProps(kid(el, 'spPr'));
	if (sp.fill) {
		out.fill = sp.fill;
	}
	if (sp.line) {
		out.line = sp.line;
	}
	const { style } = readTxPr(kid(el, 'txPr'));
	if (style) {
		out.style = style;
	}
	const pos = valOf(el, 'dLblPos');
	if (pos && LABEL_POSITIONS.has(pos)) {
		out.position = pos as ChartDataLabels['position'];
	}
	const flags: Array<[string, keyof ChartDataLabels]> = [
		['showLegendKey', 'showLegendKey'],
		['showVal', 'showValue'],
		['showCatName', 'showCategoryName'],
		['showSerName', 'showSeriesName'],
		['showPercent', 'showPercent'],
		['showBubbleSize', 'showBubbleSize'],
		['showLeaderLines', 'showLeaderLines'],
	];
	for (const [tag, key] of flags) {
		const v = boolOf(el, tag);
		if (v !== undefined) {
			(out as Record<string, unknown>)[key] = v;
		}
	}
	const sep = kid(el, 'separator');
	if (sep) {
		out.separator = textContent(sep);
	}
}

export function readDataLabels(el: XmlElement | undefined): ChartDataLabels | undefined {
	if (!el) {
		return undefined;
	}
	const out: ChartDataLabels = {};
	readLabelFields(el, out);
	const points: NonNullable<ChartDataLabels['points']> = [];
	for (const d of kids(el, 'dLbl')) {
		const point: ChartDataLabels & { idx: number; layout?: ReturnType<typeof readManualLayout> } = {
			idx: numOf(d, 'idx') ?? 0,
		};
		readLabelFields(d, point);
		const layout = readManualLayout(kid(d, 'layout'));
		if (layout) {
			point.layout = layout;
		}
		points.push(point as ChartDataLabels & { idx: number });
	}
	if (points.length) {
		out.points = points;
	}
	return out;
}

const TRENDLINE_TYPES = new Set(['linear', 'exp', 'log', 'movingAvg', 'poly', 'power']);

function readTrendline(el: XmlElement): ChartTrendline | undefined {
	const type = valOf(el, 'trendlineType');
	if (!type || !TRENDLINE_TYPES.has(type)) {
		return undefined;
	}
	const out: ChartTrendline = { type: type as ChartTrendline['type'] };
	const order = numOf(el, 'order');
	if (order !== undefined) {
		out.order = order;
	}
	const period = numOf(el, 'period');
	if (period !== undefined) {
		out.period = period;
	}
	const { line } = readShapeProps(kid(el, 'spPr'));
	if (line) {
		out.line = line;
	}
	const name = kid(el, 'name');
	if (name) {
		out.name = textContent(name);
	}
	return out;
}

function readPoint(el: XmlElement): ChartDataPoint {
	const point: ChartDataPoint = { idx: numOf(el, 'idx') ?? 0 };
	const inv = boolOf(el, 'invertIfNegative');
	if (inv !== undefined) {
		point.invertIfNegative = inv;
	}
	const marker = readMarker(kid(el, 'marker'));
	if (marker) {
		point.marker = marker;
	}
	const explosion = numOf(el, 'explosion');
	if (explosion !== undefined) {
		point.explosion = explosion;
	}
	Object.assign(point, readShapeProps(kid(el, 'spPr')));
	return point;
}

/** Read one `c:ser` element. `position` is its index within the plot. */
export function readSeries(el: XmlElement, position: number): ChartSeries {
	const idx = numOf(el, 'idx') ?? position;
	const series: ChartSeries = { idx, order: numOf(el, 'order') ?? idx };
	const name = readSeriesText(kid(el, 'tx'));
	if (name) {
		series.name = name;
	}
	Object.assign(series, readShapeProps(kid(el, 'spPr')));
	const marker = readMarker(kid(el, 'marker'));
	if (marker) {
		series.marker = marker;
	}
	const points = kids(el, 'dPt').map(readPoint);
	if (points.length) {
		series.points = points;
	}
	const labels = readDataLabels(kid(el, 'dLbls'));
	if (labels) {
		series.dataLabels = labels;
	}
	const trendlines = kids(el, 'trendline')
		.map(readTrendline)
		.filter((t): t is ChartTrendline => t !== undefined);
	if (trendlines.length) {
		series.trendlines = trendlines;
	}
	const categories = readDataRef(kid(el, 'cat') ?? kid(el, 'xVal'));
	if (categories) {
		series.categories = categories;
	}
	const values = readDataRef(kid(el, 'val') ?? kid(el, 'yVal'));
	if (values) {
		series.values = values;
	}
	const sizes = readDataRef(kid(el, 'bubbleSize'));
	if (sizes) {
		series.bubbleSizes = sizes;
	}
	const smooth = boolOf(el, 'smooth');
	if (smooth !== undefined) {
		series.smooth = smooth;
	}
	const explosion = numOf(el, 'explosion');
	if (explosion !== undefined) {
		series.explosion = explosion;
	}
	const inv = boolOf(el, 'invertIfNegative');
	if (inv !== undefined) {
		series.invertIfNegative = inv;
	}
	return series;
}
