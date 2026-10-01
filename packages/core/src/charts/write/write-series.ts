/** Builders for series content: data sources, markers, data points, labels and trendlines. */
import type {
	ChartDataLabels,
	ChartDataPoint,
	ChartDataRef,
	ChartMarker,
	ChartSeries,
	ChartText,
	ChartTrendline,
} from '../../types/chart';
import { el, textEl } from '../../xml/dom';
import type { XmlElement } from '../../xml/types';
import {
	formatXmlNumber,
	valEl,
	writeLayout,
	writeShapeProps,
	writeTextProps,
} from '../xml/drawingml-write';
import { mergeRawChildren, rawChild, rawChildByKey } from './merge-raw';
import { CHART_ORDER, seriesOrderKey } from './schema';

type CacheValue = string | number | null;

function points(cache: readonly CacheValue[], numeric: boolean): XmlElement[] {
	const out: XmlElement[] = [];
	cache.forEach((v, idx) => {
		if (v === null || v === undefined) {
			return;
		}
		if (numeric) {
			const n = typeof v === 'number' ? v : Number(v);
			if (!Number.isFinite(n)) {
				return;
			}
			out.push(el('c:pt', { idx }, [textEl('c:v', formatXmlNumber(n))]));
		} else {
			out.push(el('c:pt', { idx }, [textEl('c:v', String(v))]));
		}
	});
	return out;
}

function numData(name: string, ref: ChartDataRef): XmlElement {
	const cache = ref.cache ?? [];
	return el(name, {}, [
		textEl('c:formatCode', ref.formatCode ?? 'General'),
		valEl('c:ptCount', cache.length),
		...points(cache, true),
	]);
}

function strData(name: string, cache: readonly CacheValue[]): XmlElement {
	return el(name, {}, [valEl('c:ptCount', cache.length), ...points(cache, false)]);
}

function isNumericCache(cache: readonly CacheValue[] | undefined): boolean {
	return (cache ?? []).every((v) => v === null || typeof v === 'number');
}

/**
 * A data source container (`c:cat`, `c:val`, `c:xVal`, `c:yVal`, `c:bubbleSize`).
 * `numericOnly` containers (`val`, `yVal`, `bubbleSize`) only allow numbers.
 */
export function writeDataRef(
	name: string,
	ref: ChartDataRef | undefined,
	numericOnly: boolean,
): XmlElement | undefined {
	if (!ref) {
		return undefined;
	}
	const cache = ref.cache ?? [];
	const formula = ref.ref ? textEl('c:f', ref.ref) : undefined;
	let body: XmlElement;
	if (
		numericOnly ||
		ref.kind === 'num' ||
		(ref.kind === 'literal' && isNumericCache(cache) && ref.formatCode)
	) {
		body = formula
			? el('c:numRef', {}, [formula, numData('c:numCache', ref)])
			: numData('c:numLit', ref);
	} else if (ref.kind === 'multiLvlStr' && formula) {
		const levels = ref.levels?.length
			? ref.levels
			: [cache.map((v) => (v === null ? null : String(v)))];
		const count = Math.max(0, ...levels.map((l) => l.length));
		const lvls = [...levels].reverse().map((l) => el('c:lvl', {}, points(l, false)));
		body = el('c:multiLvlStrRef', {}, [
			formula,
			el('c:multiLvlStrCache', {}, [valEl('c:ptCount', count), ...lvls]),
		]);
	} else if (formula && ref.kind !== 'literal') {
		body = el('c:strRef', {}, [formula, strData('c:strCache', cache)]);
	} else {
		body = strData('c:strLit', cache);
	}
	return el(name, {}, [body]);
}

/** Series name (`c:tx`). */
export function writeSeriesText(text: ChartText | undefined): XmlElement | undefined {
	if (!text || (text.ref === undefined && text.text === undefined)) {
		return undefined;
	}
	if (text.ref) {
		const cache = text.text === undefined ? [] : [text.text];
		return el('c:tx', {}, [
			el('c:strRef', {}, [textEl('c:f', text.ref), strData('c:strCache', cache)]),
		]);
	}
	return el('c:tx', {}, [textEl('c:v', text.text ?? '')]);
}

export function writeMarker(
	marker: ChartMarker | undefined,
	raw?: XmlElement,
): XmlElement | undefined {
	if (!marker) {
		return undefined;
	}
	const out = el('c:marker', {}, [
		marker.symbol ? valEl('c:symbol', marker.symbol) : undefined,
		marker.size !== undefined
			? valEl('c:size', Math.min(72, Math.max(2, Math.round(marker.size))))
			: undefined,
		writeShapeProps(marker.fill, marker.line),
	]);
	return mergeRawChildren(out, raw, { order: 'marker', owned: ['symbol', 'size', 'spPr'] });
}

const SHOW_FLAGS: Array<[string, keyof ChartDataLabels]> = [
	['c:showLegendKey', 'showLegendKey'],
	['c:showVal', 'showValue'],
	['c:showCatName', 'showCategoryName'],
	['c:showSerName', 'showSeriesName'],
	['c:showPercent', 'showPercent'],
	['c:showBubbleSize', 'showBubbleSize'],
];

function labelGroup(
	labels: ChartDataLabels,
	inherit?: ChartDataLabels,
): Array<XmlElement | undefined> {
	return [
		labels.numFmt !== undefined
			? el('c:numFmt', { formatCode: labels.numFmt, sourceLinked: 0 })
			: undefined,
		writeShapeProps(labels.fill, labels.line),
		writeTextProps(labels.style),
		labels.position ? valEl('c:dLblPos', labels.position) : undefined,
		// Excel treats missing show* flags inconsistently, so all six are always written.
		...SHOW_FLAGS.map(([tag, key]) => valEl(tag, Boolean(labels[key] ?? inherit?.[key] ?? false))),
		labels.separator !== undefined ? textEl('c:separator', labels.separator) : undefined,
	];
}

const LABEL_OWNED = [
	'dLbl',
	'delete',
	'numFmt',
	'spPr',
	'txPr',
	'dLblPos',
	'showLegendKey',
	'showVal',
	'showCatName',
	'showSerName',
	'showPercent',
	'showBubbleSize',
	'separator',
	'showLeaderLines',
	'layout',
	'idx',
];

/** `c:dLbls` for a series or a plot. */
export function writeDataLabels(
	labels: ChartDataLabels | undefined,
	raw?: XmlElement,
): XmlElement | undefined {
	if (!labels) {
		return undefined;
	}
	const pointEls = (labels.points ?? []).map((p) => {
		const body = p.deleted
			? [valEl('c:delete', true)]
			: [p.layout ? writeLayout(p.layout) : undefined, ...labelGroup(p, labels)];
		const out = el('c:dLbl', {}, [valEl('c:idx', p.idx), ...body]);
		const rawPoint = rawChildByKey(raw, 'dLbl', 'idx', p.idx);
		return p.deleted ? out : mergeRawChildren(out, rawPoint, { order: 'dLbl', owned: LABEL_OWNED });
	});
	if (labels.deleted) {
		return el('c:dLbls', {}, [...pointEls, valEl('c:delete', true)]);
	}
	const out = el('c:dLbls', {}, [
		...pointEls,
		...labelGroup(labels),
		labels.showLeaderLines !== undefined
			? valEl('c:showLeaderLines', labels.showLeaderLines)
			: undefined,
	]);
	return mergeRawChildren(out, raw, { order: 'dLbls', owned: LABEL_OWNED });
}

function writeTrendline(t: ChartTrendline): XmlElement {
	return el('c:trendline', {}, [
		t.name !== undefined ? textEl('c:name', t.name) : undefined,
		writeShapeProps(undefined, t.line),
		valEl('c:trendlineType', t.type),
		t.type === 'poly' ? valEl('c:order', Math.min(6, Math.max(2, t.order ?? 2))) : undefined,
		t.type === 'movingAvg' ? valEl('c:period', Math.max(2, t.period ?? 2)) : undefined,
		valEl('c:dispRSqr', false),
		valEl('c:dispEq', false),
	]);
}

function writePoint(p: ChartDataPoint, container: string, raw?: XmlElement): XmlElement {
	const isPie = seriesOrderKey(container) === 'pieSer';
	const isBar = seriesOrderKey(container) === 'barSer' || container === 'bubbleChart';
	const hasMarker = ['lineSer', 'scatterSer', 'radarSer'].includes(seriesOrderKey(container));
	const out = el('c:dPt', {}, [
		valEl('c:idx', p.idx),
		isBar && p.invertIfNegative !== undefined
			? valEl('c:invertIfNegative', p.invertIfNegative)
			: undefined,
		hasMarker ? writeMarker(p.marker, rawChild(raw, 'marker')) : undefined,
		isPie || isBar ? valEl('c:bubble3D', false) : undefined,
		isPie && p.explosion !== undefined ? valEl('c:explosion', Math.round(p.explosion)) : undefined,
		writeShapeProps(p.fill, p.line),
	]);
	return mergeRawChildren(out, raw, {
		order: 'dPt',
		owned: ['idx', 'invertIfNegative', 'marker', 'explosion', 'spPr'],
		preferRaw: ['bubble3D'],
	});
}

const ORDER_HAS: Record<string, Set<string>> = Object.fromEntries(
	[
		'barSer',
		'lineSer',
		'pieSer',
		'areaSer',
		'scatterSer',
		'radarSer',
		'bubbleSer',
		'surfaceSer',
	].map((k) => [k, new Set(CHART_ORDER[k])]),
);

/** One `c:ser` for a plot container (`barChart`, `lineChart`, ...). */
export function writeSeries(series: ChartSeries, container: string, raw?: XmlElement): XmlElement {
	const key = seriesOrderKey(container);
	const xy = key === 'scatterSer' || key === 'bubbleSer';
	const has = (local: string): boolean => ORDER_HAS[key]?.has(local) ?? false;
	const out = el('c:ser', {}, [
		valEl('c:idx', series.idx),
		valEl('c:order', series.order),
		writeSeriesText(series.name),
		writeShapeProps(series.fill, series.line),
		has('invertIfNegative') && series.invertIfNegative !== undefined
			? valEl('c:invertIfNegative', series.invertIfNegative)
			: undefined,
		has('marker') ? writeMarker(series.marker, rawChild(raw, 'marker')) : undefined,
		has('explosion') && series.explosion !== undefined
			? valEl('c:explosion', Math.round(series.explosion))
			: undefined,
		...(has('dPt')
			? (series.points ?? []).map((p) =>
					writePoint(p, container, rawChildByKey(raw, 'dPt', 'idx', p.idx)),
				)
			: []),
		has('dLbls') ? writeDataLabels(series.dataLabels, rawChild(raw, 'dLbls')) : undefined,
		...(has('trendline') ? (series.trendlines ?? []).map(writeTrendline) : []),
		writeDataRef(xy ? 'c:xVal' : 'c:cat', series.categories, false),
		writeDataRef(xy ? 'c:yVal' : 'c:val', series.values, true),
		key === 'bubbleSer' ? writeDataRef('c:bubbleSize', series.bubbleSizes, true) : undefined,
		has('smooth') && series.smooth !== undefined ? valEl('c:smooth', series.smooth) : undefined,
	]);
	return mergeRawChildren(out, raw, {
		order: key,
		owned: [
			'idx',
			'order',
			'tx',
			'spPr',
			'invertIfNegative',
			'marker',
			'explosion',
			'dPt',
			'dLbls',
			'trendline',
			'cat',
			'val',
			'xVal',
			'yVal',
			'bubbleSize',
			'smooth',
		],
	});
}
