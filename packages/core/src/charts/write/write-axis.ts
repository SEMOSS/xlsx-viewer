/** Builders for titles, axes and the legend. */
import type { ChartAxis, ChartLegend, ChartText } from '../../types/chart';
import { el, textEl } from '../../xml/dom';
import type { XmlElement } from '../../xml/types';
import {
	formatXmlNumber,
	valEl,
	writeLayout,
	writeLine,
	writeRichText,
	writeShapeProps,
	writeTextProps,
} from '../xml/drawingml-write';
import { mergeRawChildren, rawChild } from './merge-raw';

/** `c:title`. A title without text is an automatic title ("Chart Title" / the series name). */
export function writeTitle(title: ChartText | undefined, raw?: XmlElement): XmlElement | undefined {
	if (!title) {
		return undefined;
	}
	let tx: XmlElement | undefined;
	let txPr: XmlElement | undefined;
	if (title.ref) {
		const cache =
			title.text === undefined ? [] : [el('c:pt', { idx: 0 }, [textEl('c:v', title.text)])];
		tx = el('c:tx', {}, [
			el('c:strRef', {}, [
				textEl('c:f', title.ref),
				el('c:strCache', {}, [valEl('c:ptCount', cache.length), ...cache]),
			]),
		]);
		txPr = writeTextProps(title.style, title.rotation);
	} else if (title.text !== undefined || title.paragraphs?.length) {
		tx = el('c:tx', {}, [writeRichText(title)]);
	} else {
		txPr = writeTextProps(title.style, title.rotation);
	}
	const out = el('c:title', {}, [
		tx,
		title.layout ? writeLayout(title.layout) : undefined,
		title.overlay !== undefined ? valEl('c:overlay', title.overlay) : undefined,
		txPr,
	]);
	return mergeRawChildren(out, raw, { order: 'title', owned: ['tx', 'layout', 'overlay', 'txPr'] });
}

const AXIS_TAGS: Record<ChartAxis['kind'], string> = {
	cat: 'c:catAx',
	val: 'c:valAx',
	date: 'c:dateAx',
	ser: 'c:serAx',
};

const AXIS_ORDER_KEYS: Record<ChartAxis['kind'], string> = {
	cat: 'catAx',
	val: 'valAx',
	date: 'dateAx',
	ser: 'serAx',
};

function gridlines(name: string, g: ChartAxis['majorGridlines']): XmlElement | undefined {
	if (!g) {
		return undefined;
	}
	return g === true ? el(name) : el(name, {}, [el('c:spPr', {}, [writeLine(g)])]);
}

function displayUnits(units: ChartAxis['displayUnits']): XmlElement | undefined {
	if (!units) {
		return undefined;
	}
	return el('c:dispUnits', {}, [
		typeof units.unit === 'number'
			? valEl('c:custUnit', formatXmlNumber(units.unit))
			: valEl('c:builtInUnit', units.unit),
		units.showLabel ? el('c:dispUnitsLbl', {}, [el('c:layout')]) : undefined,
	]);
}

const NUM = (name: string, v: number | undefined) =>
	v === undefined ? undefined : valEl(name, formatXmlNumber(v));

/** One axis element in schema order. */
export function writeAxis(axis: ChartAxis, raw?: XmlElement): XmlElement {
	const kind = axis.kind;
	const isCatLike = kind === 'cat' || kind === 'date';
	const children: Array<XmlElement | undefined> = [
		valEl('c:axId', axis.id),
		el('c:scaling', {}, [
			NUM('c:logBase', axis.logBase),
			axis.orientation ? valEl('c:orientation', axis.orientation) : undefined,
			NUM('c:max', axis.max),
			NUM('c:min', axis.min),
		]),
		axis.deleted !== undefined ? valEl('c:delete', axis.deleted) : undefined,
		valEl('c:axPos', axis.position),
		gridlines('c:majorGridlines', axis.majorGridlines),
		gridlines('c:minorGridlines', axis.minorGridlines),
		writeTitle(axis.title, rawChild(raw, 'title')),
		axis.numFmt
			? el('c:numFmt', {
					formatCode: axis.numFmt.code,
					sourceLinked: axis.numFmt.sourceLinked ? 1 : 0,
				})
			: undefined,
		axis.majorTickMark ? valEl('c:majorTickMark', axis.majorTickMark) : undefined,
		axis.minorTickMark ? valEl('c:minorTickMark', axis.minorTickMark) : undefined,
		axis.tickLabelPosition ? valEl('c:tickLblPos', axis.tickLabelPosition) : undefined,
		writeShapeProps(undefined, axis.line),
		writeTextProps(axis.textStyle, axis.labelRotation),
		valEl('c:crossAx', axis.crossAxisId),
		axis.crossesAt !== undefined
			? NUM('c:crossesAt', axis.crossesAt)
			: axis.crosses
				? valEl('c:crosses', axis.crosses)
				: undefined,
	];
	if (kind === 'cat') {
		children.push(
			valEl('c:auto', true),
			valEl('c:lblAlgn', 'ctr'),
			valEl('c:lblOffset', 100),
			axis.tickLabelSkip !== undefined ? valEl('c:tickLblSkip', axis.tickLabelSkip) : undefined,
			valEl('c:noMultiLvlLbl', Boolean(axis.noMultiLevelLabels)),
		);
	} else if (kind === 'date') {
		children.push(
			valEl('c:auto', true),
			valEl('c:lblOffset', 100),
			axis.baseTimeUnit ? valEl('c:baseTimeUnit', axis.baseTimeUnit) : undefined,
			NUM('c:majorUnit', axis.majorUnit),
			axis.majorTimeUnit ? valEl('c:majorTimeUnit', axis.majorTimeUnit) : undefined,
			NUM('c:minorUnit', axis.minorUnit),
		);
	} else if (kind === 'val') {
		children.push(
			axis.crossBetween ? valEl('c:crossBetween', axis.crossBetween) : undefined,
			NUM('c:majorUnit', axis.majorUnit),
			NUM('c:minorUnit', axis.minorUnit),
			displayUnits(axis.displayUnits),
		);
	} else {
		children.push(
			axis.tickLabelSkip !== undefined ? valEl('c:tickLblSkip', axis.tickLabelSkip) : undefined,
		);
	}
	const out = el(AXIS_TAGS[kind], {}, children);
	return mergeRawChildren(out, raw, {
		order: AXIS_ORDER_KEYS[kind],
		owned: [
			'axId',
			'scaling',
			'delete',
			'axPos',
			'majorGridlines',
			'minorGridlines',
			'title',
			'numFmt',
			'majorTickMark',
			'minorTickMark',
			'tickLblPos',
			'spPr',
			'txPr',
			'crossAx',
			'crosses',
			'crossesAt',
			'crossBetween',
			'majorUnit',
			'minorUnit',
			'dispUnits',
			'tickLblSkip',
			'noMultiLvlLbl',
			'baseTimeUnit',
			'majorTimeUnit',
		],
		preferRaw: isCatLike ? ['auto', 'lblAlgn', 'lblOffset'] : [],
	});
}

export function writeLegend(
	legend: ChartLegend | undefined,
	raw?: XmlElement,
): XmlElement | undefined {
	if (!legend) {
		return undefined;
	}
	const out = el('c:legend', {}, [
		valEl('c:legendPos', legend.position),
		...(legend.deletedEntries ?? []).map((idx) =>
			el('c:legendEntry', {}, [valEl('c:idx', idx), valEl('c:delete', true)]),
		),
		legend.layout ? writeLayout(legend.layout) : undefined,
		legend.overlay !== undefined ? valEl('c:overlay', legend.overlay) : undefined,
		writeShapeProps(legend.fill, legend.line),
		writeTextProps(legend.textStyle),
	]);
	return mergeRawChildren(out, raw, {
		order: 'legend',
		owned: ['legendPos', 'legendEntry', 'layout', 'overlay', 'spPr', 'txPr'],
	});
}
