import { describe, expect, it } from 'vitest';
import type { ChartText } from '../../types/chart';
import type { DrawingColor, ShapeFill, ShapeLine } from '../../types/drawing';
import { parseXmlRoot } from '../../xml/parser';
import { serializeElement } from '../../xml/serializer';
import { readColor, readFill, readLine } from './drawingml-read';
import {
	readBodyRotation,
	readManualLayout,
	readRichText,
	readTextStyle,
	readTxPr,
} from './text-read';
import {
	formatXmlNumber,
	writeColor,
	writeFill,
	writeLayout,
	writeLine,
	writeRichText,
	writeRunProps,
	writeTextProps,
} from './drawingml-write';

const NS = 'xmlns:a="a" xmlns:c="c"';
const xml = (s: string) => parseXmlRoot(`<root ${NS}>${s}</root>`);
const wrap = (el: ReturnType<typeof writeFill>) =>
	parseXmlRoot(`<root ${NS}>${el ? serializeElement(el) : ''}</root>`);

describe('colors', () => {
	it('reads every color kind with transforms', () => {
		expect(readColor(xml('<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>'))).toEqual({
			kind: 'srgb',
			value: 'FF0000',
			transforms: [{ name: 'alpha', val: 50000 }],
		});
		expect(readColor(xml('<a:schemeClr val="accent2"/>'))).toEqual({
			kind: 'scheme',
			value: 'accent2',
		});
		expect(readColor(xml('<a:sysClr val="windowText" lastClr="000000"/>'))).toEqual({
			kind: 'sys',
			value: '000000',
		});
		expect(readColor(xml('<a:prstClr val="red"/>'))).toEqual({ kind: 'preset', value: 'red' });
		expect(readColor(xml('<a:hslClr hue="0" sat="100000" lum="50000"/>'))).toEqual({
			kind: 'hsl',
			value: '0,100000,50000',
		});
		expect(readColor(xml('<a:scrgbClr r="100000" g="0" b="0"/>'))).toEqual({
			kind: 'scrgb',
			value: '100000,0,0',
		});
		expect(readColor(xml('<a:other/>'))).toBeUndefined();
	});

	it('writes colors that read back identically', () => {
		const colors: DrawingColor[] = [
			{ kind: 'srgb', value: '4472C4', transforms: [{ name: 'lumMod', val: 75000 }] },
			{
				kind: 'scheme',
				value: 'tx1',
				transforms: [
					{ name: 'lumMod', val: 65000 },
					{ name: 'lumOff', val: 35000 },
				],
			},
			{ kind: 'preset', value: 'blue' },
			{ kind: 'hsl', value: '120,50000,50000' },
			{ kind: 'scrgb', value: '0,50000,100000' },
		];
		for (const c of colors) {
			expect(readColor(wrap(writeColor(c)))).toEqual(c);
		}
		expect(readColor(wrap(writeColor({ kind: 'sys', value: '000000' })))).toEqual({
			kind: 'srgb',
			value: '000000',
		});
	});
});

describe('fills and lines', () => {
	it('round-trips solid, gradient, pattern and empty fills', () => {
		const fills: ShapeFill[] = [
			{ type: 'none' },
			{ type: 'solid', color: { kind: 'srgb', value: '00B0F0' } },
			{
				type: 'gradient',
				angle: 45,
				stops: [
					{ pos: 0, color: { kind: 'srgb', value: 'FFFFFF' } },
					{ pos: 0.5, color: { kind: 'scheme', value: 'accent1' } },
				],
			},
			{
				type: 'gradient',
				path: 'circle',
				stops: [{ pos: 1, color: { kind: 'srgb', value: '000000' } }],
			},
			{
				type: 'pattern',
				preset: 'pct25',
				fg: { kind: 'srgb', value: '000000' },
				bg: { kind: 'srgb', value: 'FFFFFF' },
			},
		];
		for (const f of fills) {
			expect(readFill(wrap(writeFill(f)))).toEqual(f);
		}
		expect(writeFill({ type: 'picture', mediaPath: 'x' })).toBeUndefined();
	});

	it('round-trips outlines with width, dash, cap and arrow heads', () => {
		const line: ShapeLine = {
			width: 28575,
			cap: 'rnd',
			fill: { type: 'solid', color: { kind: 'srgb', value: 'C00000' } },
			dash: 'sysDot',
			headEnd: { type: 'triangle', w: 'med', len: 'med' },
		};
		const el = xml(serializeElement(writeLine(line)));
		expect(readLine(el.children[0] as never)).toEqual(line);
		expect(readLine(xml('<a:ln><a:custDash/></a:ln>').children[0] as never)).toEqual({
			dash: 'dash',
		});
	});
});

describe('text', () => {
	it('reads run properties', () => {
		const rPr = xml(
			'<a:defRPr sz="1197" b="1" i="0" u="sng" strike="noStrike"><a:solidFill><a:srgbClr val="595959"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr>',
		);
		expect(readTextStyle(rPr.children[0] as never)).toEqual({
			size: 11.97,
			bold: true,
			italic: false,
			underline: true,
			strike: false,
			color: { kind: 'srgb', value: '595959' },
			font: '+mn-lt',
		});
		expect(readTextStyle(xml('<a:rPr lang="en-US"/>').children[0] as never)).toBeUndefined();
	});

	it('round-trips text properties with rotation', () => {
		const style = { size: 9, bold: false, color: { kind: 'scheme', value: 'tx1' } as DrawingColor };
		const txPr = writeTextProps(style, -45)!;
		expect(
			readTxPr(parseXmlRoot(serializeElement(txPr).replace('<c:txPr>', `<c:txPr ${NS}>`))),
		).toEqual({ style, rotation: -45 });
		expect(writeTextProps(undefined)).toBeUndefined();
		expect(serializeElement(writeRunProps('a:rPr', { underline: false, strike: true }))).toBe(
			'<a:rPr lang="en-US" u="none" strike="sngStrike"/>',
		);
	});

	it("treats Excel's -60000000 rotation as automatic and maps vertical text", () => {
		expect(
			readBodyRotation(xml('<a:bodyPr rot="-60000000"/>').children[0] as never),
		).toBeUndefined();
		expect(readBodyRotation(xml('<a:bodyPr rot="-5400000"/>').children[0] as never)).toBe(-90);
		expect(readBodyRotation(xml('<a:bodyPr vert="vert270"/>').children[0] as never)).toBe(-90);
	});

	it('round-trips rich text with runs, breaks and alignment', () => {
		const text: ChartText = {
			text: 'Sales \nOverview',
			paragraphs: [
				{ align: 'ctr', runs: [{ text: ' Sales ', style: { bold: true } }, { text: '\n' }] },
				{ runs: [{ text: 'Overview', style: { color: { kind: 'srgb', value: 'FF0000' } } }] },
			],
			style: { size: 14 },
			rotation: -90,
		};
		const el = writeRichText(text);
		expect(serializeElement(el)).toContain('<a:t xml:space="preserve"> Sales </a:t>');
		const back = readRichText(
			parseXmlRoot(serializeElement(el).replace('<c:rich>', `<c:rich ${NS}>`)),
		);
		expect(back.paragraphs).toEqual(text.paragraphs);
		expect(back.style).toEqual({ size: 14 });
		expect(back.rotation).toBe(-90);
		expect(back.text).toBe(' Sales \n\nOverview');
	});
});

describe('layout and numbers', () => {
	it('round-trips manual layouts', () => {
		const layout = {
			layoutTarget: 'inner' as const,
			xMode: 'edge' as const,
			yMode: 'factor' as const,
			x: 0.1,
			y: -0.05,
			w: 0.8,
			h: 0.75,
		};
		const el = parseXmlRoot(
			serializeElement(writeLayout(layout)).replace('<c:layout>', `<c:layout ${NS}>`),
		);
		expect(readManualLayout(el)).toEqual(layout);
		expect(serializeElement(writeLayout(undefined))).toBe('<c:layout/>');
		expect(readManualLayout(xml('<c:layout/>').children[0] as never)).toBeUndefined();
	});

	it('formats numbers like Excel', () => {
		expect(formatXmlNumber(1e-7)).toBe('1E-7');
		expect(formatXmlNumber(2.5)).toBe('2.5');
	});
});
