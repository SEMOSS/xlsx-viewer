import type { ColorContext, FillSpec, PatternType } from 'xlsx-viewer-core';
import { resolveColorCss } from 'xlsx-viewer-core';

/**
 * 8x8 pixel masks for Excel's pattern fills (1 = foreground color).
 * Each string is one row; patterns repeat seamlessly.
 */
const PATTERNS: Partial<Record<PatternType, string[]>> = {
	gray0625: [
		'10000000',
		'00000000',
		'00000000',
		'00000000',
		'00001000',
		'00000000',
		'00000000',
		'00000000',
	],
	gray125: [
		'10001000',
		'00000000',
		'00100010',
		'00000000',
		'10001000',
		'00000000',
		'00100010',
		'00000000',
	],
	lightGray: [
		'10001000',
		'00100010',
		'10001000',
		'00100010',
		'10001000',
		'00100010',
		'10001000',
		'00100010',
	],
	mediumGray: [
		'10101010',
		'01010101',
		'10101010',
		'01010101',
		'10101010',
		'01010101',
		'10101010',
		'01010101',
	],
	darkGray: [
		'11101110',
		'10111011',
		'11101110',
		'10111011',
		'11101110',
		'10111011',
		'11101110',
		'10111011',
	],
	darkHorizontal: [
		'11111111',
		'11111111',
		'00000000',
		'00000000',
		'11111111',
		'11111111',
		'00000000',
		'00000000',
	],
	darkVertical: [
		'11001100',
		'11001100',
		'11001100',
		'11001100',
		'11001100',
		'11001100',
		'11001100',
		'11001100',
	],
	darkDown: [
		'11001100',
		'01100110',
		'00110011',
		'10011001',
		'11001100',
		'01100110',
		'00110011',
		'10011001',
	],
	darkUp: [
		'00110011',
		'01100110',
		'11001100',
		'10011001',
		'00110011',
		'01100110',
		'11001100',
		'10011001',
	],
	darkGrid: [
		'11111111',
		'11001100',
		'11001100',
		'11111111',
		'11111111',
		'11001100',
		'11001100',
		'11111111',
	],
	darkTrellis: [
		'11101110',
		'01110111',
		'11101110',
		'01110111',
		'11101110',
		'01110111',
		'11101110',
		'01110111',
	],
	lightHorizontal: [
		'11111111',
		'00000000',
		'00000000',
		'00000000',
		'11111111',
		'00000000',
		'00000000',
		'00000000',
	],
	lightVertical: [
		'10001000',
		'10001000',
		'10001000',
		'10001000',
		'10001000',
		'10001000',
		'10001000',
		'10001000',
	],
	lightDown: [
		'10001000',
		'01000100',
		'00100010',
		'00010001',
		'10001000',
		'01000100',
		'00100010',
		'00010001',
	],
	lightUp: [
		'00010001',
		'00100010',
		'01000100',
		'10001000',
		'00010001',
		'00100010',
		'01000100',
		'10001000',
	],
	lightGrid: [
		'11111111',
		'10001000',
		'10001000',
		'10001000',
		'11111111',
		'10001000',
		'10001000',
		'10001000',
	],
	lightTrellis: [
		'10001000',
		'01010101',
		'00100010',
		'01010101',
		'10001000',
		'01010101',
		'00100010',
		'01010101',
	],
};

const patternCache = new Map<string, string>();

function patternDataUri(type: PatternType, fg: string): string | undefined {
	const mask = PATTERNS[type];
	if (!mask) {
		return undefined;
	}
	const key = `${type}|${fg}`;
	const hit = patternCache.get(key);
	if (hit) {
		return hit;
	}
	let rects = '';
	mask.forEach((row, y) => {
		for (let x = 0; x < row.length; x++) {
			if (row[x] === '1') {
				rects += `<rect x="${x}" y="${y}" width="1" height="1"/>`;
			}
		}
	});
	const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8" shape-rendering="crispEdges"><g fill="${fg}">${rects}</g></svg>`;
	const uri = `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
	patternCache.set(key, uri);
	return uri;
}

export interface FillCss {
	backgroundColor?: string;
	backgroundImage?: string;
}

/** CSS background for a cell fill, or `undefined` when the cell has no fill. */
export function fillToCss(fill: FillSpec | undefined, ctx: ColorContext): FillCss | undefined {
	if (!fill) {
		return undefined;
	}
	if (fill.type === 'gradient') {
		const stops = [...fill.stops]
			.sort((a, b) => a.position - b.position)
			.map((s) => `${resolveColorCss(s.color, ctx, '#ffffff')} ${Math.round(s.position * 100)}%`)
			.join(', ');
		if (fill.gradientType === 'path') {
			const cx = Math.round((((fill.left ?? 0.5) + (fill.right ?? 0.5)) / 2) * 100);
			const cy = Math.round((((fill.top ?? 0.5) + (fill.bottom ?? 0.5)) / 2) * 100);
			return { backgroundImage: `radial-gradient(circle at ${cx}% ${cy}%, ${stops})` };
		}
		return { backgroundImage: `linear-gradient(${(fill.degree ?? 0) + 90}deg, ${stops})` };
	}
	if (fill.patternType === 'none') {
		return undefined;
	}
	if (fill.patternType === 'solid') {
		// Solid fills paint the foreground color; Excel treats a missing fgColor as black.
		return { backgroundColor: resolveColorCss(fill.fgColor, ctx, '#000000') };
	}
	const fg = resolveColorCss(fill.fgColor, ctx, '#000000') ?? '#000000';
	const bg = resolveColorCss(fill.bgColor, ctx, '#ffffff') ?? '#ffffff';
	return { backgroundColor: bg, backgroundImage: patternDataUri(fill.patternType, fg) };
}

/** Fill color used where a single color is needed (e.g. the ribbon's fill swatch). */
export function fillSwatchColor(fill: FillSpec | undefined, ctx: ColorContext): string | undefined {
	const css = fillToCss(fill, ctx);
	return css?.backgroundColor;
}
