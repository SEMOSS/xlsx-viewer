/** Positioning helpers for multi-line text blocks. */
import { type ResolvedText, textAttrs } from './context';
import { baselineOffset } from './text';
import type { ChartNodeMeta, ChartTextNode, ChartTextRun } from './types';

/** Height of a block of `lineCount` lines. */
export function blockHeight(t: ResolvedText, lineCount: number): number {
	return Math.max(1, lineCount) * t.lineHeight;
}

/**
 * A text node whose block is anchored at `x` (per `anchor`) and starts at
 * `top`. `rotation` turns the block around (`pivotX`, `pivotY`), defaulting to
 * the anchor point at the block's vertical middle.
 */
export function textBlock(
	t: ResolvedText,
	lines: string[],
	x: number,
	top: number,
	anchor: ChartTextNode['anchor'],
	meta: ChartNodeMeta = {},
	extra: { rotation?: number; pivotX?: number; pivotY?: number; runs?: ChartTextRun[][] } = {},
): ChartTextNode {
	const node: ChartTextNode = {
		type: 'text',
		x,
		y: top + baselineOffset(t.sizePx),
		lines,
		anchor,
		...textAttrs(t),
		...meta,
	};
	if (extra.runs) {
		node.runs = extra.runs;
	}
	if (extra.rotation) {
		node.rotation = extra.rotation;
		node.pivotX = extra.pivotX ?? x;
		node.pivotY = extra.pivotY ?? top + blockHeight(t, lines.length) / 2;
	}
	return node;
}
