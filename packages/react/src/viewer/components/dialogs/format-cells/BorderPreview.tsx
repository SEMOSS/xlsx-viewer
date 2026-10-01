import type { BorderEdge, BorderEdgeName, Theme } from 'xlsx-viewer-core';
import { resolveColorCss } from 'xlsx-viewer-core';
import type { PointerEvent, ReactNode } from 'react';

import { borderStroke } from '../../../../engine/render/borders';
import { useT } from '../../../../i18n';

/** What the preview shows for an edge: a line, nothing, or a gray "mixed" line. */
export type EdgeDisplay = BorderEdge | null | 'mixed';

const W = 190;
const H = 130;
const PAD = 14;

/** Draw a border line between two points in the style Excel uses. */
export function StrokeLine({
	x1,
	y1,
	x2,
	y2,
	edge,
	theme,
}: {
	x1: number;
	y1: number;
	x2: number;
	y2: number;
	edge: EdgeDisplay;
	theme: Theme;
}): ReactNode {
	if (edge === null) {
		return null;
	}
	if (edge === 'mixed') {
		return <line x1={x1} y1={y1} x2={x2} y2={y2} stroke='#a6a6a6' strokeWidth={1} />;
	}
	const stroke = borderStroke(edge.style);
	if (!stroke) {
		return null;
	}
	const color = resolveColorCss(edge.color, { theme }, '#000000') ?? '#000000';
	if (stroke.double) {
		const dx = y1 === y2 ? 0 : 1.5;
		const dy = y1 === y2 ? 1.5 : 0;
		return (
			<g stroke={color} strokeWidth={1}>
				<line x1={x1 - dx} y1={y1 - dy} x2={x2 - dx} y2={y2 - dy} />
				<line x1={x1 + dx} y1={y1 + dy} x2={x2 + dx} y2={y2 + dy} />
			</g>
		);
	}
	return (
		<line
			x1={x1}
			y1={y1}
			x2={x2}
			y2={y2}
			stroke={color}
			strokeWidth={stroke.width}
			strokeDasharray={stroke.dash}
		/>
	);
}

interface BorderPreviewProps {
	edges: Record<BorderEdgeName, EdgeDisplay>;
	multiRow: boolean;
	multiCol: boolean;
	theme: Theme;
	onToggle: (edge: BorderEdgeName) => void;
}

/** Distance from a point to a segment. */
function distance(px: number, py: number, x1: number, y1: number, x2: number, y2: number): number {
	const dx = x2 - x1;
	const dy = y2 - y1;
	const t = Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy || 1)));
	return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
}

/** The Border tab's preview diagram; clicking near a line toggles it. */
export function BorderPreview({ edges, multiRow, multiCol, theme, onToggle }: BorderPreviewProps) {
	// The word drawn in each cell (`t` below is the top edge, not the translator).
	const sampleText = useT()('dialog.formatCells.previewText');
	const l = PAD;
	const t = PAD;
	const r = W - PAD;
	const b = H - PAD;
	const mx = (l + r) / 2;
	const my = (t + b) / 2;
	const cellsX = multiCol
		? [
				[l, mx],
				[mx, r],
			]
		: [[l, r]];
	const cellsY = multiRow
		? [
				[t, my],
				[my, b],
			]
		: [[t, b]];
	const lines: Array<{ name: BorderEdgeName; segs: Array<[number, number, number, number]> }> = [
		{ name: 'top', segs: [[l, t, r, t]] },
		{ name: 'bottom', segs: [[l, b, r, b]] },
		{ name: 'left', segs: [[l, t, l, b]] },
		{ name: 'right', segs: [[r, t, r, b]] },
		...(multiRow
			? [
					{
						name: 'insideHorizontal' as const,
						segs: [[l, my, r, my] as [number, number, number, number]],
					},
				]
			: []),
		...(multiCol
			? [
					{
						name: 'insideVertical' as const,
						segs: [[mx, t, mx, b] as [number, number, number, number]],
					},
				]
			: []),
		{
			name: 'diagonalUp',
			segs: cellsY.flatMap(([y1, y2]) =>
				cellsX.map(([x1, x2]) => [x1, y2, x2, y1] as [number, number, number, number]),
			),
		},
		{
			name: 'diagonalDown',
			segs: cellsY.flatMap(([y1, y2]) =>
				cellsX.map(([x1, x2]) => [x1, y1, x2, y2] as [number, number, number, number]),
			),
		},
	];

	const onPointerDown = (e: PointerEvent<SVGSVGElement>) => {
		const box = e.currentTarget.getBoundingClientRect();
		if (box.width === 0) {
			return;
		}
		const px = ((e.clientX - box.left) / box.width) * W;
		const py = ((e.clientY - box.top) / box.height) * H;
		let best: BorderEdgeName | null = null;
		let bestD = 9;
		for (const line of lines) {
			for (const [x1, y1, x2, y2] of line.segs) {
				// Diagonals are harder to hit near the corners where edges meet.
				const d = distance(px, py, x1, y1, x2, y2) + (line.name.startsWith('diagonal') ? 2 : 0);
				if (d < bestD) {
					bestD = d;
					best = line.name;
				}
			}
		}
		if (best) {
			onToggle(best);
		}
	};

	const corner = (x: number, y: number, dx: number, dy: number) => (
		<path
			d={`M ${x + dx * 6} ${y} H ${x} V ${y + dy * 6}`}
			className='fill-none stroke-muted-foreground'
			strokeWidth={1}
		/>
	);

	return (
		<svg
			width={W}
			height={H}
			viewBox={`0 0 ${W} ${H}`}
			aria-hidden='true'
			className='rounded-sm bg-white'
			onPointerDown={onPointerDown}
			data-testid='border-preview'
		>
			{corner(l - 2, t - 2, -1, -1)}
			{corner(r + 2, t - 2, 1, -1)}
			{corner(l - 2, b + 2, -1, 1)}
			{corner(r + 2, b + 2, 1, 1)}
			{cellsY.map(([y1, y2]) =>
				cellsX.map(([x1, x2]) => (
					<text
						key={`${x1}-${y1}`}
						x={(x1 + x2) / 2}
						y={(y1 + y2) / 2 + 4}
						textAnchor='middle'
						fontSize={11}
						fill='#595959'
					>
						{sampleText}
					</text>
				)),
			)}
			{lines.map((line) =>
				line.segs.map(([x1, y1, x2, y2], i) => (
					<StrokeLine
						key={`${line.name}-${i}`}
						x1={x1}
						y1={y1}
						x2={x2}
						y2={y2}
						edge={edges[line.name]}
						theme={theme}
					/>
				)),
			)}
		</svg>
	);
}

/** 16px icon of a cell with one edge drawn (edge toggle buttons). */
export function EdgeIcon({ edge }: { edge: BorderEdgeName }) {
	const paths: Record<BorderEdgeName, string> = {
		top: 'M2 2 H14',
		bottom: 'M2 14 H14',
		left: 'M2 2 V14',
		right: 'M14 2 V14',
		insideHorizontal: 'M2 8 H14',
		insideVertical: 'M8 2 V14',
		diagonalUp: 'M2 14 L14 2',
		diagonalDown: 'M2 2 L14 14',
	};
	return (
		<svg width={16} height={16} viewBox='0 0 16 16' aria-hidden='true'>
			<rect
				x={2}
				y={2}
				width={12}
				height={12}
				className='fill-none stroke-muted-foreground'
				strokeDasharray='1 1.5'
				strokeWidth={1}
			/>
			<path d={paths[edge]} className='stroke-foreground' strokeWidth={1.6} />
		</svg>
	);
}
