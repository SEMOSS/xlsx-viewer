/** Small chart-type pictures for the Change Chart Type gallery (Excel-style two-series icons). */
import type { ChartTypeTarget } from 'xlsx-viewer-core';
import type { ReactNode } from 'react';

const A = '#4472C4';
const B = '#ED7D31';
const AXIS = '#A6A6A6';

const axis = <path key='axis' d='M3 3 V21 H29' fill='none' stroke={AXIS} strokeWidth={1} />;

function columns(kind: 'clustered' | 'stacked' | 'percent', horizontal: boolean): ReactNode {
	const a = [8, 13, 10];
	const b = [6, 9, 12];
	const rects: ReactNode[] = [];
	a.forEach((h, i) => {
		const pos = 6 + i * 8;
		let shapes: Array<[number, number, number, string]>;
		if (kind === 'clustered') {
			shapes = [
				[pos, 0, h, A],
				[pos + 3, 0, b[i], B],
			];
		} else {
			const total = kind === 'percent' ? 17 : h + b[i] > 17 ? 17 : h + b[i];
			const ha = kind === 'percent' ? Math.round((17 * h) / (h + b[i])) : Math.min(h, total);
			shapes = [
				[pos, 0, ha, A],
				[pos, ha, total - ha, B],
			];
		}
		for (const [p, from, len, fill] of shapes) {
			const w = kind === 'clustered' ? 3 : 5;
			rects.push(
				horizontal ? (
					<rect
						key={`${i}-${fill}`}
						x={3 + from}
						y={21 - p - w}
						width={len}
						height={w}
						fill={fill}
					/>
				) : (
					<rect key={`${i}-${fill}`} x={p} y={21 - from - len} width={w} height={len} fill={fill} />
				),
			);
		}
	});
	return rects;
}

const lineA = 'M5 16 L12 10 L19 13 L27 6';
const lineB = 'M5 19 L12 15 L19 17 L27 12';

/** Dots at the vertices of a polyline path. */
function markers(d: string, fill: string): ReactNode {
	const nums = (d.match(/\d+(\.\d+)?/g) ?? []).map(Number);
	const dots: ReactNode[] = [];
	for (let i = 0; i + 1 < nums.length; i += 2) {
		dots.push(<circle key={`${fill}-${i}`} cx={nums[i]} cy={nums[i + 1]} r={1.6} fill={fill} />);
	}
	return dots;
}

const GLYPHS: Record<ChartTypeTarget, () => ReactNode> = {
	columnClustered: () => [axis, columns('clustered', false)],
	columnStacked: () => [axis, columns('stacked', false)],
	columnPercent: () => [axis, columns('percent', false)],
	barClustered: () => [axis, columns('clustered', true)],
	barStacked: () => [axis, columns('stacked', true)],
	barPercent: () => [axis, columns('percent', true)],
	line: () => [
		axis,
		<path key='a' d={lineA} fill='none' stroke={A} strokeWidth={1.6} />,
		<path key='b' d={lineB} fill='none' stroke={B} strokeWidth={1.6} />,
	],
	lineMarkers: () => [
		axis,
		<path key='a' d={lineA} fill='none' stroke={A} strokeWidth={1.2} />,
		<path key='b' d={lineB} fill='none' stroke={B} strokeWidth={1.2} />,
		markers(lineA, A),
		markers(lineB, B),
	],
	area: () => [
		axis,
		<path key='a' d={`${lineA} V21 H5 Z`} fill={A} />,
		<path key='b' d={`${lineB} V21 H5 Z`} fill={B} opacity={0.9} />,
	],
	areaStacked: () => [
		axis,
		<path key='a' d='M5 13 L12 8 L19 10 L27 4 V12 L19 17 L12 15 L5 19 Z' fill={A} />,
		<path key='b' d={`${lineB} V21 H5 Z`} fill={B} />,
	],
	pie: () => [
		<circle key='c' cx={16} cy={12} r={9} fill={A} />,
		<path key='s1' d='M16 12 L16 3 A9 9 0 0 1 24.2 15.7 Z' fill={B} />,
		<path key='s2' d='M16 12 L24.2 15.7 A9 9 0 0 1 12 20.1 Z' fill='#A5A5A5' />,
	],
	doughnut: () => [
		<circle key='a' cx={16} cy={12} r={7} fill='none' stroke={A} strokeWidth={4} />,
		<path key='b' d='M16 5 A7 7 0 0 1 22.6 14.3' fill='none' stroke={B} strokeWidth={4} />,
	],
	scatter: () => [
		axis,
		markers('M6 15 L10 9 L14 13 L19 7 L23 10 L27 5', A),
		markers('M7 19 L12 17 L17 18 L22 14 L26 15', B),
	],
	scatterLines: () => [
		axis,
		<path
			key='a'
			d='M6 15 L10 9 L14 13 L19 7 L23 10 L27 5'
			fill='none'
			stroke={A}
			strokeWidth={1.2}
		/>,
		markers('M6 15 L10 9 L14 13 L19 7 L23 10 L27 5', A),
	],
	radar: () => [
		<path
			key='w'
			d='M16 3 L24.5 9 L21.5 20 L10.5 20 L7.5 9 Z'
			fill='none'
			stroke={AXIS}
			strokeWidth={1}
		/>,
		<path
			key='a'
			d='M16 6 L22 10 L19 17 L12 18 L10 10 Z'
			fill='none'
			stroke={A}
			strokeWidth={1.4}
		/>,
		<path
			key='b'
			d='M16 9 L20 11 L18 15 L13 16 L12 11 Z'
			fill='none'
			stroke={B}
			strokeWidth={1.4}
		/>,
	],
};

/** A 32 x 24 picture of a chart type. */
export function ChartTypeGlyph({ type, size = 32 }: { type: ChartTypeTarget; size?: number }) {
	return (
		<svg width={size} height={(size * 3) / 4} viewBox='0 0 32 24' aria-hidden focusable='false'>
			{GLYPHS[type]()}
		</svg>
	);
}
