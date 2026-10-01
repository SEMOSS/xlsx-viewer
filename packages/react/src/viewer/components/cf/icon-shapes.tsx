import type { ReactNode } from 'react';
import type { IconDescriptor } from 'xlsx-viewer-core';

/**
 * SVG drawings of Excel's icon-set icons on a 16 x 16 grid. Each function
 * returns the children of an `<svg viewBox="0 0 16 16">`.
 */

const EMPTY = '#D9D9D9';
const WHITE = '#FFFFFF';
const RIM = '#3F3F3F';
const POLE = '#595959';
const OUTLINE = 'rgba(0, 0, 0, 0.22)';

function fmt(n: number): string {
	return String(Math.round(n * 100) / 100);
}

/** Ten points of a five-pointed star (outer, inner, outer, ...) starting at the top. */
function starPoints(cx: number, cy: number, outer: number, inner: number): Array<[number, number]> {
	const out: Array<[number, number]> = [];
	for (let k = 0; k < 10; k++) {
		const angle = ((-90 + k * 36) * Math.PI) / 180;
		const r = k % 2 === 0 ? outer : inner;
		out.push([cx + r * Math.cos(angle), cy + r * Math.sin(angle)]);
	}
	return out;
}

const STAR = starPoints(8, 8.7, 7.4, 3);
const STAR_POINTS = STAR.map(([x, y]) => `${fmt(x)},${fmt(y)}`).join(' ');
/** Left half of the star: bottom inner vertex, the left points, then the top point. */
const STAR_LEFT_POINTS = [5, 6, 7, 8, 9, 0]
	.map((i) => `${fmt(STAR[i][0])},${fmt(STAR[i][1])}`)
	.join(' ');

const ARROW_UP = '8,1.5 14.5,8 10.5,8 10.5,14.5 5.5,14.5 5.5,8 1.5,8';
const ARROW_DOWN = '8,14.5 14.5,8 10.5,8 10.5,1.5 5.5,1.5 5.5,8 1.5,8';
const ARROW_RIGHT = '14.5,8 8,1.5 8,5.5 1.5,5.5 1.5,10.5 8,10.5 8,14.5';
/** A shorter up arrow that still fits the box once rotated by 45 degrees. */
const ARROW_DIAGONAL = '8,2.5 13.5,8 10.25,8 10.25,13.5 5.75,13.5 5.75,8 2.5,8';

function pieSlice(fraction: number, r: number): string {
	const angle = fraction * 2 * Math.PI;
	const x = 8 + r * Math.sin(angle);
	const y = 8 - r * Math.cos(angle);
	const large = fraction > 0.5 ? 1 : 0;
	return `M8 8 L8 ${fmt(8 - r)} A${r} ${r} 0 ${large} 1 ${fmt(x)} ${fmt(y)} Z`;
}

function quarter(color: string, fill: number): ReactNode {
	const r = 6.5;
	return (
		<>
			<circle
				cx='8'
				cy='8'
				r={r}
				fill={fill >= 1 ? color : WHITE}
				stroke={color}
				strokeWidth='1.25'
			/>
			{fill > 0 && fill < 1 ? <path d={pieSlice(fill, r)} fill={color} /> : null}
		</>
	);
}

function rating(color: string, fill: number): ReactNode {
	const lit = Math.round(fill * 4);
	return [4, 7, 10, 13].map((h, i) => (
		<rect
			key={i}
			x={1.5 + i * 3.5}
			y={14.5 - h}
			width='2.5'
			height={h}
			fill={i < lit ? color : EMPTY}
			shapeRendering='crispEdges'
		/>
	));
}

function boxes(color: string, fill: number): ReactNode {
	const lit = Math.round(fill * 4);
	const cells: Array<[number, number]> = [
		[2, 2],
		[9, 2],
		[2, 9],
		[9, 9],
	];
	return cells.map(([x, y], i) => (
		<rect
			key={i}
			x={x}
			y={y}
			width='5'
			height='5'
			fill={i < lit ? color : EMPTY}
			shapeRendering='crispEdges'
		/>
	));
}

function symbol(
	shape: 'check' | 'cross' | 'exclamation',
	color: string,
	circled: boolean,
): ReactNode {
	const ink = circled ? WHITE : color;
	const width = circled ? 2 : 2.5;
	let mark: ReactNode;
	if (shape === 'check') {
		const d = circled ? 'M4.5 8.25 L7 10.75 L11.5 5.5' : 'M2.5 8.5 L6.25 12.25 L13.5 3.75';
		mark = (
			<path
				d={d}
				fill='none'
				stroke={ink}
				strokeWidth={width}
				strokeLinecap='round'
				strokeLinejoin='round'
			/>
		);
	} else if (shape === 'cross') {
		const d = circled
			? 'M5.25 5.25 L10.75 10.75 M10.75 5.25 L5.25 10.75'
			: 'M3.5 3.5 L12.5 12.5 M12.5 3.5 L3.5 12.5';
		mark = <path d={d} fill='none' stroke={ink} strokeWidth={width} strokeLinecap='round' />;
	} else {
		mark = circled ? (
			<>
				<rect x='7.1' y='3.5' width='1.8' height='5.75' rx='0.5' fill={ink} />
				<circle cx='8' cy='11.5' r='1.1' fill={ink} />
			</>
		) : (
			<>
				<rect x='6.75' y='1.5' width='2.5' height='9' rx='0.6' fill={ink} />
				<circle cx='8' cy='13.4' r='1.4' fill={ink} />
			</>
		);
	}
	return circled ? (
		<>
			<circle cx='8' cy='8' r='7' fill={color} />
			{mark}
		</>
	) : (
		mark
	);
}

/** Draw one icon (children of a 16 x 16 SVG). */
export function renderIconShape(icon: IconDescriptor): ReactNode {
	const { color } = icon;
	switch (icon.shape) {
		case 'arrowUp':
			return <polygon points={ARROW_UP} fill={color} />;
		case 'arrowDown':
			return <polygon points={ARROW_DOWN} fill={color} />;
		case 'arrowRight':
			return <polygon points={ARROW_RIGHT} fill={color} />;
		case 'arrowUpRight':
			return <polygon points={ARROW_DIAGONAL} fill={color} transform='rotate(45 8 8)' />;
		case 'arrowDownRight':
			return <polygon points={ARROW_DIAGONAL} fill={color} transform='rotate(135 8 8)' />;
		case 'circle':
			return <circle cx='8' cy='8' r='6.5' fill={color} stroke={OUTLINE} strokeWidth='0.75' />;
		case 'trafficLight':
			return (
				<>
					<circle cx='8' cy='8' r='7.25' fill={RIM} />
					<circle cx='8' cy='8' r='5' fill={color} />
				</>
			);
		case 'flag':
			return (
				<>
					<rect x='3' y='1.5' width='1.5' height='13' fill={POLE} shapeRendering='crispEdges' />
					<path d='M4.5 2 H13.5 L11 5.25 L13.5 8.5 H4.5 Z' fill={color} />
				</>
			);
		case 'triangleUp':
			return <polygon points='8,2.5 14.5,13.5 1.5,13.5' fill={color} />;
		case 'triangleDown':
			return <polygon points='1.5,2.5 14.5,2.5 8,13.5' fill={color} />;
		case 'diamond':
			return <polygon points='8,1.5 14.5,8 8,14.5 1.5,8' fill={color} />;
		case 'dash':
			return <rect x='2' y='6' width='12' height='4' fill={color} shapeRendering='crispEdges' />;
		case 'star':
			return (
				<polygon
					points={STAR_POINTS}
					fill={color}
					stroke={color}
					strokeWidth='0.5'
					strokeLinejoin='round'
				/>
			);
		case 'starHalf':
			return (
				<>
					<polygon
						points={STAR_POINTS}
						fill={WHITE}
						stroke={color}
						strokeWidth='1'
						strokeLinejoin='round'
					/>
					<polygon points={STAR_LEFT_POINTS} fill={color} />
				</>
			);
		case 'starEmpty':
			return (
				<polygon
					points={STAR_POINTS}
					fill={WHITE}
					stroke={color}
					strokeWidth='1'
					strokeLinejoin='round'
				/>
			);
		case 'check':
		case 'cross':
		case 'exclamation':
			return symbol(icon.shape, color, icon.circled === true);
		case 'quarter':
			return quarter(color, icon.fill ?? 0);
		case 'rating':
			return rating(color, icon.fill ?? 0);
		case 'box':
			return boxes(color, icon.fill ?? 0);
		default:
			return null;
	}
}
