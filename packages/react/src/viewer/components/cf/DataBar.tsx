import { memo, useId } from 'react';
import type { DataBarResult } from 'xlsx-viewer-core';

export interface DataBarProps {
	bar: DataBarResult;
	/** Height of the cell box in CSS pixels. */
	height: number;
	className?: string;
}

function clamp01(v: number): number {
	return v <= 0 ? 0 : v >= 1 ? 1 : v;
}

function pct(fraction: number): string {
	return `${Math.round(fraction * 100_000) / 1000}%`;
}

/** Gap between the bar and the cell's left / right edges, in CSS pixels. */
const HORIZONTAL_INSET = 1;

/** Gap above and below the bar, like Excel's (none for very short cells). */
function verticalPadding(height: number): number {
	return height >= 8 ? 2 : height >= 4 ? 1 : 0;
}

function DataBarImpl({ bar, height, className }: DataBarProps) {
	const gradientId = `xv-databar-${useId().replace(/[^A-Za-z0-9_-]/g, '')}`;
	const pad = verticalPadding(height);
	const barHeight = Math.max(0, height - pad * 2);
	const left = clamp01(Math.min(bar.start, bar.end));
	const width = clamp01(Math.max(bar.start, bar.end)) - left;
	// The gradient fades from the bar's origin (axis side) towards its value end.
	const originRight = bar.negative !== (bar.direction === 'rtl');
	const fill = bar.gradient ? `url(#${gradientId})` : bar.color;
	return (
		<svg
			className={className}
			aria-hidden='true'
			focusable='false'
			width='100%'
			height={height}
			data-negative={bar.negative ? 'true' : undefined}
			style={{
				position: 'absolute',
				left: HORIZONTAL_INSET,
				top: 0,
				width: `calc(100% - ${HORIZONTAL_INSET * 2}px)`,
				height,
				overflow: 'visible',
				pointerEvents: 'none',
			}}
		>
			{bar.gradient ? (
				<defs>
					<linearGradient
						id={gradientId}
						x1={originRight ? 1 : 0}
						x2={originRight ? 0 : 1}
						y1={0}
						y2={0}
					>
						<stop offset='0' stopColor={bar.color} />
						<stop offset='1' stopColor='#FFFFFF' />
					</linearGradient>
				</defs>
			) : null}
			{width > 0 && barHeight > 0 ? (
				<rect
					x={pct(left)}
					y={pad}
					width={pct(width)}
					height={barHeight}
					fill={fill}
					stroke={bar.borderColor}
					strokeWidth={bar.borderColor ? 1 : undefined}
					vectorEffect='non-scaling-stroke'
					shapeRendering='crispEdges'
				/>
			) : null}
			{bar.axis !== undefined ? (
				<line
					x1={pct(clamp01(bar.axis))}
					x2={pct(clamp01(bar.axis))}
					y1={0}
					y2={height}
					stroke={bar.axisColor ?? '#000000'}
					strokeWidth={1}
					strokeDasharray='2 2'
					shapeRendering='crispEdges'
				/>
			) : null}
		</svg>
	);
}

/**
 * A conditional-formatting data bar. Renders an absolutely positioned SVG
 * covering its (relatively positioned) cell box: the bar (solid, or a
 * gradient fading to white), its optional border and the dashed axis line.
 */
export const DataBar = memo(DataBarImpl);
DataBar.displayName = 'DataBar';
