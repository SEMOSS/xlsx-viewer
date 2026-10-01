import type { DataBarBox } from '../../../engine/layout/types';

interface BarLike {
	start: number;
	end: number;
	color: string;
	borderColor?: string;
	gradient: boolean;
	negative?: boolean;
	axis?: number;
	axisColor?: string;
}

function isBar(v: unknown): v is BarLike {
	return typeof v === 'object' && v !== null && 'start' in v && 'end' in v && 'color' in v;
}

/** Conditional-formatting data bar inside a cell box. */
export function DataBarSlot({ box }: { box: DataBarBox }) {
	if (!isBar(box.bar)) {
		return null;
	}
	const bar = box.bar;
	const inset = 2;
	const innerW = Math.max(0, box.width - inset * 2);
	const left = box.x + inset + innerW * Math.min(bar.start, bar.end);
	const width = innerW * Math.abs(bar.end - bar.start);
	const background = bar.gradient
		? `linear-gradient(${bar.negative ? 270 : 90}deg, ${bar.color}, #ffffff)`
		: bar.color;
	return (
		<>
			<div
				style={{
					position: 'absolute',
					left,
					top: box.y + inset,
					width,
					height: Math.max(0, box.height - inset * 2 - 1),
					background,
					border: bar.borderColor ? `1px solid ${bar.borderColor}` : undefined,
					boxSizing: 'border-box',
				}}
			/>
			{bar.axis !== undefined && (
				<div
					style={{
						position: 'absolute',
						left: box.x + inset + innerW * bar.axis,
						top: box.y,
						width: 1,
						height: box.height - 1,
						borderLeft: `1px dashed ${bar.axisColor ?? '#000000'}`,
					}}
				/>
			)}
		</>
	);
}
