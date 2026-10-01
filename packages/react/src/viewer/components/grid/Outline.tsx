import type { AxisOutline, OutlineGroup } from 'xlsx-viewer-core';
import { memo } from 'react';

import { useT } from '../../../i18n';
import type { Span } from '../../../engine/layout/types';
import { setOutlineCollapsed, showOutlineLevelCmd } from '../../commands/outline-commands';
import { useViewerStore } from '../../state/use-store';
import { useGridRender } from './grid-context';

/** Width of one outline level in the gutter, at 100% zoom. */
export const OUTLINE_LEVEL_PX = 14;

/** Size of the outline gutter for an axis (0 without outline). */
export function outlineGutterSize(levels: number, zoom: number): number {
	return levels > 0 ? Math.round(((levels + 1) * OUTLINE_LEVEL_PX + 4) * Math.max(zoom, 0.6)) : 0;
}

function ToggleButton({
	collapsed,
	size,
	style,
	label,
	onToggle,
}: {
	collapsed: boolean;
	size: number;
	style: React.CSSProperties;
	label: string;
	onToggle: () => void;
}) {
	return (
		<button
			type='button'
			className='xv-outline-toggle'
			aria-label={label}
			aria-expanded={!collapsed}
			title={label}
			style={{ ...style, width: size, height: size }}
			onPointerDown={(e) => e.stopPropagation()}
			onClick={onToggle}
		>
			{collapsed ? '+' : '-'}
		</button>
	);
}

interface GutterProps {
	outline: AxisOutline;
	/** Visible lines of the pane. */
	span: Span;
	/** Gutter thickness in px. */
	size: number;
}

/**
 * Row outline gutter (left of the row numbers): a bracket along each group's
 * detail rows and a +/- button on its summary row.
 */
export const RowOutlineGutter = memo(function RowOutlineGutter({
	outline,
	span,
	size,
}: GutterProps) {
	const t = useT();
	const { geometry: g, zoom } = useGridRender();
	const store = useViewerStore();
	const level = Math.max(10, OUTLINE_LEVEL_PX * Math.max(zoom, 0.6));
	const btn = Math.round(Math.min(level - 2, 12));
	const visible = (gr: OutlineGroup) =>
		(gr.end >= span.start && gr.start <= span.end) ||
		(gr.summary >= span.start && gr.summary <= span.end);
	const summaryHere = (gr: OutlineGroup) =>
		gr.summary < 0 ? gr.start === span.start : gr.summary >= span.start && gr.summary <= span.end;
	return (
		<div
			className='xv-outline-gutter'
			data-xv-object='outline'
			style={{
				position: 'absolute',
				left: 0,
				top: g.rowTop(span.start),
				width: size,
				height: g.rowTop(span.end + 1) - g.rowTop(span.start),
			}}
		>
			{outline.groups
				.filter((gr) => visible(gr) && (gr.summary < 0 || g.rowHeight(gr.summary) > 0))
				.map((gr) => {
					const x = (gr.level - 1) * level + level / 2;
					const top = g.rowTop(gr.start) - g.rowTop(span.start);
					const bottom = g.rowTop(gr.end + 1) - g.rowTop(span.start);
					const summaryMid =
						gr.summary >= 0
							? g.rowTop(gr.summary) + g.rowHeight(gr.summary) / 2 - g.rowTop(span.start)
							: bottom;
					return (
						<div key={`${gr.level}:${gr.start}`}>
							{!gr.collapsed && bottom > top && (
								<div
									className='xv-outline-bracket-v'
									style={{ left: x, top: top + 2, height: Math.max(0, bottom - top - 2) }}
								/>
							)}
							{summaryHere(gr) && (
								<ToggleButton
									collapsed={gr.collapsed}
									size={btn}
									label={gr.collapsed ? t('cmd.showDetail') : t('cmd.hideDetail')}
									style={{ position: 'absolute', left: x - btn / 2, top: summaryMid - btn / 2 }}
									onToggle={() => setOutlineCollapsed(store, 'rows', gr, !gr.collapsed)}
								/>
							)}
						</div>
					);
				})}
		</div>
	);
});

/** Column outline gutter (above the column letters). */
export const ColOutlineGutter = memo(function ColOutlineGutter({
	outline,
	span,
	size,
}: GutterProps) {
	const t = useT();
	const { geometry: g, zoom } = useGridRender();
	const store = useViewerStore();
	const level = Math.max(10, OUTLINE_LEVEL_PX * Math.max(zoom, 0.6));
	const btn = Math.round(Math.min(level - 2, 12));
	const visible = (gr: OutlineGroup) =>
		(gr.end >= span.start && gr.start <= span.end) ||
		(gr.summary >= span.start && gr.summary <= span.end);
	const summaryHere = (gr: OutlineGroup) =>
		gr.summary < 0 ? gr.start === span.start : gr.summary >= span.start && gr.summary <= span.end;
	return (
		<div
			className='xv-outline-gutter'
			data-xv-object='outline'
			style={{
				position: 'absolute',
				top: 0,
				left: g.colLeft(span.start),
				height: size,
				width: g.colLeft(span.end + 1) - g.colLeft(span.start),
			}}
		>
			{outline.groups
				.filter((gr) => visible(gr) && (gr.summary < 0 || g.colWidth(gr.summary) > 0))
				.map((gr) => {
					const y = (gr.level - 1) * level + level / 2;
					const left = g.colLeft(gr.start) - g.colLeft(span.start);
					const right = g.colLeft(gr.end + 1) - g.colLeft(span.start);
					const summaryMid =
						gr.summary >= 0
							? g.colLeft(gr.summary) + g.colWidth(gr.summary) / 2 - g.colLeft(span.start)
							: right;
					return (
						<div key={`${gr.level}:${gr.start}`}>
							{!gr.collapsed && right > left && (
								<div
									className='xv-outline-bracket-h'
									style={{ top: y, left: left + 2, width: Math.max(0, right - left - 2) }}
								/>
							)}
							{summaryHere(gr) && (
								<ToggleButton
									collapsed={gr.collapsed}
									size={btn}
									label={gr.collapsed ? t('cmd.showDetail') : t('cmd.hideDetail')}
									style={{ position: 'absolute', top: y - btn / 2, left: summaryMid - btn / 2 }}
									onToggle={() => setOutlineCollapsed(store, 'cols', gr, !gr.collapsed)}
								/>
							)}
						</div>
					);
				})}
		</div>
	);
});

/** The numbered level buttons in the corner (rows along the bottom, columns down the side). */
export function OutlineLevelButtons({
	rows,
	cols,
	rowGutter,
	colGutter,
	headerHeight,
}: {
	rows: AxisOutline;
	cols: AxisOutline;
	rowGutter: number;
	colGutter: number;
	headerHeight: number;
}) {
	const t = useT();
	const store = useViewerStore();
	const { zoom } = useGridRender();
	const level = Math.max(10, OUTLINE_LEVEL_PX * Math.max(zoom, 0.6));
	const btn = Math.round(Math.min(level - 1, 13));
	const buttons = (axis: 'rows' | 'cols', count: number) =>
		Array.from({ length: count + 1 }, (_, i) => {
			const n = i + 1;
			const label = t(axis === 'rows' ? 'outline.showRowLevel' : 'outline.showColumnLevel', {
				level: n,
			});
			const pos = i * level + (level - btn) / 2;
			const style: React.CSSProperties =
				axis === 'rows'
					? {
							position: 'absolute',
							left: pos,
							top: colGutter + (headerHeight - colGutter - btn) / 2,
							width: btn,
							height: btn,
						}
					: { position: 'absolute', top: pos, left: rowGutter + 2, width: btn, height: btn };
			return (
				<button
					key={`${axis}${n}`}
					type='button'
					className='xv-outline-level'
					title={label}
					aria-label={label}
					style={style}
					onPointerDown={(e) => e.stopPropagation()}
					onClick={() => showOutlineLevelCmd(store, axis, n)}
				>
					{n}
				</button>
			);
		});
	return (
		<div data-xv-object='outline-levels'>
			{rows.levels > 0 && buttons('rows', rows.levels)}
			{cols.levels > 0 && buttons('cols', cols.levels)}
		</div>
	);
}
