import { memo, useMemo } from 'react';

import type { Span } from '../../../engine/layout/types';
import { layoutSparklines } from '../../../engine/render/sparklines';
import { useViewerStore } from '../../state/use-store';
import { useGridRender } from './grid-context';

/** Sparklines (line, column, win/loss) drawn inside their cells for one pane. */
export const SparklinesLayer = memo(function SparklinesLayer({
	rows,
	cols,
}: {
	rows: Span;
	cols: Span;
}) {
	const ctx = useGridRender();
	const store = useViewerStore();
	const { geometry: g, sheet } = ctx;
	const shapes = useMemo(
		() =>
			sheet.sparklineGroups.length
				? layoutSparklines(store.workbook, sheet, g, rows, cols, ctx.colorCtx)
				: [],
		// docVersion: values come from the cells (recalculated on edits).
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[store, sheet, g, rows.start, rows.end, cols.start, cols.end, ctx.docVersion, ctx.colorCtx],
	);
	if (shapes.length === 0) {
		return null;
	}
	return (
		<svg
			className='xv-sparklines'
			style={{ position: 'absolute', left: 0, top: 0, overflow: 'visible', pointerEvents: 'none' }}
			width={1}
			height={1}
			aria-hidden
		>
			<defs>
				{shapes.map((s) => (
					<clipPath key={s.key} id={`xv-spark-${sheet.id}-${s.key.replace(/:/g, '-')}`}>
						<rect x={s.clip.x} y={s.clip.y} width={s.clip.width} height={s.clip.height} />
					</clipPath>
				))}
			</defs>
			{shapes.map((s) => (
				<g key={s.key} clipPath={`url(#xv-spark-${sheet.id}-${s.key.replace(/:/g, '-')})`}>
					{s.bars.map((b, i) => (
						<rect key={i} x={b.x} y={b.y} width={b.width} height={b.height} fill={b.color} />
					))}
					{s.paths.map((d, i) => (
						<path
							key={i}
							d={d}
							fill='none'
							stroke={s.lineColor}
							strokeWidth={s.lineWidth}
							strokeLinejoin='round'
							strokeLinecap='round'
						/>
					))}
					{s.markers.map((m, i) => (
						<circle key={`m${i}`} cx={m.cx} cy={m.cy} r={m.r} fill={m.color} />
					))}
				</g>
			))}
		</svg>
	);
});
