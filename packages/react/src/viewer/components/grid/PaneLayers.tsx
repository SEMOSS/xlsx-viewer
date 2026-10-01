import { memo, useMemo } from 'react';

import { layoutViewport } from '../../../engine/layout/viewport-layout';
import type { Span, ViewportLayout } from '../../../engine/layout/types';
import { segmentLines } from '../../../engine/render/borders';
import { CellTextView } from './CellTextView';
import { useGridRender } from './grid-context';
import { DataBarSlot } from './DataBarSlot';
import { SparklinesLayer } from './SparklinesLayer';

interface PaneLayersProps {
	rows: Span;
	cols: Span;
}

/** Gridlines, fills, borders and texts of one pane, in sheet coordinates. */
export const PaneLayers = memo(function PaneLayers({ rows, cols }: PaneLayersProps) {
	const ctx = useGridRender();
	const { geometry: g } = ctx;
	const layout = useMemo<ViewportLayout>(
		() =>
			layoutViewport({
				sheet: ctx.sheet,
				geometry: g,
				rows,
				cols,
				styles: ctx.styles,
				merges: ctx.merges,
				cf: ctx.cf,
				table: ctx.table,
				canvasColor: '#ffffff',
				describe: {
					zoom: ctx.zoom,
					colorCtx: ctx.colorCtx,
					date1904: ctx.date1904,
					defaultTextColor: '#000000',
					showFormulas: ctx.showFormulas,
					hideZeros: !ctx.showZeros,
				},
			}),
		// docVersion invalidates the layout when the sheet content changes.
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[
			ctx.sheet,
			ctx.docVersion,
			g,
			rows.start,
			rows.end,
			cols.start,
			cols.end,
			ctx.showFormulas,
			ctx.showZeros,
			ctx.cf,
			ctx.table,
		],
	);

	const x0 = g.colLeft(cols.start);
	const y0 = g.rowTop(rows.start);
	const x1 = g.colLeft(cols.end + 1);
	const y1 = g.rowTop(rows.end + 1);
	const width = Math.max(0, x1 - x0);
	const height = Math.max(0, y1 - y0);

	const gridlines = useMemo(() => {
		if (!ctx.showGridLines || width === 0 || height === 0) {
			return null;
		}
		let d = '';
		for (let c = cols.start; c <= cols.end; c++) {
			if (g.colWidth(c) === 0) {
				continue;
			}
			const x = g.colLeft(c + 1) - x0 - 0.5;
			d += `M${x} 0V${height}`;
		}
		for (let r = rows.start; r <= rows.end; r++) {
			if (g.rowHeight(r) === 0) {
				continue;
			}
			const y = g.rowTop(r + 1) - y0 - 0.5;
			d += `M0 ${y}H${width}`;
		}
		return d;
	}, [ctx.showGridLines, cols.start, cols.end, rows.start, rows.end, g, x0, y0, width, height]);

	return (
		<>
			{gridlines !== null && (
				<svg
					className='xv-gridlines'
					width={width}
					height={height}
					style={{ position: 'absolute', left: x0, top: y0 }}
					shapeRendering='crispEdges'
					aria-hidden
				>
					<path d={gridlines} stroke={ctx.gridlineColor} strokeWidth={1} fill='none' />
					{layout.gridlineMasks.map((m, i) => (
						<rect
							// biome-ignore lint: masks are positional
							key={i}
							x={m.x - x0}
							y={m.y - y0}
							width={Math.max(0, m.width)}
							height={Math.max(0, m.height)}
							fill='#ffffff'
						/>
					))}
				</svg>
			)}
			<div className='xv-fills' aria-hidden>
				{layout.fills.map((f) => (
					<div
						key={f.key}
						style={{
							position: 'absolute',
							left: f.x,
							top: f.y,
							width: f.width,
							height: f.height,
							backgroundColor: f.css.backgroundColor,
							backgroundImage: f.css.backgroundImage,
						}}
					/>
				))}
				{layout.dataBars.map((b) => (
					<DataBarSlot key={`db:${b.key}`} box={b} />
				))}
			</div>
			<SparklinesLayer rows={rows} cols={cols} />
			{layout.borders.length > 0 && (
				<svg
					className='xv-borders'
					width={width + 4}
					height={height + 4}
					style={{ position: 'absolute', left: x0 - 2, top: y0 - 2, overflow: 'visible' }}
					aria-hidden
				>
					<g transform='translate(2 2)'>
						{layout.borders.flatMap((seg, i) =>
							segmentLines(seg).map((l, j) => (
								<line
									// biome-ignore lint: border segments are positional
									key={`${i}:${j}`}
									x1={l.x1 - x0}
									y1={l.y1 - y0}
									x2={l.x2 - x0}
									y2={l.y2 - y0}
									stroke={seg.color}
									strokeWidth={l.width}
									strokeDasharray={l.dash}
									shapeRendering={l.dash ? undefined : 'crispEdges'}
								/>
							)),
						)}
					</g>
				</svg>
			)}
			<div className='xv-texts'>
				{layout.texts.map((t) => (
					<CellTextView key={t.key} box={t} zoom={ctx.zoom} />
				))}
			</div>
		</>
	);
});
