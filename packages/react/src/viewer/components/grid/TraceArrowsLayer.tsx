import { isCellError } from 'xlsx-viewer-core';
import { memo } from 'react';

import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import { useStoreSelector, useViewerStore } from '../../state/use-store';

const BLUE = '#1f5bd8';
const RED = '#d13438';
const EXTERNAL = '#3a3a3a';

/** Formula auditing arrows (Trace Precedents / Dependents) in sheet coordinates. */
export const TraceArrowsLayer = memo(function TraceArrowsLayer({
	geometry: g,
}: {
	geometry: SheetGeometry;
}) {
	const store = useViewerStore();
	const arrows = useStoreSelector((s) => s.ui.traceArrows);
	const sheetIndex = useStoreSelector((s) => s.ui.activeSheet);
	const visible = arrows.filter((a) => a.sheet === sheetIndex);
	if (visible.length === 0) {
		return null;
	}
	const sheet = store.workbook.sheets[sheetIndex];
	const center = (row: number, col: number) => {
		const r = g.cellRect(row, col);
		return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
	};
	return (
		<svg
			className='xv-trace-arrows'
			style={{
				position: 'absolute',
				left: 0,
				top: 0,
				overflow: 'visible',
				pointerEvents: 'none',
				zIndex: 11,
			}}
			width={1}
			height={1}
			aria-hidden
		>
			<defs>
				{[BLUE, RED, EXTERNAL].map((c) => (
					<marker
						key={c}
						id={`xv-arrow-${c.slice(1)}`}
						viewBox='0 0 10 10'
						refX='9'
						refY='5'
						markerWidth='7'
						markerHeight='7'
						orient='auto-start-reverse'
					>
						<path d='M0 0 L10 5 L0 10 z' fill={c} />
					</marker>
				))}
			</defs>
			{visible.map((a, i) => {
				const to = center(a.to.row, a.to.col);
				const fromCell =
					a.kind === 'precedent'
						? { row: a.from.top, col: a.from.left }
						: { row: a.from.top, col: a.from.left };
				if (a.external) {
					// Other-sheet end: a small sheet icon up and to the left, joined by a dashed line.
					const base = a.kind === 'precedent' ? to : center(fromCell.row, fromCell.col);
					const icon = { x: base.x - 46, y: base.y - 28 };
					const [x1, y1, x2, y2] =
						a.kind === 'precedent'
							? [icon.x + 8, icon.y + 6, to.x, to.y]
							: [base.x, base.y, icon.x + 8, icon.y + 6];
					return (
						<g key={i}>
							<rect x={icon.x} y={icon.y} width={16} height={12} fill='#ffffff' stroke={EXTERNAL} />
							<path
								d={`M${icon.x + 3} ${icon.y + 4} h10 M${icon.x + 3} ${icon.y + 8} h10`}
								stroke={EXTERNAL}
								strokeWidth={0.8}
							/>
							<line
								x1={x1}
								y1={y1}
								x2={x2}
								y2={y2}
								stroke={EXTERNAL}
								strokeWidth={1.2}
								strokeDasharray='4 3'
								markerEnd={`url(#xv-arrow-${EXTERNAL.slice(1)})`}
							/>
						</g>
					);
				}
				const from = center(fromCell.row, fromCell.col);
				const multi = a.from.top !== a.from.bottom || a.from.left !== a.from.right;
				const error = isCellError(sheet.cells.get(fromCell.row * 16384 + fromCell.col)?.value);
				const color = error ? RED : BLUE;
				const box = multi ? g.rangeRect(a.from) : null;
				return (
					<g key={i}>
						{box && (
							<rect
								x={box.x + 1}
								y={box.y + 1}
								width={Math.max(0, box.width - 2)}
								height={Math.max(0, box.height - 2)}
								fill='none'
								stroke={color}
								strokeWidth={2}
							/>
						)}
						<circle cx={from.x} cy={from.y} r={3} fill={color} />
						<line
							x1={from.x}
							y1={from.y}
							x2={to.x}
							y2={to.y}
							stroke={color}
							strokeWidth={1.5}
							markerEnd={`url(#xv-arrow-${color.slice(1)})`}
						/>
					</g>
				);
			})}
		</svg>
	);
});
