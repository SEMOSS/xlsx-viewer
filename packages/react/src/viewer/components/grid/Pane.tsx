import { memo, type CSSProperties, type ReactNode } from 'react';

import type { Span } from '../../../engine/layout/types';
import { PaneLayers } from './PaneLayers';

interface PaneProps {
	rows: Span;
	cols: Span;
	/** Sheet coordinates shown at the pane's top-left corner. */
	originX: number;
	originY: number;
	style: CSSProperties;
	className?: string;
	/** Draw the freeze-pane separator on the bottom / right edge. */
	freezeBottom?: boolean;
	freezeRight?: boolean;
	children?: ReactNode;
}

/**
 * One rectangular region of the grid (body, frozen rows, frozen columns or
 * the frozen corner). Content is laid out in sheet coordinates inside an
 * inner box shifted by the pane origin; the pane clips it.
 */
export const Pane = memo(function Pane({
	rows,
	cols,
	originX,
	originY,
	style,
	className,
	freezeBottom,
	freezeRight,
	children,
}: PaneProps) {
	const empty = rows.end < rows.start || cols.end < cols.start;
	return (
		<div
			className={`xv-pane ${className ?? ''}`}
			style={style}
			data-freeze-bottom={freezeBottom || undefined}
			data-freeze-right={freezeRight || undefined}
		>
			<div
				className='xv-pane-inner'
				style={{ transform: `translate(${-originX}px, ${-originY}px)` }}
			>
				{!empty && <PaneLayers rows={rows} cols={cols} />}
				{children}
			</div>
		</div>
	);
});
