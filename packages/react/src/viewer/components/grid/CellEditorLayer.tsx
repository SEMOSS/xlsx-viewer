import type { CSSProperties } from 'react';

import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { MergeIndex } from '../../../engine/layout/style-at';
import type { EditState } from '../../state/types';
import { CellEditor } from './CellEditor';

interface CellEditorLayerProps {
	edit: EditState;
	geometry: SheetGeometry;
	merges: MergeIndex;
	panes: { bodyRow: number; bodyCol: number };
	header: { headerWidth: number; headerHeight: number };
	/** Sheet coordinates at the top-left of the frozen and body panes. */
	origins: { frozenX: number; frozenY: number; bodyX: number; bodyY: number };
	maxRight: number;
	onDone: () => void;
}

/** Grid placement of the editor layer: the same area and sticky offsets as the pane holding the cell. */
function layerStyle(
	edit: { row: number; col: number },
	panes: CellEditorLayerProps['panes'],
	header: CellEditorLayerProps['header'],
): CSSProperties {
	const frozenRow = edit.row < panes.bodyRow;
	const frozenCol = edit.col < panes.bodyCol;
	return {
		gridArea: `${frozenRow ? 2 : 3} / ${frozenCol ? 2 : 3}`,
		position: frozenRow || frozenCol ? 'sticky' : 'relative',
		top: frozenRow ? header.headerHeight : undefined,
		left: frozenCol ? header.headerWidth : undefined,
		width: 0,
		height: 0,
		zIndex: 7,
	};
}

/**
 * The in-cell editor in an unclipped layer that scrolls like the pane
 * holding the cell, so it can grow across the freeze lines (and cover
 * merges spanning them).
 */
export function CellEditorLayer({
	edit,
	geometry,
	merges,
	panes,
	header,
	origins,
	maxRight,
	onDone,
}: CellEditorLayerProps) {
	const x = edit.col < panes.bodyCol ? origins.frozenX : origins.bodyX;
	const y = edit.row < panes.bodyRow ? origins.frozenY : origins.bodyY;
	return (
		<div className='xv-editor-layer' style={layerStyle(edit, panes, header)}>
			<div className='xv-pane-inner' style={{ transform: `translate(${-x}px, ${-y}px)` }}>
				<CellEditor
					edit={edit}
					geometry={geometry}
					merges={merges}
					maxRight={maxRight}
					onDone={onDone}
				/>
			</div>
		</div>
	);
}
