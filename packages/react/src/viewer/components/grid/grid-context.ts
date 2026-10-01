import { createContext, useContext } from 'react';

import type { ColorContext, DifferentialFormat, StyleResolver, Worksheet } from 'xlsx-viewer-core';

import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { CfOverlay, Span } from '../../../engine/layout/types';
import type { PaneLayout } from '../../../engine/layout/panes';
import type { MergeIndex } from '../../../engine/layout/style-at';

/** Values shared by every pane of the grid for one render. */
export interface GridRenderContext {
	sheet: Worksheet;
	sheetIndex: number;
	geometry: SheetGeometry;
	panes: PaneLayout;
	merges: MergeIndex;
	styles: StyleResolver;
	colorCtx: ColorContext;
	zoom: number;
	docVersion: number;
	date1904: boolean;
	showGridLines: boolean;
	showFormulas: boolean;
	showZeros: boolean;
	gridlineColor: string;
	cf?: (row: number, col: number) => CfOverlay | undefined;
	table?: (row: number, col: number) => DifferentialFormat | undefined;
	/** Visible body window (with overscan). */
	bodyRows: Span;
	bodyCols: Span;
}

export const GridRenderContextReact = createContext<GridRenderContext | null>(null);

export function useGridRender(): GridRenderContext {
	const ctx = useContext(GridRenderContextReact);
	if (!ctx) {
		throw new Error('Grid render context missing');
	}
	return ctx;
}
