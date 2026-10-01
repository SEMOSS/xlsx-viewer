import { parseCellAddress } from '../address/cell-address';
import { parseSqref } from '../address/range';
import { createDefaultSheetView } from '../sheet/worksheet';
import type { PaneName, SheetPane, SheetSelection, SheetView } from '../types/sheet';
import type { XmlElement } from '../xml/types';
import { bool, int, kid, kids, num, oneOf, str } from './xml-utils';

const PANES: readonly PaneName[] = ['bottomRight', 'topRight', 'bottomLeft', 'topLeft'];

function parsePane(el: XmlElement | undefined): SheetPane | undefined {
	if (!el) {
		return undefined;
	}
	const pane: SheetPane = {
		xSplit: num(el, 'xSplit') ?? 0,
		ySplit: num(el, 'ySplit') ?? 0,
		state: oneOf(str(el, 'state'), ['frozen', 'frozenSplit', 'split'] as const) ?? 'split',
	};
	const topLeft = str(el, 'topLeftCell');
	const tl = topLeft ? parseCellAddress(topLeft) : null;
	if (tl) {
		pane.topLeftCell = tl;
	}
	const active = oneOf(str(el, 'activePane'), PANES);
	if (active) {
		pane.activePane = active;
	}
	return pane;
}

function parseSelection(el: XmlElement): SheetSelection {
	const sqref = str(el, 'sqref');
	const ranges = sqref ? parseSqref(sqref) : [];
	const active = str(el, 'activeCell');
	const activeCell =
		(active ? parseCellAddress(active) : null) ??
		(ranges[0] ? { row: ranges[0].top, col: ranges[0].left } : { row: 0, col: 0 });
	const selection: SheetSelection = {
		activeCell,
		ranges:
			ranges.length > 0
				? ranges
				: [
						{
							top: activeCell.row,
							left: activeCell.col,
							bottom: activeCell.row,
							right: activeCell.col,
						},
					],
	};
	const pane = oneOf(str(el, 'pane'), PANES);
	if (pane) {
		selection.pane = pane;
	}
	return selection;
}

/**
 * The first `<sheetView>` of `<sheetViews>`. When there are several
 * selections (one per pane), the one for the active pane is used.
 */
export function parseSheetView(sheetViews: XmlElement | undefined): SheetView {
	const view = createDefaultSheetView();
	const el = kid(sheetViews, 'sheetView');
	if (!el) {
		return view;
	}
	view.showGridLines = bool(el, 'showGridLines') ?? true;
	view.showRowColHeaders = bool(el, 'showRowColHeaders') ?? true;
	view.showZeros = bool(el, 'showZeros') ?? true;
	view.showFormulas = bool(el, 'showFormulas') ?? false;
	view.rightToLeft = bool(el, 'rightToLeft') ?? false;
	const zoom = int(el, 'zoomScale');
	if (zoom !== undefined && zoom > 0) {
		view.zoomScale = Math.max(10, Math.min(400, zoom));
	}
	const tabSelected = bool(el, 'tabSelected');
	if (tabSelected !== undefined) {
		view.tabSelected = tabSelected;
	}
	const mode = oneOf(str(el, 'view'), ['normal', 'pageBreakPreview', 'pageLayout'] as const);
	if (mode) {
		view.view = mode;
	}
	const topLeft = str(el, 'topLeftCell');
	const tl = topLeft ? parseCellAddress(topLeft) : null;
	if (tl) {
		view.topLeftCell = tl;
	}
	if (bool(el, 'defaultGridColor') === false) {
		view.gridColor = { indexed: int(el, 'colorId') ?? 64 };
	}
	const pane = parsePane(kid(el, 'pane'));
	if (pane) {
		view.pane = pane;
	}
	const selections = kids(el, 'selection');
	if (selections.length > 0) {
		const activePane = pane?.activePane ?? 'topLeft';
		const chosen =
			selections.find((s) => (str(s, 'pane') ?? 'topLeft') === activePane) ?? selections[0];
		view.selection = parseSelection(chosen);
	}
	return view;
}
