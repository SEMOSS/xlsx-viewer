import type { ResolvedCellStyle } from 'xlsx-viewer-core';
import { classifyFormat } from 'xlsx-viewer-core';

import { styleIdAt } from '../../engine/layout/style-at';
import { useStoreSelector } from '../state/use-store';

export interface ActiveFormat {
	style: ResolvedCellStyle;
	category: ReturnType<typeof classifyFormat>;
	merged: boolean;
}

function same(a: ActiveFormat, b: ActiveFormat): boolean {
	return a.style === b.style && a.merged === b.merged;
}

/** Formatting of the active cell, for the ribbon's toggle states. */
export function useActiveFormat(): ActiveFormat {
	return useStoreSelector((snap, store) => {
		void snap.docVersion;
		const sheet = store.workbook.sheets[snap.ui.activeSheet];
		const sel = store.selectionOf(snap.ui.activeSheet);
		const { row, col } = sel.active;
		const style = store.styles.resolve(styleIdAt(sheet, row, col));
		return {
			style,
			category: classifyFormat(style.numFmt),
			merged: sheet.merges.some((m) => m.top === row && m.left === col),
		};
	}, same);
}
