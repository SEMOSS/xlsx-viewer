import type { RangeAddress } from 'xlsx-viewer-core';

import type { TraceArrow } from '../state/types';
import type { ViewerStore } from '../state/viewer-store';

const key = (a: TraceArrow) =>
	`${a.kind}:${a.sheet}:${a.from.top}:${a.from.left}:${a.from.bottom}:${a.from.right}>${a.to.row}:${a.to.col}`;

function addArrows(store: ViewerStore, arrows: TraceArrow[]): boolean {
	const seen = new Set(store.state.traceArrows.map(key));
	const fresh = arrows.filter((a) => {
		const k = key(a);
		if (seen.has(k)) {
			return false;
		}
		seen.add(k);
		return true;
	});
	if (fresh.length) {
		store.setUi({ traceArrows: [...store.state.traceArrows, ...fresh] });
	}
	return fresh.length > 0;
}

const cellRange = (row: number, col: number): RangeAddress => ({
	top: row,
	left: col,
	bottom: row,
	right: col,
});

/**
 * Formulas > Trace Precedents: arrows from the cells a formula reads to the
 * formula. Each further click adds the next level, as in Excel.
 */
export function tracePrecedents(store: ViewerStore): void {
	const sheet = store.state.activeSheet;
	const { row, col } = store.selectionOf().active;
	// Next level: the precedents of every cell already at the arrows' source end.
	const frontier = store.state.traceArrows.filter(
		(a) => a.kind === 'precedent' && a.sheet === sheet && !a.external,
	);
	const targets = frontier.flatMap((a) =>
		a.from.top === a.from.bottom && a.from.left === a.from.right
			? [{ row: a.from.top, col: a.from.left }]
			: [],
	);
	const arrows: TraceArrow[] = [];
	for (const t of [{ row, col }, ...targets]) {
		for (const p of store.calc.getPrecedents({ sheet, row: t.row, col: t.col })) {
			arrows.push({
				sheet,
				from: p.range,
				to: t,
				kind: 'precedent',
				external: p.sheet !== sheet || undefined,
			});
		}
	}
	if (
		!addArrows(store, arrows) &&
		arrows.length === 0 &&
		!store.activeSheet.cells.get(row * 16384 + col)?.formula
	) {
		store.setUi({ message: store.t('message.noPrecedents') });
	}
}

/** Formulas > Trace Dependents: arrows from the cell to the formulas that read it. */
export function traceDependents(store: ViewerStore): void {
	const sheet = store.state.activeSheet;
	const { row, col } = store.selectionOf().active;
	const frontier = store.state.traceArrows
		.filter((a) => a.kind === 'dependent' && a.sheet === sheet && !a.external)
		.map((a) => a.to);
	const sources = frontier.length ? [{ row, col }, ...frontier] : [{ row, col }];
	const arrows: TraceArrow[] = [];
	for (const s of sources) {
		for (const d of store.calc.getDependents({ sheet, row: s.row, col: s.col })) {
			arrows.push({
				sheet,
				from: cellRange(s.row, s.col),
				to: { row: d.row, col: d.col },
				kind: 'dependent',
				external: d.sheet !== sheet || undefined,
			});
		}
	}
	if (!addArrows(store, arrows) && arrows.length === 0) {
		store.setUi({ message: store.t('message.noDependents') });
	}
}

/** Formulas > Remove Arrows (all, or only precedent / dependent ones). */
export function removeArrows(store: ViewerStore, kind?: TraceArrow['kind']): void {
	store.setUi({ traceArrows: kind ? store.state.traceArrows.filter((a) => a.kind !== kind) : [] });
}
