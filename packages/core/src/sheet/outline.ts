import type { Worksheet } from '../types/sheet';

/** One outline group: a run of rows (or columns) at `level` or deeper, with its summary line. */
export interface OutlineGroup {
	level: number;
	start: number;
	end: number;
	/** Summary row / column holding the +/- button (below / right of the detail by default). */
	summary: number;
	/** The detail is hidden. */
	collapsed: boolean;
}

export interface AxisOutline {
	/** Deepest level used (0 = no outline). */
	levels: number;
	groups: OutlineGroup[];
}

const MAX_LEVEL = 7;

function lineInfo(
	sheet: Worksheet,
	axis: 'rows' | 'cols',
): Array<{ index: number; level: number; hidden: boolean; collapsed: boolean }> {
	const out: Array<{ index: number; level: number; hidden: boolean; collapsed: boolean }> = [];
	if (axis === 'rows') {
		for (const [index, info] of sheet.rows) {
			out.push({
				index,
				level: Math.min(MAX_LEVEL, info.outlineLevel ?? 0),
				hidden: !!info.hidden,
				collapsed: !!info.collapsed,
			});
		}
	} else {
		for (const c of sheet.cols) {
			for (let i = c.min; i <= Math.min(c.max, c.min + 16384); i++) {
				out.push({
					index: i,
					level: Math.min(MAX_LEVEL, c.outlineLevel ?? 0),
					hidden: !!c.hidden,
					collapsed: !!c.collapsed,
				});
			}
		}
	}
	return out.sort((a, b) => a.index - b.index);
}

/**
 * The outline groups of a sheet axis, computed from the per-row (per-column)
 * outline levels. A group at level L is a maximal run of consecutive lines
 * with level >= L; its summary line follows it (summaryBelow / summaryRight)
 * or precedes it.
 */
export function computeOutline(sheet: Worksheet, axis: 'rows' | 'cols'): AxisOutline {
	const lines = lineInfo(sheet, axis);
	const below = axis === 'rows' ? sheet.summaryBelow !== false : sheet.summaryRight !== false;
	const levelAt = new Map(lines.map((l) => [l.index, l]));
	let levels = 0;
	for (const l of lines) {
		levels = Math.max(levels, l.level);
	}
	const groups: OutlineGroup[] = [];
	for (let level = 1; level <= levels; level++) {
		let start = -1;
		let prev = -2;
		// Lines exactly at this level inside the current run, and whether all are hidden.
		let own = 0;
		let ownHidden = 0;
		const flush = (end: number) => {
			if (start < 0) {
				return;
			}
			const summary = below ? end + 1 : start - 1;
			const collapsed =
				(levelAt.get(summary)?.collapsed ?? false) || (own > 0 && own === ownHidden);
			groups.push({ level, start, end, summary, collapsed });
			start = -1;
		};
		for (const l of lines) {
			if (l.level >= level) {
				if (start < 0 || l.index !== prev + 1) {
					flush(prev);
					start = l.index;
					own = 0;
					ownHidden = 0;
				}
				if (l.level === level) {
					own++;
					ownHidden += l.hidden ? 1 : 0;
				}
				prev = l.index;
			}
		}
		flush(prev);
	}
	return { levels, groups };
}
