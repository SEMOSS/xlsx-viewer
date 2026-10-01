import type { AxisSize } from './metrics';
import type { PrintSpan } from './types';

/**
 * Page breaks along one axis (rows or columns): pages are filled greedily
 * with whole rows (columns), a manual break always starts a new page, and
 * a row taller than the page gets a page of its own (and is clipped).
 */

const EPSILON = 1e-6;

export interface AxisInput {
	/** First and last row (column) of the area, inclusive. */
	span: PrintSpan;
	/** Size in points at 100% (0 for hidden lines). */
	size: AxisSize;
	/**
	 * Room for body lines on a page whose body starts at `start`, in points
	 * at the given scale (titles and headings already taken off).
	 */
	capacity: (start: number, scale: number) => number;
	/** Lines that start a new page (manual breaks). */
	manual?: ReadonlySet<number>;
}

/** Repeated title lines printed above (left of) a page starting at `start`. */
export function titlesBefore(titles: PrintSpan | undefined, start: number): PrintSpan | undefined {
	if (!titles || titles.start >= start) {
		return undefined;
	}
	return { start: titles.start, end: Math.min(titles.end, start - 1) };
}

/** Bands of lines, one per page, for a scale (1 = 100%). Bands of hidden lines only are dropped. */
export function paginateAxis(input: AxisInput, scale: number): PrintSpan[] {
	const { span, size, manual } = input;
	const out: PrintSpan[] = [];
	let i = span.start;
	while (i <= span.end) {
		const cap = input.capacity(i, scale);
		let used = 0;
		let visible = false;
		let j = i;
		while (j <= span.end) {
			if (j > i && manual?.has(j)) {
				break;
			}
			const s = size(j) * scale;
			if (visible && used + s > cap + EPSILON) {
				break;
			}
			used += s;
			visible ||= s > 0;
			j++;
		}
		if (visible) {
			out.push({ start: i, end: j - 1 });
		} else if (out.length > 0) {
			// Trailing hidden lines belong to the previous page.
			out[out.length - 1].end = j - 1;
		}
		i = j;
	}
	return out;
}

/** Total size of the area in points at 100%. */
export function axisTotal(input: AxisInput): number {
	let total = 0;
	for (let i = input.span.start; i <= input.span.end; i++) {
		total += input.size(i);
	}
	return total;
}

/**
 * Largest whole percentage (10 to 100) at which the area takes at most
 * `pages` pages along this axis, as Fit to N pages does. `printable` is the
 * page length and `fixed` the repeated titles plus headings, both at 100%
 * (for the first estimate). Returns 0.1 when even 10% does not fit.
 */
export function fitAxisScale(
	input: AxisInput,
	pages: number,
	printable: number,
	fixed: number,
): number {
	const total = axisTotal(input);
	if (pages <= 0 || total <= 0) {
		return 1;
	}
	const count = (percent: number) => paginateAxis(input, percent / 100).length;
	const estimate = Math.min(1, (pages * Math.max(1, printable)) / (total + pages * fixed));
	let p = Math.max(10, Math.min(100, Math.floor(estimate * 100 + EPSILON)));
	while (p > 10 && count(p) > pages) {
		p--;
	}
	while (p < 100 && count(p + 1) <= pages) {
		p++;
	}
	return p / 100;
}
