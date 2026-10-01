import { cellKey } from '../address/cell-address';
import { translateFormula } from '../formula/translate';
import { isDateFormat } from '../numfmt';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Workbook } from '../types/workbook';
import { dateToSerial, daysInMonth, serialToDateParts } from '../values/dates';
import type { EditTransaction } from './transaction';

const DAY_LISTS: string[][] = [
	['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
	['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
	[
		'January',
		'February',
		'March',
		'April',
		'May',
		'June',
		'July',
		'August',
		'September',
		'October',
		'November',
		'December',
	],
	['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
];

function matchCase(template: string, word: string): string {
	if (template === template.toUpperCase()) {
		return word.toUpperCase();
	}
	if (template === template.toLowerCase()) {
		return word.toLowerCase();
	}
	return word;
}

function listIndex(text: string): { list: string[]; index: number } | null {
	const lower = text.trim().toLowerCase();
	for (const list of DAY_LISTS) {
		const i = list.findIndex((w) => w.toLowerCase() === lower);
		if (i >= 0) {
			return { list, index: i };
		}
	}
	return null;
}

type Generator = (n: number) => Pick<Cell, 'value' | 'formula'>;

/**
 * Build a generator continuing the source values: `n` is the offset from the
 * first source cell (n >= k continues forward, n < 0 continues backward).
 */
function seriesFor(
	cells: Array<Cell | undefined>,
	isDate: boolean[],
	date1904: boolean,
	copyOnly: boolean,
): Generator {
	const k = cells.length;
	const repeat: Generator = (n) => {
		const src = cells[((n % k) + k) % k];
		return { value: src?.value ?? null, formula: src?.formula };
	};
	if (copyOnly || cells.some((c) => c?.formula !== undefined)) {
		return (n) => {
			const idx = ((n % k) + k) % k;
			const src = cells[idx];
			if (src?.formula !== undefined) {
				return { value: src.value, formula: undefined };
			}
			return repeat(n);
		};
	}
	const values = cells.map((c) => c?.value ?? null);
	if (values.every((v) => typeof v === 'number')) {
		const nums = values as number[];
		if (k === 1) {
			if (isDate[0]) {
				return (n) => ({ value: nums[0] + n });
			}
			return () => ({ value: nums[0] });
		}
		if (isDate.every(Boolean)) {
			// Month / year steps when the source dates share the same day of month.
			const parts = nums.map((v) => serialToDateParts(v, date1904));
			const sameDay = parts.every((p) => p.day === parts[0].day);
			const monthDiff = (parts[1].year - parts[0].year) * 12 + (parts[1].month - parts[0].month);
			const consistentMonths = parts.every(
				(p, i) =>
					i === 0 ||
					(p.year - parts[i - 1].year) * 12 + (p.month - parts[i - 1].month) === monthDiff,
			);
			if (sameDay && consistentMonths && monthDiff !== 0) {
				return (n) => {
					const total = parts[0].month - 1 + monthDiff * n;
					const year = parts[0].year + Math.floor(total / 12);
					const month = (((total % 12) + 12) % 12) + 1;
					const day = Math.min(parts[0].day, daysInMonth(year, month));
					return { value: dateToSerial(year, month, day, date1904) + (nums[0] % 1) };
				};
			}
		}
		// Linear trend (least squares), as Excel's AutoFill does for numbers.
		const xs = nums.map((_, i) => i);
		const mx = (k - 1) / 2;
		const my = nums.reduce((a, b) => a + b, 0) / k;
		let sxy = 0;
		let sxx = 0;
		for (let i = 0; i < k; i++) {
			sxy += (xs[i] - mx) * (nums[i] - my);
			sxx += (xs[i] - mx) ** 2;
		}
		const slope = sxx === 0 ? 0 : sxy / sxx;
		const intercept = my - slope * mx;
		return (n) => {
			if (n >= 0 && n < k) {
				return { value: nums[n] };
			}
			return { value: Number((intercept + slope * n).toPrecision(15)) };
		};
	}
	if (values.every((v) => typeof v === 'string')) {
		const strs = values as string[];
		const lists = strs.map(listIndex);
		if (lists.every((l) => l !== null) && lists.every((l) => l!.list === lists[0]!.list)) {
			const list = lists[0]!.list;
			const step = k > 1 ? lists[1]!.index - lists[0]!.index : 1;
			return (n) => {
				if (n >= 0 && n < k) {
					return { value: strs[n] };
				}
				const idx = (((lists[0]!.index + step * n) % list.length) + list.length) % list.length;
				return { value: matchCase(strs[0], list[idx]) };
			};
		}
		const m = strs.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s));
		if (m.every((x) => x !== null) && m.every((x) => x![1] === m[0]![1] && x![3] === m[0]![3])) {
			const prefix = m[0]![1];
			const suffix = m[0]![3];
			const width = m[0]![2].length;
			const nums = m.map((x) => Number(x![2]));
			const step = k > 1 ? nums[1] - nums[0] : 1;
			return (n) => {
				if (n >= 0 && n < k) {
					return { value: strs[n] };
				}
				const v = Math.abs(nums[0] + step * n);
				const digits = String(v);
				return {
					value: `${prefix}${digits.length < width ? digits.padStart(width, '0') : digits}${suffix}`,
				};
			};
		}
	}
	return repeat;
}

/**
 * AutoFill (fill handle / Ctrl+D / Ctrl+R): extend `source` into `target`
 * (which contains source and grows in one direction). Numbers and dates
 * continue as series, text with numbers increments, day / month names
 * cycle, formulas are copied with relative references adjusted, and
 * formats are copied cyclically. `copyOnly` repeats values without series.
 */
export function fillRange(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	source: RangeAddress,
	target: RangeAddress,
	options: { copyOnly?: boolean; formatsOnly?: boolean; valuesOnly?: boolean } = {},
): void {
	const sheet = workbook.sheets[sheetIndex];
	const vertical = target.top !== source.top || target.bottom !== source.bottom;
	const forward = vertical ? target.bottom > source.bottom : target.right > source.right;
	tx.touchRange(sheetIndex, target);
	const lines = vertical
		? Array.from({ length: source.right - source.left + 1 }, (_, i) => source.left + i)
		: Array.from({ length: source.bottom - source.top + 1 }, (_, i) => source.top + i);
	const k = vertical ? source.bottom - source.top + 1 : source.right - source.left + 1;
	for (const line of lines) {
		const srcCells: Array<Cell | undefined> = [];
		const srcPos: Array<[number, number]> = [];
		for (let i = 0; i < k; i++) {
			const row = vertical ? source.top + i : line;
			const col = vertical ? line : source.left + i;
			srcPos.push([row, col]);
			srcCells.push(sheet.cells.get(cellKey(row, col)));
		}
		const isDate = srcCells.map((c) => {
			const xf = workbook.styles.cellXfs[c?.style ?? 0];
			return (
				typeof c?.value === 'number' &&
				isDateFormat(numberFormatCode(xf?.numFmtId ?? 0, workbook.styles.numFmts))
			);
		});
		const gen = seriesFor(srcCells, isDate, workbook.date1904, options.copyOnly === true);
		const startOff = vertical
			? forward
				? source.bottom + 1
				: target.top
			: forward
				? source.right + 1
				: target.left;
		const endOff = vertical
			? forward
				? target.bottom
				: source.top - 1
			: forward
				? target.right
				: source.left - 1;
		for (let p = startOff; p <= endOff; p++) {
			const n = vertical ? p - source.top : p - source.left;
			const row = vertical ? p : line;
			const col = vertical ? line : p;
			const srcIdx = ((n % k) + k) % k;
			const src = srcCells[srcIdx];
			const key = cellKey(row, col);
			const existing = sheet.cells.get(key);
			const next: Cell = options.formatsOnly
				? { ...(existing ?? { value: null }) }
				: { value: null };
			if (!options.formatsOnly) {
				if (src?.formula !== undefined && !options.valuesOnly) {
					const [sr, sc] = srcPos[srcIdx];
					next.formula = translateFormula(src.formula, row - sr, col - sc);
					next.value = null;
					if (src.array?.dynamic) {
						next.array = { ref: { top: row, left: col, bottom: row, right: col }, dynamic: true };
					}
				} else {
					const g = gen(n);
					next.value = g.value;
				}
			}
			const style = options.valuesOnly ? existing?.style : src?.style;
			if (style !== undefined && style !== 0) {
				next.style = style;
			} else {
				delete next.style;
			}
			if (next.value === null && next.formula === undefined && next.style === undefined) {
				sheet.cells.delete(key);
			} else {
				sheet.cells.set(key, next);
			}
		}
	}
	sheet.dirty = true;
}
