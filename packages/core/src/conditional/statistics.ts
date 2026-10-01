import type { RangeAddress } from '../types/address';
import type { CellValue } from '../types/cell';
import type { Worksheet } from '../types/sheet';
import { forEachCellInRanges } from './range-cells';

/**
 * Identity of a value for duplicate / unique rules: text compares
 * case-insensitively, numbers by value (15 significant digits, like Excel),
 * and different types never match. Blanks and errors have no key (ignored).
 */
export function duplicateKey(value: CellValue): string | undefined {
	switch (typeof value) {
		case 'number':
			return Number.isFinite(value) ? `n:${Number(value.toPrecision(15))}` : undefined;
		case 'string':
			return `s:${value.toLowerCase()}`;
		case 'boolean':
			return value ? 'b:1' : 'b:0';
		default:
			return undefined;
	}
}

/**
 * Lazily computed statistics over the cells of a conditional format's
 * ranges. Only numbers take part in the numeric statistics (text, booleans,
 * errors and blanks are ignored, like Excel's color scales, data bars, icon
 * sets, top/bottom and average rules). Each statistic is computed once on
 * first use; create a new instance (or drop the owner) to refresh.
 */
export class RangeStatistics {
	private sorted: Float64Array | undefined;
	private total = 0;
	private sampleStdDev: number | undefined;
	private counts: Map<string, number> | undefined;

	constructor(
		private readonly sheet: Worksheet,
		private readonly ranges: readonly RangeAddress[],
	) {}

	/** Numeric values, sorted ascending. */
	get values(): Float64Array {
		if (!this.sorted) {
			const out: number[] = [];
			forEachCellInRanges(this.sheet, this.ranges, (v) => {
				if (typeof v === 'number' && Number.isFinite(v)) {
					out.push(v);
				}
			});
			const arr = Float64Array.from(out);
			arr.sort();
			let sum = 0;
			for (let i = 0; i < arr.length; i++) {
				sum += arr[i];
			}
			this.total = sum;
			this.sorted = arr;
		}
		return this.sorted;
	}

	get count(): number {
		return this.values.length;
	}

	get min(): number | undefined {
		const v = this.values;
		return v.length ? v[0] : undefined;
	}

	get max(): number | undefined {
		const v = this.values;
		return v.length ? v[v.length - 1] : undefined;
	}

	get sum(): number {
		void this.values;
		return this.total;
	}

	get average(): number | undefined {
		const n = this.count;
		return n ? this.total / n : undefined;
	}

	/** Sample standard deviation (STDEV.S); 0 for a single value, `undefined` when empty. */
	get stdDev(): number | undefined {
		const v = this.values;
		if (v.length === 0) {
			return undefined;
		}
		if (this.sampleStdDev === undefined) {
			if (v.length === 1) {
				this.sampleStdDev = 0;
			} else {
				const mean = this.total / v.length;
				let sq = 0;
				for (let i = 0; i < v.length; i++) {
					const d = v[i] - mean;
					sq += d * d;
				}
				this.sampleStdDev = Math.sqrt(sq / (v.length - 1));
			}
		}
		return this.sampleStdDev;
	}

	/** PERCENTILE.INC with `p` in [0, 1] (clamped). */
	percentile(p: number): number | undefined {
		const v = this.values;
		if (v.length === 0 || !Number.isFinite(p)) {
			return undefined;
		}
		const rank = Math.min(1, Math.max(0, p)) * (v.length - 1);
		const lo = Math.floor(rank);
		const frac = rank - lo;
		return lo + 1 < v.length ? v[lo] + frac * (v[lo + 1] - v[lo]) : v[lo];
	}

	/** k-th largest value (1-based, like LARGE), or `undefined` when out of range. */
	large(k: number): number | undefined {
		const v = this.values;
		return k >= 1 && k <= v.length ? v[v.length - Math.floor(k)] : undefined;
	}

	/** k-th smallest value (1-based, like SMALL), or `undefined` when out of range. */
	small(k: number): number | undefined {
		const v = this.values;
		return k >= 1 && k <= v.length ? v[Math.floor(k) - 1] : undefined;
	}

	/** Number of cells whose {@link duplicateKey} equals `key`. */
	occurrences(key: string): number {
		if (!this.counts) {
			const counts = new Map<string, number>();
			forEachCellInRanges(this.sheet, this.ranges, (v) => {
				const k = duplicateKey(v);
				if (k !== undefined) {
					counts.set(k, (counts.get(k) ?? 0) + 1);
				}
			});
			this.counts = counts;
		}
		return this.counts.get(key) ?? 0;
	}
}
