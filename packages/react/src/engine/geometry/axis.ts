/**
 * One axis (rows or columns) of a sheet: a default size plus sparse
 * overrides, with O(log n) offset and hit-testing via prefix sums.
 * Sizes are in zoomed CSS pixels; hidden lines have size 0.
 */
export class AxisMetrics {
	/** Sorted indices that differ from the default size. */
	private readonly keys: number[];
	private readonly sizes: number[];
	/** prefixDelta[i] = sum of (size - default) for keys[0..i-1]. */
	private readonly prefixDelta: number[];

	constructor(
		readonly defaultSize: number,
		overrides: ReadonlyMap<number, number>,
		readonly count: number,
	) {
		const entries = [...overrides.entries()]
			.filter(([k, v]) => k >= 0 && k < count && v !== defaultSize)
			.sort((a, b) => a[0] - b[0]);
		this.keys = entries.map((e) => e[0]);
		this.sizes = entries.map((e) => e[1]);
		this.prefixDelta = new Array(entries.length + 1);
		this.prefixDelta[0] = 0;
		for (let i = 0; i < entries.length; i++) {
			this.prefixDelta[i + 1] = this.prefixDelta[i] + (this.sizes[i] - defaultSize);
		}
	}

	/** Index of the first override >= index (binary search). */
	private lowerBound(index: number): number {
		let lo = 0;
		let hi = this.keys.length;
		while (lo < hi) {
			const mid = (lo + hi) >> 1;
			if (this.keys[mid] < index) {
				lo = mid + 1;
			} else {
				hi = mid;
			}
		}
		return lo;
	}

	size(index: number): number {
		const i = this.lowerBound(index);
		return i < this.keys.length && this.keys[i] === index ? this.sizes[i] : this.defaultSize;
	}

	/** Offset of the start of `index` from the start of the axis. */
	offset(index: number): number {
		const i = this.lowerBound(index);
		return index * this.defaultSize + this.prefixDelta[i];
	}

	/** Offset of the end of `index`. */
	end(index: number): number {
		return this.offset(index) + this.size(index);
	}

	/** Total length of the first `n` lines (default: all). */
	total(n = this.count): number {
		return this.offset(n);
	}

	/**
	 * Index of the line containing `pos`. Positions past the end clamp to the
	 * last line. Zero-size (hidden) lines are never returned.
	 */
	indexAt(pos: number): number {
		if (pos <= 0) {
			return this.nextVisible(0);
		}
		let lo = 0;
		let hi = this.count - 1;
		while (lo < hi) {
			const mid = (lo + hi + 1) >> 1;
			if (this.offset(mid) <= pos) {
				lo = mid;
			} else {
				hi = mid - 1;
			}
		}
		// Step back over hidden lines that share the same offset.
		let idx = lo;
		while (idx > 0 && this.size(idx) === 0) {
			idx--;
		}
		if (this.size(idx) === 0) {
			return this.nextVisible(idx);
		}
		return idx;
	}

	isHidden(index: number): boolean {
		return this.size(index) === 0;
	}

	/** First visible line at or after `index` (or `index` when none). */
	nextVisible(index: number): number {
		let i = index;
		while (i < this.count - 1 && this.size(i) === 0) {
			i++;
		}
		return i;
	}

	/** Last visible line at or before `index` (or `index` when none). */
	prevVisible(index: number): number {
		let i = index;
		while (i > 0 && this.size(i) === 0) {
			i--;
		}
		return i;
	}
}
