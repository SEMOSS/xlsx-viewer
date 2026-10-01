/**
 * Value-axis scaling the way Excel computes automatic minimum, maximum and
 * major unit:
 *
 *  - All-positive data starts at zero unless the smallest value exceeds 5/6 of
 *    the largest; then the minimum is `min - (max - min) / 2`. All-negative
 *    data mirrors this and ends at zero.
 *  - The far end gets 5% of the data span as headroom (none for 100% stacks).
 *  - The major unit is the smallest 1-2-5 step that labels the padded range
 *    with at most `maxIntervals` intervals (10 at most, fewer on short axes);
 *    the bounds are then rounded outwards to whole units.
 */

export interface ValueScale {
	min: number;
	max: number;
	majorUnit: number;
	minorUnit: number;
	/** Logarithmic base (log axes only). */
	logBase?: number;
	/** Major tick values from `min` to `max`. */
	ticks: number[];
	minorTicks: number[];
}

export interface ValueScaleOptions {
	/** Fixed bounds and units from the axis (`c:scaling`, `c:majorUnit`). */
	min?: number;
	max?: number;
	majorUnit?: number;
	minorUnit?: number;
	logBase?: number;
	/** Upper bound on major intervals (derived from the axis length); Excel never exceeds 10. */
	maxIntervals?: number;
	/** 100% stacked data: no headroom, bounds clamped to [-1, 1]. */
	percent?: boolean;
}

const MANTISSAS = [1, 2, 5];
const EPS = 1e-9;

/** The smallest 1-2-5 step that is at least `raw`. */
export function niceStep(raw: number): number {
	if (!Number.isFinite(raw) || raw <= 0) {
		return 1;
	}
	let exp = Math.floor(Math.log10(raw));
	for (;;) {
		const pow = 10 ** exp;
		for (const m of MANTISSAS) {
			if (m * pow >= raw * (1 - EPS)) {
				return m * pow;
			}
		}
		exp++;
	}
}

/** The next 1-2-5 step after `step`. */
function nextStep(step: number): number {
	return niceStep(step * 1.000001);
}

/** Round to the precision of `unit` to remove floating-point noise (0.1 + 0.2). */
export function roundToUnit(value: number, unit: number): number {
	const decimals = Math.max(0, Math.min(15, Math.ceil(-Math.log10(unit)) + 2));
	return Number(value.toFixed(decimals));
}

function floorTo(v: number, unit: number): number {
	return roundToUnit(Math.floor(v / unit + EPS) * unit, unit);
}

function ceilTo(v: number, unit: number): number {
	return roundToUnit(Math.ceil(v / unit - EPS) * unit, unit);
}

function ticksBetween(min: number, max: number, unit: number, limit = 1000): number[] {
	const out: number[] = [];
	if (!(unit > 0) || max < min) {
		return [min];
	}
	const start = Math.ceil(min / unit - EPS);
	for (let k = start; out.length < limit; k++) {
		const v = roundToUnit(k * unit, unit);
		if (v > max + Math.abs(unit) * EPS) {
			break;
		}
		out.push(v);
	}
	return out;
}

function logScale(dataMin: number, dataMax: number, opts: ValueScaleOptions): ValueScale {
	const base = opts.logBase && opts.logBase > 1 ? opts.logBase : 10;
	const lg = (v: number) => Math.log(v) / Math.log(base);
	const positiveMin = dataMin > 0 ? dataMin : dataMax > 0 ? dataMax / base : 1;
	const min = opts.min && opts.min > 0 ? opts.min : base ** Math.floor(lg(positiveMin) + EPS);
	let max =
		opts.max && opts.max > min ? opts.max : base ** Math.ceil(lg(Math.max(dataMax, min)) - EPS);
	if (max <= min) {
		max = min * base;
	}
	const unit = opts.majorUnit && opts.majorUnit > 1 ? opts.majorUnit : base;
	const ticks: number[] = [];
	for (let v = min; v <= max * (1 + EPS) && ticks.length < 400; v *= unit) {
		ticks.push(v);
	}
	const minorTicks: number[] = [];
	for (const t of ticks) {
		for (let k = 2; k < base && minorTicks.length < 4000; k++) {
			const v = t * k;
			if (v < max) {
				minorTicks.push(v);
			}
		}
	}
	return { min, max, majorUnit: unit, minorUnit: unit, logBase: base, ticks, minorTicks };
}

/** Excel's automatic bounds before rounding to the major unit. */
function paddedBounds(low: number, high: number, percent: boolean): [number, number] {
	if (percent) {
		return [low < 0 ? Math.max(-1, low) : 0, high > 0 ? Math.min(1, high) : 0];
	}
	if (low === high) {
		if (low === 0) {
			return [0, 1];
		}
		return low > 0 ? [0, high * 1.05] : [low * 1.05, 0];
	}
	const span = high - low;
	if (low >= 0) {
		const min = low > (high * 5) / 6 ? low - span / 2 : 0;
		return [min, high + span * 0.05];
	}
	if (high <= 0) {
		const max = high < (low * 5) / 6 ? high + span / 2 : 0;
		return [low - span * 0.05, max];
	}
	return [low - span * 0.05, high + span * 0.05];
}

/** Compute an automatic (or partially fixed) linear or logarithmic value scale. */
export function computeValueScale(
	dataMin: number,
	dataMax: number,
	opts: ValueScaleOptions = {},
): ValueScale {
	let low = Number.isFinite(dataMin) ? dataMin : 0;
	let high = Number.isFinite(dataMax) ? dataMax : 0;
	if (low > high) {
		[low, high] = [high, low];
	}
	if (opts.logBase !== undefined) {
		return logScale(low, high, opts);
	}
	const maxIntervals = Math.max(1, Math.min(10, Math.floor(opts.maxIntervals ?? 10)));
	let [autoMin, autoMax] = paddedBounds(low, high, Boolean(opts.percent));
	// A fixed bound on one side only shifts the automatic range.
	if (opts.min !== undefined) {
		autoMin = opts.min;
		if (opts.max === undefined && autoMax <= autoMin) {
			autoMax = autoMin + Math.max(1, Math.abs(autoMin) * 0.1);
		}
	}
	if (opts.max !== undefined) {
		autoMax = opts.max;
		if (opts.min === undefined && autoMin >= autoMax) {
			autoMin = autoMax - Math.max(1, Math.abs(autoMax) * 0.1);
		}
	}
	if (autoMax <= autoMin) {
		autoMax = autoMin + 1;
	}
	let unit = opts.majorUnit && opts.majorUnit > 0 ? opts.majorUnit : 0;
	if (!unit) {
		unit = niceStep((autoMax - autoMin) / maxIntervals);
		for (let guard = 0; guard < 40; guard++) {
			const lo = opts.min !== undefined || autoMin === 0 ? autoMin : floorTo(autoMin, unit);
			const hi = opts.max !== undefined || autoMax === 0 ? autoMax : ceilTo(autoMax, unit);
			if ((hi - lo) / unit <= maxIntervals + EPS) {
				break;
			}
			unit = nextStep(unit);
		}
	}
	const min = opts.min !== undefined || autoMin === 0 ? autoMin : floorTo(autoMin, unit);
	let max = opts.max !== undefined || autoMax === 0 ? autoMax : ceilTo(autoMax, unit);
	if (max <= min) {
		max = min + unit;
	}
	const minorUnit = opts.minorUnit && opts.minorUnit > 0 ? opts.minorUnit : unit / 5;
	const ticks = ticksBetween(min, max, unit);
	const minorTicks = ticksBetween(min, max, minorUnit, 5000).filter(
		(v) => Math.abs(v / unit - Math.round(v / unit)) > 1e-6,
	);
	return { min, max, majorUnit: unit, minorUnit, ticks, minorTicks };
}

/** Map a value to [0, 1] along a scale (log-aware), unclamped. */
export function scaleFraction(scale: ValueScale, value: number): number {
	if (scale.logBase) {
		if (value <= 0) {
			return -Infinity;
		}
		const lg = (v: number) => Math.log(v);
		return (lg(value) - lg(scale.min)) / (lg(scale.max) - lg(scale.min));
	}
	return (value - scale.min) / (scale.max - scale.min);
}

/**
 * How many major intervals fit on an axis: Excel keeps tick labels at least
 * ~1.5 line heights apart on vertical axes and one label width plus a gap
 * apart on horizontal axes, and never uses more than 10 intervals.
 */
export function maxIntervalsFor(axisLength: number, labelExtent: number): number {
	if (!(axisLength > 0) || !(labelExtent > 0)) {
		return 10;
	}
	return Math.max(1, Math.min(10, Math.floor(axisLength / labelExtent)));
}
