/**
 * Date (time-scaled) category axes: points sit at their date, and ticks fall
 * on whole days, months or years like Excel's date axis.
 */
import { dateToSerial, serialToDateParts } from '../../values/dates';

export type TimeUnit = 'days' | 'months' | 'years';

export interface DateScale {
	/** Domain in date serials (padded by half a base unit for bars). */
	min: number;
	max: number;
	baseUnit: TimeUnit;
	/** Width of one base-unit slot in days (bar widths). */
	slotDays: number;
	/** Major tick serials. */
	ticks: number[];
}

export interface DateScaleOptions {
	between: boolean;
	baseUnit?: TimeUnit;
	majorUnit?: number;
	majorTimeUnit?: TimeUnit;
	min?: number;
	max?: number;
	maxIntervals?: number;
	date1904?: boolean;
}

const DAY_STEPS = [1, 2, 3, 7, 14];
const MONTH_STEPS = [1, 2, 3, 6, 12];
const YEAR_STEPS = [1, 2, 5, 10, 20, 50, 100];

/** Base unit Excel would pick: the coarsest unit every date sits on. */
export function detectBaseUnit(dates: readonly number[], date1904 = false): TimeUnit {
	const valid = dates.filter((d) => Number.isFinite(d));
	if (valid.length < 2) {
		return 'days';
	}
	const parts = valid.map((d) => serialToDateParts(Math.floor(d), date1904));
	const sameDay = parts.every((p) => p.day === parts[0].day);
	const sameMonth = parts.every((p) => p.month === parts[0].month);
	if (sameDay && sameMonth) {
		return 'years';
	}
	if (sameDay) {
		return 'months';
	}
	return 'days';
}

function addUnits(serial: number, unit: TimeUnit, n: number, date1904: boolean): number {
	if (unit === 'days') {
		return serial + n;
	}
	const p = serialToDateParts(Math.floor(serial), date1904);
	if (unit === 'months') {
		return dateToSerial(p.year, p.month + n, p.day, date1904);
	}
	return dateToSerial(p.year + n, p.month, p.day, date1904);
}

function unitDays(unit: TimeUnit): number {
	return unit === 'days' ? 1 : unit === 'months' ? 30.4375 : 365.25;
}

export function computeDateScale(dates: readonly number[], opts: DateScaleOptions): DateScale {
	const date1904 = Boolean(opts.date1904);
	const valid = dates.filter((d) => Number.isFinite(d)).map((d) => Math.floor(d));
	const lo = opts.min ?? (valid.length ? Math.min(...valid) : 0);
	const hi = opts.max ?? (valid.length ? Math.max(...valid) : lo + 1);
	const baseUnit = opts.baseUnit ?? detectBaseUnit(valid, date1904);
	const slotDays = unitDays(baseUnit);
	const pad = opts.between ? slotDays / 2 : 0;
	const min = lo - (opts.min === undefined ? pad : 0);
	const max = Math.max(hi + (opts.max === undefined ? pad : 0), min + (pad ? slotDays : 1));
	const maxIntervals = Math.max(1, opts.maxIntervals ?? 10);
	const majorUnitKind = opts.majorTimeUnit ?? baseUnit;
	let step = opts.majorUnit && opts.majorUnit > 0 ? opts.majorUnit : 0;
	if (!step) {
		const spanUnits = (hi - lo) / unitDays(majorUnitKind);
		const steps =
			majorUnitKind === 'days' ? DAY_STEPS : majorUnitKind === 'months' ? MONTH_STEPS : YEAR_STEPS;
		step = steps.find((s) => spanUnits / s <= maxIntervals) ?? Math.ceil(spanUnits / maxIntervals);
	}
	const ticks: number[] = [];
	for (
		let t = lo;
		t <= hi + 1e-9 && ticks.length < 1000;
		t = addUnits(t, majorUnitKind, step, date1904)
	) {
		ticks.push(t);
	}
	return { min, max, baseUnit, slotDays, ticks };
}

/** Position of a serial in [0, 1] along the scale. */
export function dateFraction(scale: DateScale, serial: number): number {
	return (serial - scale.min) / (scale.max - scale.min || 1);
}
