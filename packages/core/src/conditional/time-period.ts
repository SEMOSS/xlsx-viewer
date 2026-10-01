import type { CfTimePeriod } from '../types/conditional-format';
import { serialToDateParts } from '../values/dates';

/**
 * "A Date Occurring" test, equivalent to the formulas Excel writes for
 * `timePeriod` rules (`FLOOR(A1,1)=TODAY()`, ...). `today` is a date serial;
 * weeks run Sunday to Saturday; months compare calendar month and year.
 */
export function matchesTimePeriod(
	value: number,
	period: CfTimePeriod,
	today: number,
	date1904 = false,
): boolean {
	if (!Number.isFinite(value) || value < 0) {
		return false;
	}
	const day = Math.floor(value);
	const t = Math.floor(today);
	switch (period) {
		case 'today':
			return day === t;
		case 'yesterday':
			return day === t - 1;
		case 'tomorrow':
			return day === t + 1;
		case 'last7Days':
			return day <= t && t - day <= 6;
		case 'thisWeek':
		case 'lastWeek':
		case 'nextWeek': {
			const sunday = t - serialToDateParts(t, date1904).weekday;
			const start =
				period === 'thisWeek' ? sunday : period === 'lastWeek' ? sunday - 7 : sunday + 7;
			return day >= start && day <= start + 6;
		}
		case 'thisMonth':
		case 'lastMonth':
		case 'nextMonth': {
			const now = serialToDateParts(t, date1904);
			let month = now.month + (period === 'lastMonth' ? -1 : period === 'nextMonth' ? 1 : 0);
			let year = now.year;
			if (month < 1) {
				month = 12;
				year -= 1;
			} else if (month > 12) {
				month = 1;
				year += 1;
			}
			const d = serialToDateParts(day, date1904);
			return d.year === year && d.month === month;
		}
		default:
			return false;
	}
}
