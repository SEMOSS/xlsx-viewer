import { formatValue } from 'xlsx-viewer-core';

import { formatNumber } from '../../i18n/number-format';

/** Beyond this magnitude the status bar switches to scientific notation. */
const SCIENTIFIC_FROM = 1e15;

/**
 * Text of a status-bar statistic (Sum, Average, Min, Max). With the General
 * format the number follows the UI locale (`Intl.NumberFormat`: `1,234.5`,
 * `1.234,5`); an explicit number format of the active cell (currency, date,
 * percent...) is applied the way the cell itself displays, as Excel does.
 */
export function formatStatistic(
	value: number,
	numFmt: string,
	locale: string,
	date1904 = false,
): string {
	if (numFmt.trim() === '' || /^general$/i.test(numFmt.trim())) {
		const options: Intl.NumberFormatOptions =
			Math.abs(value) >= SCIENTIFIC_FROM
				? { notation: 'scientific', maximumFractionDigits: 5 }
				: { maximumFractionDigits: 10 };
		return formatNumber(value, locale, options);
	}
	return formatValue(value, numFmt, { date1904 }).text;
}
