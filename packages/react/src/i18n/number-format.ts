/**
 * Locale-aware numbers for the viewer chrome (zoom, counts, page numbers).
 * Cell values are not formatted here: they follow the workbook's number formats.
 */

const formatters = new Map<string, Intl.NumberFormat>();

/** A cached `Intl.NumberFormat` for `locale` and `options`. */
export function numberFormatter(
	locale: string,
	options: Intl.NumberFormatOptions = {},
): Intl.NumberFormat {
	const id = `${locale}|${JSON.stringify(options)}`;
	let formatter = formatters.get(id);
	if (!formatter) {
		formatter = new Intl.NumberFormat(locale, options);
		formatters.set(id, formatter);
	}
	return formatter;
}

/** `value` in the conventions of `locale` (`1,234.5`, `1.234,5`, `1 234,5`). */
export function formatNumber(
	value: number,
	locale: string,
	options?: Intl.NumberFormatOptions,
): string {
	return numberFormatter(locale, options).format(value);
}

/** A whole percentage (`125` for 125%) as `locale` writes it (`125%`, `125 %`). */
export function formatPercent(percent: number, locale: string): string {
	return numberFormatter(locale, { style: 'percent', maximumFractionDigits: 0 }).format(
		percent / 100,
	);
}
