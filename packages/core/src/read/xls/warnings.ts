import type { XlsxWarning } from '../../types/warnings';

/** Part name used in warnings: the BIFF workbook stream. */
const XLS_PART = 'Workbook';

/**
 * Warning collector for the `.xls` reader. Like the `.xlsx` reader's
 * context, a warning with the same code and sheet is recorded once unless
 * `dedupe` is turned off (a string adds to the key), so a damaged sheet
 * does not flood the list.
 */
export class XlsWarnings {
	readonly list: XlsxWarning[] = [];
	private readonly seen = new Set<string>();

	warn(
		code: string,
		message: string,
		details: { sheet?: string; cell?: string } = {},
		dedupe: boolean | string = true,
	): void {
		if (dedupe !== false) {
			const extra = typeof dedupe === 'string' ? dedupe : '';
			const key = `${code}\u0000${details.sheet ?? ''}\u0000${extra}`;
			if (this.seen.has(key)) {
				return;
			}
			this.seen.add(key);
		}
		const w: XlsxWarning = { code, message, part: XLS_PART };
		if (details.sheet !== undefined) {
			w.sheet = details.sheet;
		}
		if (details.cell !== undefined) {
			w.cell = details.cell;
		}
		this.list.push(w);
	}

	/** Record a `corrupt-record` warning for an exception thrown while reading a record. */
	corrupt(what: string, error: unknown, sheet?: string): void {
		const reason = error instanceof Error ? error.message : String(error);
		this.warn('corrupt-record', `Could not read ${what}: ${reason}`, { sheet }, what);
	}
}
