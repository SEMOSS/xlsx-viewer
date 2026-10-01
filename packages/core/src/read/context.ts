import type { OpcPackage } from '../package/opc-package';
import { createDefaultStyleSheet, createDefaultTheme } from '../styles/defaults';
import type { RichTextRun } from '../types/cell';
import type { MediaItem } from '../types/drawing';
import type { StyleSheet } from '../types/style';
import type { Theme } from '../types/theme';
import type { XlsxWarning } from '../types/warnings';

export interface ReadOptions {
	/** Password for encrypted workbooks. */
	password?: string;
	/** Skip drawings (pictures, charts, shapes) for faster loading. */
	skipDrawings?: boolean;
	/** Year used to complete dates typed without one (testing aid). */
	now?: Date;
	/**
	 * Name of the file the bytes came from. A `.csv`, `.tsv` or `.txt` name
	 * opens the text as delimited values (a one-sheet workbook named after the file).
	 */
	fileName?: string;
}

/** The shared-string table (`xl/sharedStrings.xml`) in file order. */
export interface SharedStrings {
	/** Plain text of every `<si>` (rich runs concatenated, phonetic runs dropped). */
	strings: string[];
	/** Rich-text runs by string index, only for formatted strings. */
	rich: Map<number, RichTextRun[]>;
}

export function emptySharedStrings(): SharedStrings {
	return { strings: [], rich: new Map() };
}

/** State shared by every part reader while one workbook loads. */
export class ReadContext {
	readonly warnings: XlsxWarning[] = [];
	/** Binary parts (images) loaded so far, keyed by package path. Becomes `workbook.media`. */
	readonly media = new Map<string, MediaItem>();
	date1904 = false;
	sharedStrings: SharedStrings = emptySharedStrings();
	styles: StyleSheet = createDefaultStyleSheet();
	theme: Theme = createDefaultTheme();
	/**
	 * Cell-metadata indices (`cm`, 1-based) that `xl/metadata.xml` marks as
	 * dynamic arrays. `undefined` when the workbook has no metadata part.
	 */
	dynamicArrayCm?: ReadonlySet<number>;
	private readonly warningKeys = new Set<string>();

	constructor(
		readonly pkg: OpcPackage,
		readonly options: ReadOptions = {},
	) {}

	/**
	 * Record a warning. By default a warning with the same code, part and
	 * sheet is only recorded once, so repeated problems in one part do not
	 * flood the list.
	 */
	warn(
		code: string,
		message: string,
		details: { part?: string; sheet?: string; cell?: string } = {},
		dedupe = true,
	): void {
		if (dedupe) {
			const key = `${code}\u0000${details.part ?? ''}\u0000${details.sheet ?? ''}`;
			if (this.warningKeys.has(key)) {
				return;
			}
			this.warningKeys.add(key);
		}
		const w: XlsxWarning = { code, message };
		if (details.part !== undefined) {
			w.part = details.part;
		}
		if (details.sheet !== undefined) {
			w.sheet = details.sheet;
		}
		if (details.cell !== undefined) {
			w.cell = details.cell;
		}
		this.warnings.push(w);
	}

	/** Record an `unreadable-part` warning for an exception thrown while reading `part`. */
	unreadable(part: string, error: unknown, what = 'part', sheet?: string): void {
		const reason = error instanceof Error ? error.message : String(error);
		this.warn(
			'unreadable-part',
			`Could not read ${what} ${part}: ${reason}`,
			{ part, sheet },
			false,
		);
	}
}

/** Run `fn`; on failure record an `unreadable-part` warning and return `fallback`. */
export async function attempt<T>(
	ctx: ReadContext,
	part: string,
	what: string,
	fn: () => Promise<T> | T,
	fallback: T,
	sheet?: string,
): Promise<T> {
	try {
		return await fn();
	} catch (e) {
		ctx.unreadable(part, e, what, sheet);
		return fallback;
	}
}

/** Synchronous variant of {@link attempt}. */
export function attemptSync<T>(
	ctx: ReadContext,
	part: string,
	what: string,
	fn: () => T,
	fallback: T,
	sheet?: string,
): T {
	try {
		return fn();
	} catch (e) {
		ctx.unreadable(part, e, what, sheet);
		return fallback;
	}
}
