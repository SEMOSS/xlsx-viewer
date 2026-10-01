import type {
	AlignmentSpec,
	BorderSpec,
	CellXf,
	DifferentialFormat,
	FillSpec,
	FontSpec,
	ProtectionSpec,
	StyleSheet,
} from '../types/style';
import {
	builtinNumberFormatId,
	FIRST_CUSTOM_NUMFMT_ID,
	numberFormatCode,
} from './builtin-number-formats';

/** Changes to apply on top of an existing cell format. `null` clears a component. */
export interface StylePatch {
	font?: Partial<FontSpec>;
	/** Replace the fill; `null` removes it. */
	fill?: FillSpec | null;
	/** Merge edges into the border; set an edge to `undefined` to keep it, or `{ style: 'none' }` to clear. */
	border?: Partial<BorderSpec>;
	/** Replace the whole border. */
	borderReplace?: BorderSpec;
	/** Merge into the alignment; `null` resets alignment. */
	alignment?: Partial<AlignmentSpec> | null;
	/** Number format code. */
	numFmt?: string;
	protection?: Partial<ProtectionSpec>;
	quotePrefix?: boolean;
}

/** Stable JSON with sorted keys and `undefined` dropped, used as a dedupe key. */
export function canonicalKey(value: unknown): string {
	return JSON.stringify(value, (_k, v: unknown) => {
		if (v && typeof v === 'object' && !Array.isArray(v)) {
			const out: Record<string, unknown> = {};
			for (const key of Object.keys(v as Record<string, unknown>).sort()) {
				const item = (v as Record<string, unknown>)[key];
				if (item !== undefined) {
					out[key] = item;
				}
			}
			return out;
		}
		return v;
	});
}

function stripUndefined<T extends object>(o: T): T {
	const out = {} as T;
	for (const k of Object.keys(o) as Array<keyof T>) {
		if (o[k] !== undefined) {
			out[k] = o[k];
		}
	}
	return out;
}

/**
 * Append-only interning of style components. Existing entries are never
 * modified or reordered (cells in untouched sheets keep their indices);
 * identical new components are reused.
 */
export class StylePool {
	private fontIndex?: Map<string, number>;
	private fillIndex?: Map<string, number>;
	private borderIndex?: Map<string, number>;
	private xfIndex?: Map<string, number>;
	private dxfIndex?: Map<string, number>;

	constructor(readonly styles: StyleSheet) {}

	private static buildIndex<T>(items: readonly T[]): Map<string, number> {
		const m = new Map<string, number>();
		items.forEach((item, i) => {
			const k = canonicalKey(item);
			if (!m.has(k)) {
				m.set(k, i);
			}
		});
		return m;
	}

	private intern<T>(items: T[], index: Map<string, number>, item: T): number {
		const k = canonicalKey(item);
		const hit = index.get(k);
		if (hit !== undefined) {
			return hit;
		}
		items.push(item);
		index.set(k, items.length - 1);
		return items.length - 1;
	}

	internFont(font: FontSpec): number {
		this.fontIndex ??= StylePool.buildIndex(this.styles.fonts);
		return this.intern(this.styles.fonts, this.fontIndex, stripUndefined(font));
	}

	internFill(fill: FillSpec): number {
		this.fillIndex ??= StylePool.buildIndex(this.styles.fills);
		return this.intern(this.styles.fills, this.fillIndex, fill);
	}

	internBorder(border: BorderSpec): number {
		this.borderIndex ??= StylePool.buildIndex(this.styles.borders);
		return this.intern(this.styles.borders, this.borderIndex, stripUndefined(border));
	}

	internXf(xf: CellXf): number {
		this.xfIndex ??= StylePool.buildIndex(this.styles.cellXfs);
		return this.intern(this.styles.cellXfs, this.xfIndex, stripUndefined(xf));
	}

	internDxf(dxf: DifferentialFormat): number {
		this.dxfIndex ??= StylePool.buildIndex(this.styles.dxfs);
		return this.intern(this.styles.dxfs, this.dxfIndex, stripUndefined(dxf));
	}

	/** Number format id for a code: built-in id when one matches, else a (new) custom id. */
	internNumFmt(code: string): number {
		if (code === 'General') {
			return 0;
		}
		const builtin = builtinNumberFormatId(code);
		if (builtin !== undefined) {
			return builtin;
		}
		for (const [id, c] of Object.entries(this.styles.numFmts)) {
			if (c === code) {
				return Number(id);
			}
		}
		let next = FIRST_CUSTOM_NUMFMT_ID;
		for (const id of Object.keys(this.styles.numFmts)) {
			next = Math.max(next, Number(id) + 1);
		}
		this.styles.numFmts[next] = code;
		return next;
	}

	/** Format code currently used by an xf. */
	numFmtCodeOf(xfId: number): string {
		const xf = this.styles.cellXfs[xfId] ?? this.styles.cellXfs[0];
		return numberFormatCode(xf?.numFmtId ?? 0, this.styles.numFmts);
	}

	/** Return the id of a cell format equal to `baseXfId` with `patch` applied (created on demand). */
	derive(baseXfId: number | undefined, patch: StylePatch): number {
		const base: CellXf = {
			...(this.styles.cellXfs[baseXfId ?? 0] ?? {
				numFmtId: 0,
				fontId: 0,
				fillId: 0,
				borderId: 0,
				xfId: 0,
			}),
		};
		const next: CellXf = { ...base };
		if (patch.font) {
			const font = { ...(this.styles.fonts[base.fontId] ?? {}), ...patch.font };
			next.fontId = this.internFont(font);
			next.applyFont = true;
		}
		if (patch.fill !== undefined) {
			next.fillId = patch.fill === null ? 0 : this.internFill(patch.fill);
			next.applyFill = true;
		}
		if (patch.borderReplace) {
			next.borderId = this.internBorder(patch.borderReplace);
			next.applyBorder = true;
		} else if (patch.border) {
			const merged: BorderSpec = { ...(this.styles.borders[base.borderId] ?? {}) };
			for (const [k, v] of Object.entries(patch.border) as Array<[keyof BorderSpec, unknown]>) {
				if (v === undefined) {
					continue;
				}
				if (typeof v === 'object' && v !== null && (v as { style?: string }).style === 'none') {
					delete merged[k];
				} else {
					(merged as Record<string, unknown>)[k] = v;
				}
			}
			next.borderId = this.internBorder(merged);
			next.applyBorder = true;
		}
		if (patch.alignment !== undefined) {
			if (patch.alignment === null) {
				delete next.alignment;
			} else {
				const merged = stripUndefined({ ...(base.alignment ?? {}), ...patch.alignment });
				if (Object.keys(merged).length) {
					next.alignment = merged;
				} else {
					delete next.alignment;
				}
			}
			next.applyAlignment = true;
		}
		if (patch.numFmt !== undefined) {
			next.numFmtId = this.internNumFmt(patch.numFmt);
			next.applyNumberFormat = true;
		}
		if (patch.protection) {
			next.protection = stripUndefined({ ...(base.protection ?? {}), ...patch.protection });
			next.applyProtection = true;
		}
		if (patch.quotePrefix !== undefined) {
			if (patch.quotePrefix) {
				next.quotePrefix = true;
			} else {
				delete next.quotePrefix;
			}
		}
		if (next.xfId === undefined) {
			next.xfId = 0;
		}
		return this.internXf(next);
	}
}
