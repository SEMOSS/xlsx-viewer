/**
 * Shared state of one save: the output package, the source package, unique
 * part-name allocation, relationships written so far and warnings.
 */
import { nextRelId, normalizePartPath, OpcPackage, PackageWriter, relsPathFor } from '../package';
import { StylePool } from '../styles/style-pool';
import type { Relationship } from '../types/sheet';
import type { DifferentialFormat, StyleSheet } from '../types/style';
import type { XlsxWarning } from '../types/warnings';
import type { Workbook } from '../types/workbook';

export interface WriteOptions {
	/**
	 * Encrypt the file with this password, as Excel's Encrypt with Password does:
	 * ECMA-376 Agile encryption (AES-256, SHA-512, 100000 spins). Empty means none.
	 */
	password?: string;
	/** Force regeneration of every part, ignoring dirty tracking (testing aid). */
	regenerateAll?: boolean;
	/** Application name written to docProps/app.xml. */
	application?: string;
	/** Timestamp used for `dcterms:modified` and new comments (defaults to now). */
	now?: Date;
	/** Receives approximations and skipped content found while saving. */
	onWarning?: (warning: XlsxWarning) => void;
}

/** Duck-typed check so a source package from another module copy is still accepted. */
export function asOpcPackage(value: unknown): OpcPackage | undefined {
	if (value instanceof OpcPackage) {
		return value;
	}
	const v = value as Partial<OpcPackage> | null | undefined;
	if (v && typeof v === 'object' && v.archive && v.contentTypes && typeof v.rels === 'function') {
		return value as OpcPackage;
	}
	return undefined;
}

export class WriteContext {
	readonly writer: PackageWriter;
	readonly now: Date;
	readonly warnings: XlsxWarning[] = [];
	/** Some part was written or removed (the package differs from the source). */
	changed = false;
	private readonly relsWritten = new Map<string, Relationship[]>();
	private readonly present = new Map<string, string>();
	private readonly reserved = new Set<string>();
	private savedStyles?: StyleSheet;
	private dxfPool?: StylePool;

	constructor(
		readonly workbook: Workbook,
		readonly source: OpcPackage | undefined,
		readonly options: WriteOptions,
	) {
		this.writer = new PackageWriter(source);
		for (const p of this.writer.partPaths()) {
			if (p !== '[Content_Types].xml') {
				this.present.set(p.toLowerCase(), p);
			}
		}
		this.now = options.now ?? new Date();
	}

	/**
	 * The stylesheet as written: the model's, plus differential formats that
	 * were interned during this save (the caller's model is not modified).
	 */
	get styles(): StyleSheet {
		return this.savedStyles ?? this.workbook.styles;
	}

	/** Index of a differential format in the saved stylesheet, appended when new. */
	internDxf(dxf: DifferentialFormat): number {
		if (!this.savedStyles) {
			const styles = this.workbook.styles;
			this.savedStyles = { ...styles, dxfs: [...styles.dxfs] };
			this.dxfPool = new StylePool(this.savedStyles);
		}
		return this.dxfPool!.internDxf(dxf);
	}

	/** Every part is regenerated (no source package, or the testing flag). */
	get regenerateAll(): boolean {
		return !this.source || !!this.options.regenerateAll;
	}

	/** The part exists in the output package. */
	has(path: string): boolean {
		return this.present.has(normalizePartPath(path).toLowerCase());
	}

	/** The part exists in the source package. */
	sourceHas(path: string): boolean {
		return !!this.source?.has(path);
	}

	async sourceText(path: string): Promise<string | undefined> {
		return this.source?.has(path) ? this.source.text(path) : undefined;
	}

	/** Reserve a fresh part path `dir/baseN.ext` that is not used in the output. */
	allocate(dir: string, base: string, ext: string): string {
		for (let n = 1; ; n++) {
			const p = `${dir}/${base}${n}.${ext}`;
			const key = p.toLowerCase();
			if (!this.present.has(key) && !this.reserved.has(key)) {
				this.reserved.add(key);
				return p;
			}
		}
	}

	writeText(path: string, xml: string, contentType?: string): void {
		const p = normalizePartPath(path);
		this.writer.setText(p, xml, contentType);
		this.present.set(p.toLowerCase(), p);
		this.changed = true;
	}

	writeBytes(path: string, data: Uint8Array, contentType?: string): void {
		const p = normalizePartPath(path);
		this.writer.setBytes(p, data, contentType);
		this.present.set(p.toLowerCase(), p);
		this.changed = true;
	}

	/** Write the relationships of a part (an empty list removes the `.rels` part). */
	writeRels(sourcePart: string, rels: readonly Relationship[]): void {
		const key = normalizePartPath(sourcePart);
		this.relsWritten.set(key, rels.slice());
		this.writer.setRels(key, rels);
		const relsPath = key === '' ? '_rels/.rels' : relsPathFor(key);
		if (rels.length) {
			this.present.set(relsPath.toLowerCase(), relsPath);
		} else {
			this.present.delete(relsPath.toLowerCase());
		}
		this.changed = true;
	}

	/** Write rels only when they differ from the source rels. */
	async writeRelsIfChanged(sourcePart: string, rels: readonly Relationship[]): Promise<void> {
		const before = await this.rels(sourcePart);
		if (!sameRels(before, rels)) {
			this.writeRels(sourcePart, rels);
		}
	}

	/** Current relationships of a part in the output (written ones win over the source). */
	async rels(sourcePart: string): Promise<Relationship[]> {
		const key = normalizePartPath(sourcePart);
		const written = this.relsWritten.get(key);
		if (written) {
			return written;
		}
		if (!this.source || (key !== '' && !this.source.has(key))) {
			return [];
		}
		return this.source.rels(key);
	}

	/** Remove a part, its relationships part and its content-type override. */
	remove(path: string): void {
		const p = normalizePartPath(path);
		if (this.has(p)) {
			this.writer.remove(p);
			this.present.delete(p.toLowerCase());
			this.changed = true;
		}
		this.writer.contentTypes.removeOverride(p);
		const relsPath = relsPathFor(p);
		if (this.has(relsPath)) {
			this.writer.remove(relsPath);
			this.present.delete(relsPath.toLowerCase());
		}
		this.relsWritten.delete(p);
	}

	/** Output part paths. */
	partPaths(): string[] {
		return [...this.present.values()];
	}

	warn(code: string, message: string, extra: Partial<XlsxWarning> = {}): void {
		const w: XlsxWarning = { code, message, ...extra };
		this.warnings.push(w);
		this.options.onWarning?.(w);
	}
}

/** Relationship lists are equal as sets (order-insensitive). */
export function sameRels(a: readonly Relationship[], b: readonly Relationship[]): boolean {
	if (a.length !== b.length) {
		return false;
	}
	const key = (r: Relationship): string =>
		`${r.id}\u0000${r.type}\u0000${r.target}\u0000${r.targetMode ?? ''}`;
	const set = new Set(a.map(key));
	return b.every((r) => set.has(key(r)));
}

/** Append a relationship with a fresh id; returns the id. */
export function addRel(
	rels: Relationship[],
	type: string,
	target: string,
	targetMode?: 'External',
): string {
	const id = nextRelId(rels);
	rels.push({ id, type, target, ...(targetMode ? { targetMode } : {}) });
	return id;
}
