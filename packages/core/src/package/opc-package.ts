import type { Relationship } from '../types/sheet';
import { ContentTypes } from './content-types';
import { normalizePartPath, relsPathFor, resolveRelTarget } from './paths';
import { parseRelationships, serializeRelationships } from './relationships';
import { buildZipArchive, openZipArchive, type PendingEntry, type ZipArchive } from './zip-backend';

export type BinaryInput = Uint8Array | ArrayBuffer;

/** Signature of an OLE2 compound file (encrypted OOXML packages and legacy `.xls`). */
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function toUint8Array(data: BinaryInput): Uint8Array {
	return data instanceof Uint8Array ? data : new Uint8Array(data);
}

/** The bytes are an OLE2 compound file rather than a ZIP package. */
export function isCompoundFile(data: BinaryInput): boolean {
	const bytes = toUint8Array(data);
	if (bytes.length < 8) {
		return false;
	}
	return CFB_SIGNATURE.every((b, i) => bytes[i] === b);
}

/** The bytes start with a ZIP local file header. */
export function isZipFile(data: BinaryInput): boolean {
	const bytes = toUint8Array(data);
	return (
		bytes.length >= 4 &&
		bytes[0] === 0x50 &&
		bytes[1] === 0x4b &&
		bytes[2] === 0x03 &&
		bytes[3] === 0x04
	);
}

const utf8 = new TextDecoder();
const encoder = new TextEncoder();

/**
 * Read-only view of an Open Packaging Conventions (ZIP) package.
 * Relationships are parsed lazily and cached per source part.
 */
export class OpcPackage {
	private readonly relsCache = new Map<string, Relationship[]>();
	private readonly nameIndex = new Map<string, string>();

	private constructor(
		readonly archive: ZipArchive,
		readonly contentTypes: ContentTypes,
	) {
		for (const name of archive.names()) {
			this.nameIndex.set(normalizePartPath(name).toLowerCase(), name);
		}
	}

	static async open(data: BinaryInput): Promise<OpcPackage> {
		const archive = await openZipArchive(toUint8Array(data));
		const ctEntry = archive.names().find((n) => n.toLowerCase() === '[content_types].xml');
		const ctBytes = ctEntry ? await archive.bytes(ctEntry) : undefined;
		const contentTypes = ctBytes
			? ContentTypes.parse(utf8.decode(ctBytes))
			: ContentTypes.createDefault();
		return new OpcPackage(archive, contentTypes);
	}

	/** Exact ZIP entry name for a part path (part names are case-insensitive in OPC). */
	entryName(partPath: string): string | undefined {
		return this.nameIndex.get(normalizePartPath(partPath).toLowerCase());
	}

	has(partPath: string): boolean {
		return this.entryName(partPath) !== undefined;
	}

	/** All part paths (normalised, without leading slash). */
	partPaths(): string[] {
		return [...this.nameIndex.values()].map((n) => normalizePartPath(n));
	}

	/** ZIP entry names in archive order (exact case). */
	entryNames(): string[] {
		return this.archive.names();
	}

	async text(partPath: string): Promise<string | undefined> {
		const bytes = await this.bytes(partPath);
		return bytes ? utf8.decode(bytes) : undefined;
	}

	async bytes(partPath: string): Promise<Uint8Array | undefined> {
		const name = this.entryName(partPath);
		return name ? this.archive.bytes(name) : undefined;
	}

	/** Relationships whose source is `partPath` (use `''` for the package-level `_rels/.rels`). */
	async rels(partPath: string): Promise<Relationship[]> {
		const key = normalizePartPath(partPath);
		const cached = this.relsCache.get(key);
		if (cached) {
			return cached;
		}
		const relsPath = key === '' ? '_rels/.rels' : relsPathFor(key);
		const xml = await this.text(relsPath);
		const rels = xml ? parseRelationships(xml) : [];
		this.relsCache.set(key, rels);
		return rels;
	}

	/** Resolve an internal relationship target to a package path. */
	resolve(sourcePart: string, rel: Relationship): string {
		return resolveRelTarget(sourcePart, rel.target);
	}
}

/**
 * Builds the output ZIP. Starts as a copy of a source package (so untouched
 * parts are carried over without being decoded, and when possible without
 * being recompressed) or empty.
 */
export class PackageWriter {
	readonly contentTypes: ContentTypes;
	/** Output entries by exact name, in archive order. */
	private readonly entries = new Map<string, PendingEntry>();

	constructor(private readonly source?: OpcPackage) {
		if (source) {
			for (const name of source.entryNames()) {
				if (name.toLowerCase() !== '[content_types].xml') {
					this.entries.set(name, { name, source: source.archive, sourceName: name });
				}
			}
			this.contentTypes = ContentTypes.parse(source.contentTypes.serialize());
		} else {
			this.contentTypes = ContentTypes.createDefault();
		}
	}

	private findEntry(partPath: string): string | undefined {
		const want = normalizePartPath(partPath).toLowerCase();
		for (const name of this.entries.keys()) {
			if (normalizePartPath(name).toLowerCase() === want) {
				return name;
			}
		}
		return undefined;
	}

	has(partPath: string): boolean {
		return this.findEntry(partPath) !== undefined;
	}

	/** Write (or replace) a text part. `contentType` registers an override. */
	setText(partPath: string, xml: string, contentType?: string): void {
		this.setBytes(partPath, encoder.encode(xml), contentType);
	}

	/** Write (or replace) a binary part. */
	setBytes(partPath: string, data: Uint8Array, contentType?: string): void {
		const p = normalizePartPath(partPath);
		const existing = this.findEntry(p);
		if (existing && existing !== p) {
			this.entries.delete(existing);
		}
		this.entries.set(p, { name: p, data });
		if (contentType) {
			this.contentTypes.setOverride(p, contentType);
		}
	}

	setRels(sourcePart: string, rels: readonly Relationship[]): void {
		const relsPath = sourcePart === '' ? '_rels/.rels' : relsPathFor(sourcePart);
		if (rels.length === 0) {
			this.remove(relsPath);
			return;
		}
		this.setText(relsPath, serializeRelationships(rels));
	}

	remove(partPath: string): void {
		const existing = this.findEntry(partPath);
		if (existing) {
			this.entries.delete(existing);
		}
		this.contentTypes.removeOverride(partPath);
	}

	/** Part paths currently in the output. */
	partPaths(): string[] {
		return [...this.entries.keys()].map((n) => normalizePartPath(n));
	}

	/** Produce the final package bytes; `[Content_Types].xml` comes first, as Excel writes it. */
	async generate(): Promise<Uint8Array> {
		const ct: PendingEntry = {
			name: '[Content_Types].xml',
			data: encoder.encode(this.contentTypes.serialize()),
		};
		return buildZipArchive([ct, ...this.entries.values()]);
	}
}
