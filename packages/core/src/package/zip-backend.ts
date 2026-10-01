import JSZip from 'jszip';

import { hasNativeDeflate, writeZip, ZipReader, type RawZipEntry, type ZipWriteEntry } from './zip';

/**
 * The ZIP container behind an Open XML package. With the platform's
 * compression streams (browsers, Node 21+) entries are inflated natively and
 * unchanged entries are copied into saved files without recompression;
 * elsewhere JSZip does the work.
 */
export interface ZipArchive {
	/** Entry names (files only), in archive order. */
	names(): string[];
	bytes(name: string): Promise<Uint8Array | undefined>;
	/** Compressed bytes of an entry, when they can be copied as they are. */
	raw(name: string): RawZipEntry | undefined;
}

class NativeArchive implements ZipArchive {
	constructor(private readonly reader: ZipReader) {}

	names(): string[] {
		return this.reader.entries.map((e) => e.name);
	}

	bytes(name: string): Promise<Uint8Array | undefined> {
		return this.reader.bytes(name);
	}

	raw(name: string): RawZipEntry | undefined {
		return this.reader.raw(name);
	}
}

class JsZipArchive implements ZipArchive {
	constructor(private readonly zip: JSZip) {}

	names(): string[] {
		return Object.keys(this.zip.files).filter((n) => !this.zip.files[n].dir);
	}

	async bytes(name: string): Promise<Uint8Array | undefined> {
		const file = this.zip.files[name];
		return file && !file.dir ? file.async('uint8array') : undefined;
	}

	raw(): RawZipEntry | undefined {
		return undefined;
	}
}

export async function openZipArchive(bytes: Uint8Array): Promise<ZipArchive> {
	if (hasNativeDeflate()) {
		return new NativeArchive(new ZipReader(bytes));
	}
	return new JsZipArchive(await JSZip.loadAsync(bytes));
}

/** A part of the archive being written: new contents, or an entry of the source archive. */
export type PendingEntry =
	| { name: string; data: Uint8Array }
	| { name: string; source: ZipArchive; sourceName: string };

/** Build the archive (native streams when available, else JSZip). */
export async function buildZipArchive(entries: readonly PendingEntry[]): Promise<Uint8Array> {
	if (hasNativeDeflate()) {
		const out: ZipWriteEntry[] = [];
		for (const e of entries) {
			if ('data' in e) {
				out.push({ name: e.name, data: e.data });
				continue;
			}
			const raw = e.source.raw(e.sourceName);
			const data = raw ? undefined : await e.source.bytes(e.sourceName);
			if (raw) {
				out.push({ name: e.name, raw });
			} else if (data) {
				out.push({ name: e.name, data });
			}
		}
		return writeZip(out);
	}
	const zip = new JSZip();
	for (const e of entries) {
		const data = 'data' in e ? e.data : await e.source.bytes(e.sourceName);
		if (data) {
			zip.file(e.name, data, { binary: true });
		}
	}
	return zip.generateAsync({
		type: 'uint8array',
		compression: 'DEFLATE',
		compressionOptions: { level: 6 },
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	});
}
