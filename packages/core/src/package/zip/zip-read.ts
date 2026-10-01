import { inflateRaw } from './streams';

/** A stored entry: its compressed bytes can be copied into a new archive unchanged. */
export interface RawZipEntry {
	method: 0 | 8;
	crc32: number;
	size: number;
	compressed: Uint8Array;
}

export interface ZipEntryInfo {
	name: string;
	method: number;
	crc32: number;
	compressedSize: number;
	size: number;
	offset: number;
}

const EOCD = 0x06054b50;
const ZIP64_LOCATOR = 0x07064b50;
const ZIP64_EOCD = 0x06064b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;

export class ZipFormatError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'ZipFormatError';
	}
}

const decoder = new TextDecoder();
const latin1 = new TextDecoder('latin1');

function u64(view: DataView, at: number): number {
	return view.getUint32(at, true) + view.getUint32(at + 4, true) * 2 ** 32;
}

/**
 * Reader for ZIP archives (the container of every Open XML package), with
 * entries inflated by the platform's DecompressionStream. Supports ZIP64
 * and data descriptors; encrypted entries are rejected.
 */
export class ZipReader {
	readonly entries: ZipEntryInfo[] = [];
	private readonly byName = new Map<string, ZipEntryInfo>();

	constructor(private readonly data: Uint8Array) {
		const bytes = data;
		const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
		let eocd = -1;
		for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 0xffff); i--) {
			if (view.getUint32(i, true) === EOCD) {
				eocd = i;
				break;
			}
		}
		if (eocd < 0) {
			throw new ZipFormatError('Not a ZIP archive (no end of central directory)');
		}
		let count = view.getUint16(eocd + 10, true);
		let dirOffset = view.getUint32(eocd + 16, true);
		const locator = eocd - 20;
		if (locator >= 0 && view.getUint32(locator, true) === ZIP64_LOCATOR) {
			const z = u64(view, locator + 8);
			if (z + 56 <= bytes.length && view.getUint32(z, true) === ZIP64_EOCD) {
				count = u64(view, z + 32);
				dirOffset = u64(view, z + 48);
			}
		}
		let p = dirOffset;
		for (let i = 0; i < count; i++) {
			if (p + 46 > bytes.length || view.getUint32(p, true) !== CENTRAL) {
				throw new ZipFormatError('Damaged ZIP central directory');
			}
			const flags = view.getUint16(p + 8, true);
			const method = view.getUint16(p + 10, true);
			const crc = view.getUint32(p + 16, true);
			let compressedSize = view.getUint32(p + 20, true);
			let size = view.getUint32(p + 24, true);
			const nameLen = view.getUint16(p + 28, true);
			const extraLen = view.getUint16(p + 30, true);
			const commentLen = view.getUint16(p + 32, true);
			let offset = view.getUint32(p + 42, true);
			const rawName = bytes.subarray(p + 46, p + 46 + nameLen);
			const name = (flags & 0x800 ? decoder : latin1).decode(rawName);
			// ZIP64 extra field: the 0xFFFFFFFF fields follow in order.
			let e = p + 46 + nameLen;
			const extraEnd = e + extraLen;
			while (e + 4 <= extraEnd) {
				const id = view.getUint16(e, true);
				const len = view.getUint16(e + 2, true);
				if (id === 0x0001) {
					let q = e + 4;
					if (size === 0xffffffff) {
						size = u64(view, q);
						q += 8;
					}
					if (compressedSize === 0xffffffff) {
						compressedSize = u64(view, q);
						q += 8;
					}
					if (offset === 0xffffffff) {
						offset = u64(view, q);
					}
				}
				e += 4 + len;
			}
			if (flags & 1) {
				throw new ZipFormatError(`Encrypted ZIP entries are not supported (${name})`);
			}
			if (!name.endsWith('/')) {
				const info = { name, method, crc32: crc, compressedSize, size, offset };
				this.entries.push(info);
				this.byName.set(name, info);
			}
			p = extraEnd + commentLen;
		}
	}

	has(name: string): boolean {
		return this.byName.has(name);
	}

	/** The entry's compressed bytes and metadata, for copying into another archive. */
	raw(name: string): RawZipEntry | undefined {
		const info = this.byName.get(name);
		if (!info || (info.method !== 0 && info.method !== 8)) {
			return undefined;
		}
		const view = new DataView(this.data.buffer, this.data.byteOffset, this.data.byteLength);
		const at = info.offset;
		if (at + 30 > this.data.length || view.getUint32(at, true) !== LOCAL) {
			throw new ZipFormatError(`Damaged ZIP entry ${name}`);
		}
		const start = at + 30 + view.getUint16(at + 26, true) + view.getUint16(at + 28, true);
		const end = start + info.compressedSize;
		if (end > this.data.length) {
			throw new ZipFormatError(`Truncated ZIP entry ${name}`);
		}
		return {
			method: info.method,
			crc32: info.crc32,
			size: info.size,
			compressed: this.data.subarray(start, end),
		};
	}

	/** The entry's uncompressed contents. */
	async bytes(name: string): Promise<Uint8Array | undefined> {
		const raw = this.raw(name);
		if (!raw) {
			if (this.byName.has(name)) {
				throw new ZipFormatError(`Unsupported ZIP compression in ${name}`);
			}
			return undefined;
		}
		return raw.method === 0 ? raw.compressed.slice() : inflateRaw(raw.compressed);
	}
}
