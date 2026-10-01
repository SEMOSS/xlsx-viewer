/**
 * Reader for OLE2 compound files ([MS-CFB]): the container of encrypted
 * OOXML packages (`EncryptionInfo`, `EncryptedPackage`) and of legacy
 * `.xls` workbooks (`Workbook` / `Book` streams).
 */

import { isCompoundFile } from '../opc-package';

const MAX_REGULAR_SECTOR = 0xfffffffa;
const END_OF_CHAIN = 0xfffffffe;
const NO_STREAM = 0xffffffff;
const HEADER_DIFAT_ENTRIES = 109;
const DIRECTORY_ENTRY_SIZE = 128;

export class CfbFormatError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'CfbFormatError';
	}
}

export interface CfbEntry {
	name: string;
	/** Slash-separated path below the root, e.g. `EncryptionInfo`, `\u0006DataSpaces/Version`. */
	path: string;
	type: 'root' | 'storage' | 'stream';
	size: number;
}

export interface CompoundFile {
	/** Every storage and stream (the root first), in directory order. */
	entries: CfbEntry[];
	/** Contents of a stream by path (case-insensitive, as names in CFB are). */
	stream(path: string): Uint8Array | undefined;
}

interface RawEntry {
	name: string;
	type: number;
	left: number;
	right: number;
	child: number;
	start: number;
	size: number;
}

/** Parse a compound file. Throws CfbFormatError for corrupt headers or chains. */
export function readCompoundFile(bytes: Uint8Array): CompoundFile {
	if (bytes.length < 512 || !isCompoundFile(bytes)) {
		throw new CfbFormatError('Not a compound file');
	}
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const shift = view.getUint16(0x1e, true);
	const miniShift = view.getUint16(0x20, true);
	if ((shift !== 9 && shift !== 12) || miniShift !== 6) {
		throw new CfbFormatError(`Unsupported sector size (shift ${shift}, mini shift ${miniShift})`);
	}
	const size = 1 << shift;
	const miniSize = 1 << miniShift;
	const cutoff = view.getUint32(0x38, true);
	const sectorOffset = (n: number): number => (n + 1) * size;
	const readSector = (n: number): Uint8Array => {
		const off = sectorOffset(n);
		if (n >= MAX_REGULAR_SECTOR || off + size > bytes.length) {
			// Files may be truncated at the end of the last sector: pad with zeros.
			if (off < bytes.length && n < MAX_REGULAR_SECTOR) {
				const out = new Uint8Array(size);
				out.set(bytes.subarray(off));
				return out;
			}
			throw new CfbFormatError(`Sector ${n} is outside the file`);
		}
		return bytes.subarray(off, off + size);
	};

	// FAT sector numbers: 109 in the header, the rest along the DIFAT chain.
	const fatSectors: number[] = [];
	for (let i = 0; i < HEADER_DIFAT_ENTRIES; i++) {
		const s = view.getUint32(0x4c + i * 4, true);
		if (s < MAX_REGULAR_SECTOR) {
			fatSectors.push(s);
		}
	}
	const seen = new Set<number>();
	for (let difat = view.getUint32(0x44, true); difat < MAX_REGULAR_SECTOR; ) {
		if (seen.has(difat)) {
			throw new CfbFormatError('DIFAT chain loops');
		}
		seen.add(difat);
		const sector = readSector(difat);
		const sv = new DataView(sector.buffer, sector.byteOffset, sector.byteLength);
		for (let i = 0; i < size / 4 - 1; i++) {
			const s = sv.getUint32(i * 4, true);
			if (s < MAX_REGULAR_SECTOR) {
				fatSectors.push(s);
			}
		}
		difat = sv.getUint32(size - 4, true);
	}
	const fat = new Uint32Array(fatSectors.length * (size / 4));
	fatSectors.forEach((s, i) => {
		const sector = readSector(s);
		const sv = new DataView(sector.buffer, sector.byteOffset, sector.byteLength);
		for (let j = 0; j < size / 4; j++) {
			fat[i * (size / 4) + j] = sv.getUint32(j * 4, true);
		}
	});

	const chain = (start: number, table: Uint32Array, limit: number): number[] => {
		const out: number[] = [];
		for (let s = start; s < MAX_REGULAR_SECTOR; s = table[s] ?? END_OF_CHAIN) {
			if (out.length > limit) {
				throw new CfbFormatError('Sector chain loops');
			}
			out.push(s);
		}
		return out;
	};
	const readChain = (start: number, length?: number): Uint8Array => {
		const sectors = chain(start, fat, fat.length);
		const out = new Uint8Array(sectors.length * size);
		sectors.forEach((s, i) => out.set(readSector(s), i * size));
		return length === undefined ? out : out.subarray(0, Math.min(length, out.length));
	};

	// Directory entries.
	const dir = readChain(view.getUint32(0x30, true));
	const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
	const raw: RawEntry[] = [];
	const major = view.getUint16(0x1a, true);
	for (let e = 0; e + DIRECTORY_ENTRY_SIZE <= dir.length; e += DIRECTORY_ENTRY_SIZE) {
		const len = dv.getUint16(e + 0x40, true);
		let name = '';
		for (let i = 0; i + 2 < len && i < 64; i += 2) {
			name += String.fromCharCode(dv.getUint16(e + i, true));
		}
		const low = dv.getUint32(e + 0x78, true);
		// Version 3 files may leave garbage in the high size word.
		const high = major >= 4 ? dv.getUint32(e + 0x7c, true) : 0;
		raw.push({
			name,
			type: dir[e + 0x42],
			left: dv.getUint32(e + 0x44, true),
			right: dv.getUint32(e + 0x48, true),
			child: dv.getUint32(e + 0x4c, true),
			start: dv.getUint32(e + 0x74, true),
			size: high * 2 ** 32 + low,
		});
	}
	const root = raw[0];
	if (!root || root.type !== 5) {
		throw new CfbFormatError('Missing root entry');
	}

	// Mini stream (inside the root entry's chain) and its allocation table.
	const miniFatStart = view.getUint32(0x3c, true);
	const miniFatBytes =
		miniFatStart < MAX_REGULAR_SECTOR ? readChain(miniFatStart) : new Uint8Array(0);
	const miniFat = new Uint32Array(miniFatBytes.length / 4);
	const mv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
	for (let i = 0; i < miniFat.length; i++) {
		miniFat[i] = mv.getUint32(i * 4, true);
	}
	let miniStream: Uint8Array | undefined;
	const readMini = (start: number, length: number): Uint8Array => {
		miniStream ??=
			root.start < MAX_REGULAR_SECTOR ? readChain(root.start, root.size) : new Uint8Array(0);
		const sectors = chain(start, miniFat, miniFat.length);
		const out = new Uint8Array(sectors.length * miniSize);
		sectors.forEach((s, i) =>
			out.set(miniStream!.subarray(s * miniSize, (s + 1) * miniSize), i * miniSize),
		);
		return out.subarray(0, Math.min(length, out.length));
	};

	// Walk the red-black trees to build paths.
	const entries: CfbEntry[] = [{ name: root.name, path: '', type: 'root', size: root.size }];
	const byPath = new Map<string, RawEntry>();
	const visit = (id: number, parent: string, guard: Set<number>): void => {
		if (id === NO_STREAM || id >= raw.length || guard.has(id)) {
			return;
		}
		guard.add(id);
		const e = raw[id];
		visit(e.left, parent, guard);
		const path = parent ? `${parent}/${e.name}` : e.name;
		if (e.type === 1 || e.type === 2) {
			entries.push({ name: e.name, path, type: e.type === 1 ? 'storage' : 'stream', size: e.size });
			byPath.set(path.toLowerCase(), e);
			if (e.type === 1) {
				visit(e.child, path, guard);
			}
		}
		visit(e.right, parent, guard);
	};
	visit(root.child, '', new Set());

	return {
		entries,
		stream(path: string): Uint8Array | undefined {
			const e = byPath.get(path.toLowerCase());
			if (!e || e.type !== 2) {
				return undefined;
			}
			if (e.size === 0) {
				return new Uint8Array(0);
			}
			return e.size < cutoff ? readMini(e.start, e.size) : readChain(e.start, e.size);
		},
	};
}
