/**
 * Writer for OLE2 compound files ([MS-CFB]), the container of encrypted OOXML
 * packages. Produces version 3 files: 512-byte sectors, 64-byte mini sectors
 * and the usual 4096-byte mini-stream cutoff. Storages are created for the
 * path prefixes of the streams; every storage's children form a valid
 * red-black tree sorted the way CFB compares names.
 */

const SECTOR_SIZE = 512;
const MINI_SECTOR_SIZE = 64;
const MINI_STREAM_CUTOFF = 4096;
const HEADER_DIFAT_ENTRIES = 109;
const IDS_PER_SECTOR = SECTOR_SIZE / 4;
const DIRECTORY_ENTRY_SIZE = 128;
const ENTRIES_PER_SECTOR = SECTOR_SIZE / DIRECTORY_ENTRY_SIZE;
const MAX_NAME_LENGTH = 31;

const FREE_SECTOR = 0xffffffff;
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const DIFAT_SECTOR = 0xfffffffc;
const NO_STREAM = 0xffffffff;
const EMPTY = new Uint8Array(0);

/** A stream to store: `path` is slash-separated, e.g. `\u0006DataSpaces/Version`. */
export interface CfbStreamInput {
	path: string;
	data: Uint8Array;
}

interface DirNode {
	name: string;
	type: 'root' | 'storage' | 'stream';
	data: Uint8Array;
	children: DirNode[];
	id: number;
	left: number;
	right: number;
	child: number;
	black: boolean;
	/** First sector (mini sector for small streams); ENDOFCHAIN when there is no data. */
	start: number;
	/** Stream size; the mini stream size for the root entry. */
	size: number;
}

function dirNode(name: string, type: DirNode['type'], data: Uint8Array = EMPTY): DirNode {
	return {
		name,
		type,
		data,
		children: [],
		id: 0,
		left: NO_STREAM,
		right: NO_STREAM,
		child: NO_STREAM,
		black: true,
		start: END_OF_CHAIN,
		size: type === 'stream' ? data.length : 0,
	};
}

/** Upper-cased UTF-16 code unit (simple case mapping only, as CFB compares names). */
function upperUnit(name: string, i: number): number {
	const up = name[i].toUpperCase();
	return up.length === 1 ? up.charCodeAt(0) : name.charCodeAt(i);
}

/** CFB name order: shorter names first, then upper-cased code units. */
export function compareCfbNames(a: string, b: string): number {
	if (a.length !== b.length) {
		return a.length - b.length;
	}
	for (let i = 0; i < a.length; i++) {
		const d = upperUnit(a, i) - upperUnit(b, i);
		if (d !== 0) {
			return d;
		}
	}
	return 0;
}

function checkName(name: string, path: string): void {
	if (!name || name.length > MAX_NAME_LENGTH || /[/\\:!]/.test(name)) {
		throw new Error(`Invalid compound file entry name in "${path}"`);
	}
}

/** Build the storage tree for the streams (storages come from the path prefixes). */
function buildTree(streams: readonly CfbStreamInput[]): DirNode {
	const root = dirNode('Root Entry', 'root');
	for (const { path, data } of streams) {
		const parts = path.split('/');
		let parent = root;
		parts.forEach((name, i) => {
			checkName(name, path);
			const existing = parent.children.find((c) => compareCfbNames(c.name, name) === 0);
			if (i === parts.length - 1) {
				if (existing) {
					throw new Error(`Duplicate compound file entry "${path}"`);
				}
				parent.children.push(dirNode(name, 'stream', data));
				return;
			}
			if (existing && existing.type !== 'storage') {
				throw new Error(`"${path}" is below a stream`);
			}
			const storage = existing ?? dirNode(name, 'storage');
			if (!existing) {
				parent.children.push(storage);
			}
			parent = storage;
		});
	}
	return root;
}

/**
 * Link sorted siblings into a balanced binary search tree and return its root
 * id. Splitting at the midpoint leaves every empty link at depth `m` or `m + 1`
 * (`m = floor(log2(n + 1))`); coloring the nodes at depth `m` red and all others
 * black gives each path the same black height and no red node a red child.
 */
function linkSiblings(sorted: DirNode[]): number {
	const redDepth = Math.floor(Math.log2(sorted.length + 1));
	const link = (lo: number, hi: number, depth: number): number => {
		if (lo > hi) {
			return NO_STREAM;
		}
		const mid = (lo + hi) >> 1;
		const n = sorted[mid];
		n.black = depth !== redDepth;
		n.left = link(lo, mid - 1, depth + 1);
		n.right = link(mid + 1, hi, depth + 1);
		return n.id;
	};
	return link(0, sorted.length - 1, 0);
}

/** Directory order: the root first, then depth-first with siblings sorted. */
function flatten(root: DirNode): DirNode[] {
	const out: DirNode[] = [];
	const visit = (n: DirNode): void => {
		n.id = out.length;
		out.push(n);
		n.children.sort((a, b) => compareCfbNames(a.name, b.name));
		n.children.forEach(visit);
	};
	visit(root);
	for (const n of out) {
		if (n.children.length) {
			n.child = linkSiblings(n.children);
		}
	}
	return out;
}

const sectorsFor = (bytes: number, size: number): number => Math.ceil(bytes / size);

/** FAT and DIFAT sector counts that cover `dataSectors` plus themselves. */
function allocationSectors(dataSectors: number): { fat: number; difat: number } {
	let fat = 1;
	for (;;) {
		const difat = Math.max(0, Math.ceil((fat - HEADER_DIFAT_ENTRIES) / (IDS_PER_SECTOR - 1)));
		const needed = sectorsFor(dataSectors + fat + difat, IDS_PER_SECTOR);
		if (needed <= fat) {
			return { fat, difat };
		}
		fat = needed;
	}
}

interface Layout {
	fatCount: number;
	difatStart: number;
	difatCount: number;
	directoryStart: number;
	directoryCount: number;
	miniFatStart: number;
	miniFatCount: number;
	miniStreamStart: number;
	miniStreamSectors: number;
}

function writeHeader(view: DataView, bytes: Uint8Array, l: Layout): void {
	bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
	view.setUint16(0x18, 0x003e, true);
	view.setUint16(0x1a, 0x0003, true);
	view.setUint16(0x1c, 0xfffe, true);
	view.setUint16(0x1e, 9, true);
	view.setUint16(0x20, 6, true);
	view.setUint32(0x2c, l.fatCount, true);
	view.setUint32(0x30, l.directoryStart, true);
	view.setUint32(0x38, MINI_STREAM_CUTOFF, true);
	view.setUint32(0x3c, l.miniFatCount ? l.miniFatStart : END_OF_CHAIN, true);
	view.setUint32(0x40, l.miniFatCount, true);
	view.setUint32(0x44, l.difatCount ? l.difatStart : END_OF_CHAIN, true);
	view.setUint32(0x48, l.difatCount, true);
	// DIFAT: the first 109 FAT sectors in the header, the rest in DIFAT sectors.
	for (let i = 0; i < HEADER_DIFAT_ENTRIES; i++) {
		view.setUint32(0x4c + i * 4, i < l.fatCount ? i : FREE_SECTOR, true);
	}
}

function writeDirectoryEntry(view: DataView, bytes: Uint8Array, off: number, n?: DirNode): void {
	if (!n) {
		// Unused slot: all zeros except NOSTREAM links.
		view.setUint32(off + 0x44, NO_STREAM, true);
		view.setUint32(off + 0x48, NO_STREAM, true);
		view.setUint32(off + 0x4c, NO_STREAM, true);
		return;
	}
	for (let i = 0; i < n.name.length; i++) {
		view.setUint16(off + i * 2, n.name.charCodeAt(i), true);
	}
	view.setUint16(off + 0x40, (n.name.length + 1) * 2, true);
	bytes[off + 0x42] = n.type === 'root' ? 5 : n.type === 'storage' ? 1 : 2;
	bytes[off + 0x43] = n.black ? 1 : 0;
	view.setUint32(off + 0x44, n.left, true);
	view.setUint32(off + 0x48, n.right, true);
	view.setUint32(off + 0x4c, n.child, true);
	view.setUint32(off + 0x74, n.type === 'storage' ? 0 : n.start, true);
	view.setUint32(off + 0x78, n.size, true);
}

/** Serialize streams (and the storages above them) as a version 3 compound file. */
export function writeCompoundFile(streams: readonly CfbStreamInput[]): Uint8Array {
	const entries = flatten(buildTree(streams));
	const withData = entries.filter((n) => n.type === 'stream' && n.size > 0);
	const small = withData.filter((n) => n.size < MINI_STREAM_CUTOFF);
	const large = withData.filter((n) => n.size >= MINI_STREAM_CUTOFF);

	// Small streams are packed into the mini stream, in 64-byte mini sectors.
	let miniSectors = 0;
	for (const n of small) {
		n.start = miniSectors;
		miniSectors += sectorsFor(n.size, MINI_SECTOR_SIZE);
	}
	const miniFatCount = sectorsFor(miniSectors * 4, SECTOR_SIZE);
	const miniStreamSectors = sectorsFor(miniSectors * MINI_SECTOR_SIZE, SECTOR_SIZE);
	const directoryCount = sectorsFor(entries.length, ENTRIES_PER_SECTOR);
	const largeSectors = large.reduce((s, n) => s + sectorsFor(n.size, SECTOR_SIZE), 0);
	const alloc = allocationSectors(directoryCount + miniFatCount + miniStreamSectors + largeSectors);

	// Sector order: FAT, DIFAT, directory, mini FAT, mini stream, large streams.
	const difatStart = alloc.fat;
	const directoryStart = difatStart + alloc.difat;
	const miniFatStart = directoryStart + directoryCount;
	const miniStreamStart = miniFatStart + miniFatCount;
	const layout: Layout = {
		fatCount: alloc.fat,
		difatStart,
		difatCount: alloc.difat,
		directoryStart,
		directoryCount,
		miniFatStart,
		miniFatCount,
		miniStreamStart,
		miniStreamSectors,
	};
	let next = miniStreamStart + miniStreamSectors;
	for (const n of large) {
		n.start = next;
		next += sectorsFor(n.size, SECTOR_SIZE);
	}
	const root = entries[0];
	root.start = miniSectors ? miniStreamStart : END_OF_CHAIN;
	root.size = miniSectors * MINI_SECTOR_SIZE;

	const bytes = new Uint8Array((next + 1) * SECTOR_SIZE);
	const view = new DataView(bytes.buffer);
	const at = (sector: number): number => (sector + 1) * SECTOR_SIZE;
	writeHeader(view, bytes, layout);
	writeAllocationTables(view, layout, large);
	for (let i = 0; i < directoryCount * ENTRIES_PER_SECTOR; i++) {
		writeDirectoryEntry(view, bytes, at(directoryStart) + i * DIRECTORY_ENTRY_SIZE, entries[i]);
	}

	// Mini FAT chains and the mini stream contents.
	const miniFat = at(miniFatStart);
	for (let i = 0; i < miniFatCount * IDS_PER_SECTOR; i++) {
		view.setUint32(miniFat + i * 4, FREE_SECTOR, true);
	}
	for (const n of small) {
		const count = sectorsFor(n.size, MINI_SECTOR_SIZE);
		for (let i = 0; i < count; i++) {
			const value = i === count - 1 ? END_OF_CHAIN : n.start + i + 1;
			view.setUint32(miniFat + (n.start + i) * 4, value, true);
		}
		bytes.set(n.data, at(miniStreamStart) + n.start * MINI_SECTOR_SIZE);
	}
	for (const n of large) {
		bytes.set(n.data, at(n.start));
	}
	return bytes;
}

/** FAT (chains of every sector run, FAT / DIFAT markers) and the DIFAT sectors. */
function writeAllocationTables(view: DataView, l: Layout, large: DirNode[]): void {
	const at = (sector: number): number => (sector + 1) * SECTOR_SIZE;
	const fat = new Uint32Array(l.fatCount * IDS_PER_SECTOR).fill(FREE_SECTOR);
	const chain = (start: number, count: number): void => {
		for (let i = 0; i < count; i++) {
			fat[start + i] = i === count - 1 ? END_OF_CHAIN : start + i + 1;
		}
	};
	fat.fill(FAT_SECTOR, 0, l.fatCount);
	fat.fill(DIFAT_SECTOR, l.difatStart, l.difatStart + l.difatCount);
	chain(l.directoryStart, l.directoryCount);
	chain(l.miniFatStart, l.miniFatCount);
	chain(l.miniStreamStart, l.miniStreamSectors);
	for (const n of large) {
		chain(n.start, sectorsFor(n.size, SECTOR_SIZE));
	}
	for (let i = 0; i < fat.length; i++) {
		view.setUint32(at(0) + i * 4, fat[i], true);
	}

	const perDifat = IDS_PER_SECTOR - 1;
	for (let d = 0; d < l.difatCount; d++) {
		const off = at(l.difatStart + d);
		for (let i = 0; i < perDifat; i++) {
			const fatSector = HEADER_DIFAT_ENTRIES + d * perDifat + i;
			view.setUint32(off + i * 4, fatSector < l.fatCount ? fatSector : FREE_SECTOR, true);
		}
		const nextDifat = d === l.difatCount - 1 ? END_OF_CHAIN : l.difatStart + d + 1;
		view.setUint32(off + SECTOR_SIZE - 4, nextDifat, true);
	}
}
