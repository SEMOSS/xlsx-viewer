/**
 * Minimal compound file writer for tests ([MS-CFB] version 3, 512-byte
 * sectors). Streams under 4096 bytes go to the mini stream (64-byte mini
 * sectors), larger ones to regular sectors. Layout: FAT sectors, directory,
 * mini FAT, mini stream container, then the regular streams. All streams
 * sit directly under the root, chained as right siblings.
 */

const SECTOR = 512;
const MINI = 64;
const CUTOFF = 4096;
const FREE = 0xffffffff;
const END = 0xfffffffe;
const FATSECT = 0xfffffffd;

export interface CfbStream {
	name: string;
	data: Uint8Array;
}

const sectorsFor = (bytes: number, size: number): number => Math.ceil(bytes / size);

export function writeCompoundFileForTest(streams: readonly CfbStream[]): Uint8Array {
	const small = streams.filter((s) => s.data.length > 0 && s.data.length < CUTOFF);
	const big = streams.filter((s) => s.data.length >= CUTOFF);
	const miniCounts = small.map((s) => sectorsFor(s.data.length, MINI));
	const miniTotal = miniCounts.reduce((a, b) => a + b, 0);
	const dirSectors = sectorsFor((streams.length + 1) * 128, SECTOR);
	const miniFatSectors = sectorsFor(miniTotal * 4, SECTOR);
	const miniDataSectors = sectorsFor(miniTotal * MINI, SECTOR);
	const bigCounts = big.map((s) => sectorsFor(s.data.length, SECTOR));
	const others =
		dirSectors + miniFatSectors + miniDataSectors + bigCounts.reduce((a, b) => a + b, 0);
	let fatSectors = 1;
	while (fatSectors * (SECTOR / 4) < fatSectors + others) {
		fatSectors++;
	}
	if (fatSectors > 109) {
		throw new Error('test compound file too large');
	}
	const total = fatSectors + others;
	const out = new Uint8Array(SECTOR * (total + 1));
	const v = new DataView(out.buffer);
	const at = (sector: number): number => (sector + 1) * SECTOR;

	// FAT: chain helper.
	const fat = new Uint32Array(fatSectors * (SECTOR / 4)).fill(FREE);
	for (let i = 0; i < fatSectors; i++) {
		fat[i] = FATSECT;
	}
	let next = fatSectors;
	const allocate = (count: number): number => {
		if (count === 0) {
			return END;
		}
		const start = next;
		for (let i = 0; i < count; i++) {
			fat[start + i] = i + 1 < count ? start + i + 1 : END;
		}
		next += count;
		return start;
	};
	const dirStart = allocate(dirSectors);
	const miniFatStart = allocate(miniFatSectors);
	const miniDataStart = allocate(miniDataSectors);
	const bigStarts = bigCounts.map((n) => allocate(n));

	// Header.
	out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
	v.setUint16(0x18, 0x3e, true);
	v.setUint16(0x1a, 3, true);
	v.setUint16(0x1c, 0xfffe, true);
	v.setUint16(0x1e, 9, true);
	v.setUint16(0x20, 6, true);
	v.setUint32(0x2c, fatSectors, true);
	v.setUint32(0x30, dirStart, true);
	v.setUint32(0x38, CUTOFF, true);
	v.setUint32(0x3c, miniFatSectors ? miniFatStart : END, true);
	v.setUint32(0x40, miniFatSectors, true);
	v.setUint32(0x44, END, true);
	for (let i = 0; i < 109; i++) {
		v.setUint32(0x4c + i * 4, i < fatSectors ? i : FREE, true);
	}
	fat.forEach((entry, i) => v.setUint32(at(Math.floor(i / 128)) + (i % 128) * 4, entry, true));

	// Mini FAT and mini stream data.
	let mini = 0;
	const miniStarts = small.map((s, k) => {
		const start = mini;
		for (let i = 0; i < miniCounts[k]; i++) {
			const entry = i + 1 < miniCounts[k] ? start + i + 1 : END;
			v.setUint32(at(miniFatStart) + (start + i) * 4, entry, true);
		}
		out.set(s.data, at(miniDataStart) + start * MINI);
		mini += miniCounts[k];
		return start;
	});
	for (let i = mini; i < miniFatSectors * (SECTOR / 4); i++) {
		v.setUint32(at(miniFatStart) + i * 4, FREE, true);
	}
	big.forEach((s, k) => out.set(s.data, at(bigStarts[k])));

	// Directory: root, then the streams chained through their right siblings.
	const entry = (
		index: number,
		name: string,
		type: number,
		start: number,
		size: number,
		right: number,
		child: number,
	): void => {
		const e = at(dirStart) + index * 128;
		for (let c = 0; c < name.length; c++) {
			v.setUint16(e + c * 2, name.charCodeAt(c), true);
		}
		v.setUint16(e + 0x40, (name.length + 1) * 2, true);
		out[e + 0x42] = type;
		out[e + 0x43] = 1;
		v.setUint32(e + 0x44, FREE, true);
		v.setUint32(e + 0x48, right, true);
		v.setUint32(e + 0x4c, child, true);
		v.setUint32(e + 0x74, start, true);
		v.setUint32(e + 0x78, size, true);
	};
	entry(
		0,
		'Root Entry',
		5,
		miniTotal ? miniDataStart : END,
		miniTotal * MINI,
		FREE,
		streams.length ? 1 : FREE,
	);
	streams.forEach((s, i) => {
		const smallIndex = small.indexOf(s);
		const bigIndex = big.indexOf(s);
		const start =
			smallIndex >= 0 ? miniStarts[smallIndex] : bigIndex >= 0 ? bigStarts[bigIndex] : END;
		entry(i + 1, s.name, 2, start, s.data.length, i + 1 < streams.length ? i + 2 : FREE, FREE);
	});
	for (let i = streams.length + 1; i < dirSectors * 4; i++) {
		const e = at(dirStart) + i * 128;
		v.setUint32(e + 0x44, FREE, true);
		v.setUint32(e + 0x48, FREE, true);
		v.setUint32(e + 0x4c, FREE, true);
	}
	return out;
}
