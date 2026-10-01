import { describe, expect, it } from 'vitest';

import { isCompoundFile } from '../opc-package';
import { CfbFormatError, readCompoundFile } from './read';

const FREE = 0xffffffff;
const END = 0xfffffffe;
const FATSECT = 0xfffffffd;

/**
 * A version 3 compound file with a mini-stream stream ("Small"), a regular
 * stream ("Big", 5000 bytes) and a storage holding a stream ("Folder/Inner").
 * Sectors: 0 FAT, 1-2 directory, 3 mini FAT, 4 mini stream, 5-14 "Big".
 */
function sampleFile(): {
	bytes: Uint8Array;
	small: Uint8Array;
	big: Uint8Array;
	inner: Uint8Array;
} {
	const size = 512;
	const bytes = new Uint8Array(size * 16);
	const v = new DataView(bytes.buffer);
	bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
	v.setUint16(0x18, 0x3e, true);
	v.setUint16(0x1a, 3, true);
	v.setUint16(0x1c, 0xfffe, true);
	v.setUint16(0x1e, 9, true);
	v.setUint16(0x20, 6, true);
	v.setUint32(0x2c, 1, true);
	v.setUint32(0x30, 1, true);
	v.setUint32(0x38, 4096, true);
	v.setUint32(0x3c, 3, true);
	v.setUint32(0x40, 1, true);
	v.setUint32(0x44, END, true);
	for (let i = 0; i < 109; i++) {
		v.setUint32(0x4c + i * 4, i === 0 ? 0 : FREE, true);
	}
	const sector = (n: number) => (n + 1) * size;
	// FAT
	const fat = [FATSECT, 2, END, END, END, 6, 7, 8, 9, 10, 11, 12, 13, 14, END];
	for (let i = 0; i < size / 4; i++) {
		v.setUint32(sector(0) + i * 4, fat[i] ?? FREE, true);
	}
	// Mini FAT
	const miniFat = [1, END, END];
	for (let i = 0; i < size / 4; i++) {
		v.setUint32(sector(3) + i * 4, miniFat[i] ?? FREE, true);
	}
	// Directory
	const entry = (
		i: number,
		name: string,
		type: number,
		o: { left?: number; right?: number; child?: number; start?: number; size?: number },
	) => {
		const e = sector(1) + i * 128;
		for (let c = 0; c < name.length; c++) {
			v.setUint16(e + c * 2, name.charCodeAt(c), true);
		}
		v.setUint16(e + 0x40, (name.length + 1) * 2, true);
		bytes[e + 0x42] = type;
		v.setUint32(e + 0x44, o.left ?? FREE, true);
		v.setUint32(e + 0x48, o.right ?? FREE, true);
		v.setUint32(e + 0x4c, o.child ?? FREE, true);
		v.setUint32(e + 0x74, o.start ?? 0, true);
		v.setUint32(e + 0x78, o.size ?? 0, true);
	};
	entry(0, 'Root Entry', 5, { child: 2, start: 4, size: 192 });
	entry(1, 'Small', 2, { start: 0, size: 100 });
	entry(2, 'Big', 2, { left: 1, right: 3, start: 5, size: 5000 });
	entry(3, 'Folder', 1, { child: 4 });
	entry(4, 'Inner', 2, { start: 2, size: 10 });
	// Contents
	const small = Uint8Array.from({ length: 100 }, (_, i) => i + 1);
	const inner = Uint8Array.from({ length: 10 }, (_, i) => 200 + i);
	const big = Uint8Array.from({ length: 5000 }, (_, i) => (i * 7) % 251);
	bytes.set(small, sector(4));
	bytes.set(inner, sector(4) + 2 * 64);
	bytes.set(big, sector(5));
	return { bytes, small, big, inner };
}

describe('readCompoundFile', () => {
	it('lists storages and streams with their paths', () => {
		const { bytes } = sampleFile();
		expect(isCompoundFile(bytes)).toBe(true);
		const cfb = readCompoundFile(bytes);
		expect(cfb.entries.map((e) => `${e.type}:${e.path}`)).toEqual([
			'root:',
			'stream:Small',
			'stream:Big',
			'storage:Folder',
			'stream:Folder/Inner',
		]);
	});

	it('reads mini-stream and regular streams', () => {
		const { bytes, small, big, inner } = sampleFile();
		const cfb = readCompoundFile(bytes);
		expect(cfb.stream('Small')).toEqual(small);
		expect(cfb.stream('big')).toEqual(big);
		expect(cfb.stream('Folder/Inner')).toEqual(inner);
		expect(cfb.stream('Folder')).toBeUndefined();
		expect(cfb.stream('Missing')).toBeUndefined();
	});

	it('rejects non-compound data and looping chains', () => {
		expect(() => readCompoundFile(new Uint8Array(600))).toThrow(CfbFormatError);
		const { bytes } = sampleFile();
		// Make the "Big" chain loop back onto itself.
		new DataView(bytes.buffer).setUint32(512 + 14 * 4, 5, true);
		expect(() => readCompoundFile(bytes).stream('Big')).toThrow(/loops/);
	});
});
