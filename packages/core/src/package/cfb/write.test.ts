import { describe, expect, it } from 'vitest';

import { readCompoundFile } from './read';
import { compareCfbNames, writeCompoundFile } from './write';

const NO_STREAM = 0xffffffff;

function pattern(length: number, seed = 1): Uint8Array {
	const out = new Uint8Array(length);
	let x = seed;
	for (let i = 0; i < length; i++) {
		x = (x * 1103515245 + 12345) >>> 0;
		out[i] = x >>> 24;
	}
	return out;
}

/** Fast equality for large arrays (deep `toEqual` is slow on megabytes). */
function sameBytes(a: Uint8Array | undefined, b: Uint8Array): boolean {
	return !!a && a.length === b.length && a.every((v, i) => v === b[i]);
}

interface RawEntry {
	name: string;
	type: number;
	black: boolean;
	left: number;
	right: number;
	child: number;
}

/** Directory entries of a file written with the header DIFAT only (small files). */
function rawDirectory(bytes: Uint8Array): RawEntry[] {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const at = (sector: number) => (sector + 1) * 512;
	const fatSectors = view.getUint32(0x2c, true);
	const next = (sector: number) => {
		const fatSector = view.getUint32(0x4c + Math.floor(sector / 128) * 4, true);
		expect(Math.floor(sector / 128)).toBeLessThan(fatSectors);
		return view.getUint32(at(fatSector) + (sector % 128) * 4, true);
	};
	const out: RawEntry[] = [];
	for (let s = view.getUint32(0x30, true); s < 0xfffffffa; s = next(s)) {
		for (let e = at(s); e < at(s) + 512; e += 128) {
			const len = view.getUint16(e + 0x40, true);
			let name = '';
			for (let i = 0; i + 2 < len; i += 2) {
				name += String.fromCharCode(view.getUint16(e + i, true));
			}
			out.push({
				name,
				type: bytes[e + 0x42],
				black: bytes[e + 0x43] === 1,
				left: view.getUint32(e + 0x44, true),
				right: view.getUint32(e + 0x48, true),
				child: view.getUint32(e + 0x4c, true),
			});
		}
	}
	return out;
}

/** Check one storage's sibling tree: ordering, black root, no red-red, equal black height. */
function checkRedBlackTree(entries: RawEntry[], rootId: number): string[] {
	const names: string[] = [];
	const height = (id: number, parentRed: boolean): number => {
		if (id === NO_STREAM) {
			return 1;
		}
		const e = entries[id];
		expect(parentRed && !e.black, `red node ${e.name} has a red parent`).toBe(false);
		const left = height(e.left, !e.black);
		names.push(e.name);
		const right = height(e.right, !e.black);
		expect(left, `black heights differ below ${e.name}`).toBe(right);
		return left + (e.black ? 1 : 0);
	};
	if (rootId !== NO_STREAM) {
		expect(entries[rootId].black, 'tree root is black').toBe(true);
	}
	height(rootId, false);
	const sorted = [...names].sort(compareCfbNames);
	expect(names).toEqual(sorted);
	return names;
}

describe('writeCompoundFile', () => {
	it('round-trips mini-stream and regular streams through readCompoundFile', () => {
		const small = pattern(100, 1);
		const edge = pattern(4095, 2);
		const big = pattern(4096, 3);
		const larger = pattern(70_000, 4);
		const bytes = writeCompoundFile([
			{ path: 'Small', data: small },
			{ path: 'Edge', data: edge },
			{ path: 'Big', data: big },
			{ path: 'Larger', data: larger },
			{ path: 'Empty', data: new Uint8Array(0) },
		]);
		expect(bytes.length % 512).toBe(0);
		const view = new DataView(bytes.buffer);
		expect(view.getUint16(0x1a, true)).toBe(3);
		expect(view.getUint16(0x1e, true)).toBe(9);
		expect(view.getUint32(0x38, true)).toBe(4096);
		const cfb = readCompoundFile(bytes);
		expect(cfb.stream('Small')).toEqual(small);
		expect(cfb.stream('edge')).toEqual(edge);
		expect(cfb.stream('Big')).toEqual(big);
		expect(cfb.stream('Larger')).toEqual(larger);
		expect(cfb.stream('Empty')).toEqual(new Uint8Array(0));
		expect(cfb.entries.find((e) => e.path === 'Larger')?.size).toBe(70_000);
	});

	it('creates storages for path prefixes', () => {
		const version = pattern(76, 5);
		const primary = pattern(200, 6);
		const bytes = writeCompoundFile([
			{ path: 'EncryptionInfo', data: pattern(1000, 7) },
			{ path: '\u0006DataSpaces/Version', data: version },
			{ path: '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', data: pattern(64) },
			{
				path: '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary',
				data: primary,
			},
		]);
		const cfb = readCompoundFile(bytes);
		const kinds = Object.fromEntries(cfb.entries.map((e) => [e.path, e.type]));
		expect(kinds).toEqual({
			'': 'root',
			EncryptionInfo: 'stream',
			'\u0006DataSpaces': 'storage',
			'\u0006DataSpaces/Version': 'stream',
			'\u0006DataSpaces/DataSpaceInfo': 'storage',
			'\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace': 'stream',
			'\u0006DataSpaces/TransformInfo': 'storage',
			'\u0006DataSpaces/TransformInfo/StrongEncryptionTransform': 'storage',
			'\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary': 'stream',
		});
		expect(cfb.stream('\u0006dataspaces/version')).toEqual(version);
		expect(
			cfb.stream('\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary'),
		).toEqual(primary);
	});

	it('builds valid red-black sibling trees in CFB name order', () => {
		for (const count of [1, 2, 3, 4, 5, 6, 7, 8, 9, 15, 16, 17, 40]) {
			const names = Array.from({ length: count }, (_, i) => `S${'x'.repeat(i % 5)}${i}`);
			const bytes = writeCompoundFile(names.map((n) => ({ path: n, data: pattern(10) })));
			const dir = rawDirectory(bytes);
			expect(dir[0]).toMatchObject({ name: 'Root Entry', type: 5 });
			const seen = checkRedBlackTree(dir, dir[0].child);
			expect(seen).toHaveLength(count);
		}
		// Nested storages get their own trees.
		const bytes = writeCompoundFile([
			{ path: 'a/b1', data: pattern(1) },
			{ path: 'a/B2', data: pattern(1) },
			{ path: 'a/c', data: pattern(1) },
			{ path: 'zz', data: pattern(1) },
		]);
		const dir = rawDirectory(bytes);
		expect(checkRedBlackTree(dir, dir[0].child)).toEqual(['a', 'zz']);
		const storage = dir.find((e) => e.name === 'a')!;
		expect(storage.type).toBe(1);
		expect(checkRedBlackTree(dir, storage.child)).toEqual(['c', 'b1', 'B2']);
	});

	it('orders names by length, then case-insensitively', () => {
		expect(compareCfbNames('Zz', 'aaa')).toBeLessThan(0);
		expect(compareCfbNames('abc', 'ABD')).toBeLessThan(0);
		expect(compareCfbNames('EncryptionInfo', 'encryptioninfo')).toBe(0);
		expect(compareCfbNames('\u0006DataSpaces', 'EncryptionInfo')).toBeLessThan(0);
	});

	it('uses DIFAT sectors for large files (10 MB)', () => {
		const data = pattern(10 * 1024 * 1024, 9);
		const small = pattern(300, 10);
		const bytes = writeCompoundFile([
			{ path: 'EncryptedPackage', data },
			{ path: 'EncryptionInfo', data: small },
		]);
		const view = new DataView(bytes.buffer);
		expect(view.getUint32(0x2c, true)).toBeGreaterThan(109);
		expect(view.getUint32(0x48, true)).toBeGreaterThan(0);
		const cfb = readCompoundFile(bytes);
		expect(sameBytes(cfb.stream('EncryptedPackage'), data)).toBe(true);
		expect(cfb.stream('EncryptionInfo')).toEqual(small);
	});

	it('rejects invalid and duplicate names', () => {
		const data = pattern(4);
		expect(() => writeCompoundFile([{ path: 'x'.repeat(32), data }])).toThrow(/name/);
		expect(() => writeCompoundFile([{ path: 'a:b', data }])).toThrow(/name/);
		expect(() => writeCompoundFile([{ path: 'a//b', data }])).toThrow(/name/);
		expect(() =>
			writeCompoundFile([
				{ path: 'Same', data },
				{ path: 'SAME', data },
			]),
		).toThrow(/Duplicate/);
		expect(() =>
			writeCompoundFile([
				{ path: 'Stream', data },
				{ path: 'Stream/Child', data },
			]),
		).toThrow(/below a stream/);
	});

	it('writes an empty compound file', () => {
		const cfb = readCompoundFile(writeCompoundFile([]));
		expect(cfb.entries.map((e) => e.type)).toEqual(['root']);
	});
});
