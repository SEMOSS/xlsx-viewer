import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { crc32 } from './crc32';
import { ZipFormatError, ZipReader } from './zip-read';
import { writeZip } from './zip-write';

const enc = new TextEncoder();
const dec = new TextDecoder();
const big = enc.encode('<row>'.repeat(20000) + 'é ünïcode');

describe('zip', () => {
	it('computes CRC-32 like the reference value', () => {
		expect(crc32(enc.encode('123456789'))).toBe(0xcbf43926);
		expect(crc32(new Uint8Array(0))).toBe(0);
	});

	it('round-trips stored and deflated entries, including UTF-8 names', async () => {
		const bytes = await writeZip([
			{ name: '[Content_Types].xml', data: enc.encode('<Types/>') },
			{ name: 'xl/worksheets/sheet1.xml', data: big },
			{ name: 'xl/média/ünï.bin', data: new Uint8Array([1, 2, 3]) },
		]);
		const zip = new ZipReader(bytes);
		expect(zip.entries.map((e) => e.name)).toEqual([
			'[Content_Types].xml',
			'xl/worksheets/sheet1.xml',
			'xl/média/ünï.bin',
		]);
		expect(dec.decode(await zip.bytes('xl/worksheets/sheet1.xml'))).toBe(dec.decode(big));
		expect(zip.raw('xl/worksheets/sheet1.xml')?.method).toBe(8);
		expect(zip.raw('[Content_Types].xml')?.method).toBe(0);
		expect(await zip.bytes('xl/média/ünï.bin')).toEqual(new Uint8Array([1, 2, 3]));
		expect(await zip.bytes('missing')).toBeUndefined();
	});

	it('interoperates with JSZip in both directions', async () => {
		const ours = await writeZip([{ name: 'a.xml', data: big }]);
		const viaJsZip = await JSZip.loadAsync(ours);
		expect(await viaJsZip.file('a.xml')!.async('string')).toBe(dec.decode(big));
		const theirs = new JSZip();
		theirs.file('b.xml', big);
		theirs.file('dir/', null, { dir: true });
		const reader = new ZipReader(
			await theirs.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
		);
		expect(reader.entries.map((e) => e.name)).toEqual(['b.xml']);
		expect(dec.decode(await reader.bytes('b.xml'))).toBe(dec.decode(big));
	});

	it('copies raw entries unchanged and writes deterministic bytes', async () => {
		const first = await writeZip([{ name: 'a.xml', data: big }]);
		const raw = new ZipReader(first).raw('a.xml')!;
		const copied = await writeZip([{ name: 'a.xml', raw }]);
		expect(copied).toEqual(first);
		expect(await writeZip([{ name: 'a.xml', data: big }])).toEqual(first);
	});

	it('rejects data that is not a ZIP archive', () => {
		expect(() => new ZipReader(enc.encode('not a zip at all, sorry'))).toThrow(ZipFormatError);
	});
});
