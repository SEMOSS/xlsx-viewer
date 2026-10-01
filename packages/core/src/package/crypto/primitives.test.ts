import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	aesCbc,
	aesEcbDecrypt,
	base64Decode,
	base64Encode,
	digest,
	equalBytes,
	fitBytes,
	hmac,
	padToBlock,
	randomBytes,
	utf16le,
} from './primitives';

const hex = (b: Uint8Array) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('aesCbc', () => {
	it('encrypts and decrypts whole blocks without padding, in WebCrypto or software', async () => {
		for (const length of [16, 24, 32]) {
			const key = randomBytes(length);
			const iv = randomBytes(16);
			const data = randomBytes(4096 + 32);
			const web = await aesCbc(key);
			const soft = await aesCbc(key, true);
			const encrypted = await web.encrypt(iv, data);
			expect(encrypted).toHaveLength(data.length);
			expect(hex(await soft.encrypt(iv, data))).toBe(hex(encrypted));
			expect(hex(await web.decrypt(iv, encrypted))).toBe(hex(data));
			expect(hex(await soft.decrypt(iv, encrypted))).toBe(hex(data));
			expect(await web.decrypt(iv, new Uint8Array(0))).toHaveLength(0);
		}
	});

	it('decrypts ECB through one CBC pass', async () => {
		const key = randomBytes(16);
		const cipher = await aesCbc(key);
		const blocks = [randomBytes(16), randomBytes(16), randomBytes(16)];
		// ECB encryption is CBC with a zero IV, one block at a time.
		const encrypted = new Uint8Array(48);
		for (let i = 0; i < 3; i++) {
			encrypted.set(await cipher.encrypt(new Uint8Array(16), blocks[i]), i * 16);
		}
		const plain = await aesEcbDecrypt(cipher, encrypted);
		expect(hex(plain)).toBe(blocks.map(hex).join(''));
	});

	it('rejects invalid key sizes', async () => {
		await expect(aesCbc(new Uint8Array(20))).rejects.toThrow(/key length/);
	});
});

describe('digest and hmac', () => {
	it('fall back to the TypeScript hashes without WebCrypto', async () => {
		const data = randomBytes(1000);
		const key = randomBytes(64);
		const viaWeb = [await digest('SHA-512', data), await hmac('SHA-256', key, data)].map(hex);
		vi.stubGlobal('crypto', { getRandomValues: (a: Uint8Array) => a });
		const viaTs = [await digest('SHA-512', data), await hmac('SHA-256', key, data)].map(hex);
		expect(viaTs).toEqual(viaWeb);
		// The software cipher takes over as well.
		const cipher = await aesCbc(new Uint8Array(32));
		expect(await cipher.encrypt(new Uint8Array(16), new Uint8Array(16))).toHaveLength(16);
	});
});

describe('byte helpers', () => {
	it('encode passwords as UTF-16LE and base64 round-trips', () => {
		expect(Array.from(utf16le('A€'))).toEqual([0x41, 0, 0xac, 0x20]);
		expect(base64Encode(Uint8Array.of(0x48, 0x65, 0x6c, 0x6c, 0x6f))).toBe('SGVsbG8=');
		expect(Array.from(base64Decode('SGVs\r\nbG8='))).toEqual([0x48, 0x65, 0x6c, 0x6c, 0x6f]);
		expect(base64Encode(new Uint8Array(0))).toBe('');
	});

	it('fit, pad and compare', () => {
		expect(Array.from(fitBytes(Uint8Array.of(1, 2), 4, 0x36))).toEqual([1, 2, 0x36, 0x36]);
		expect(Array.from(fitBytes(Uint8Array.of(1, 2, 3), 2, 0))).toEqual([1, 2]);
		expect(padToBlock(new Uint8Array(17), 16)).toHaveLength(32);
		expect(padToBlock(new Uint8Array(32), 16)).toHaveLength(32);
		expect(equalBytes(Uint8Array.of(1, 2), Uint8Array.of(1, 2))).toBe(true);
		expect(equalBytes(Uint8Array.of(1, 2), Uint8Array.of(1, 3))).toBe(false);
		expect(equalBytes(Uint8Array.of(1), Uint8Array.of(1, 2))).toBe(false);
	});
});
