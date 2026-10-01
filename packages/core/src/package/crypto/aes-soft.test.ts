import { describe, expect, it } from 'vitest';

import { SoftwareAes } from './aes-soft';

const fromHex = (h: string) => Uint8Array.from(h.match(/../g) ?? [], (b) => parseInt(b, 16));
const hex = (b: Uint8Array) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');

function bytes(length: number, seed: number): Uint8Array<ArrayBuffer> {
	const out = new Uint8Array(length);
	let x = seed;
	for (let i = 0; i < length; i++) {
		x = (x * 1664525 + 1013904223) >>> 0;
		out[i] = x >>> 24;
	}
	return out;
}

describe('SoftwareAes', () => {
	const plain = fromHex('00112233445566778899aabbccddeeff');
	const zeroIv = new Uint8Array(16);
	// FIPS 197, appendix C.
	const vectors: Array<[string, string]> = [
		['000102030405060708090a0b0c0d0e0f', '69c4e0d86a7b0430d8cdb78070b4c55a'],
		['000102030405060708090a0b0c0d0e0f1011121314151617', 'dda97ca4864cdfe06eaf70a0ec0d7191'],
		[
			'000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f',
			'8ea2b7ca516745bfeafc49904b496089',
		],
	];

	it('matches the FIPS 197 known answers for 128, 192 and 256-bit keys', () => {
		for (const [key, cipher] of vectors) {
			const aes = new SoftwareAes(fromHex(key));
			expect(hex(aes.encryptCbc(zeroIv, plain))).toBe(cipher);
			expect(hex(aes.decryptCbc(zeroIv, fromHex(cipher)))).toBe(hex(plain));
		}
	});

	it('matches WebCrypto AES-CBC for multi-block data', async () => {
		for (const keyLength of [16, 24, 32]) {
			const key = bytes(keyLength, keyLength);
			const iv = bytes(16, 99);
			const data = bytes(4096 + 48, 7);
			const ours = new SoftwareAes(key).encryptCbc(iv, data);
			const k = await crypto.subtle.importKey('raw', key, 'AES-CBC', false, ['encrypt']);
			const theirs = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-CBC', iv }, k, data));
			// WebCrypto appends one block of PKCS#7 padding.
			expect(hex(ours)).toBe(hex(theirs.subarray(0, data.length)));
			expect(hex(new SoftwareAes(key).decryptCbc(iv, ours))).toBe(hex(data));
		}
	});

	it('rejects invalid key lengths', () => {
		expect(() => new SoftwareAes(new Uint8Array(20))).toThrow(/key length/);
	});
});
