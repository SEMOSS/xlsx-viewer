import { describe, expect, it } from 'vitest';

import { HASH_SIZE, hashSync, hmacSync, type HashAlgorithm } from './sha';

const ALGORITHMS: HashAlgorithm[] = ['SHA-1', 'SHA-256', 'SHA-384', 'SHA-512'];
const hex = (b: Uint8Array) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');
const text = (s: string) => new TextEncoder().encode(s);

function bytes(length: number, seed: number): Uint8Array {
	const out = new Uint8Array(length);
	let x = seed;
	for (let i = 0; i < length; i++) {
		x = (x * 1664525 + 1013904223) >>> 0;
		out[i] = x >>> 24;
	}
	return out;
}

async function subtleDigest(alg: HashAlgorithm, data: Uint8Array): Promise<string> {
	return hex(new Uint8Array(await crypto.subtle.digest(alg, new Uint8Array(data))));
}

async function subtleHmac(alg: HashAlgorithm, key: Uint8Array, data: Uint8Array): Promise<string> {
	const k = await crypto.subtle.importKey(
		'raw',
		new Uint8Array(key),
		{ name: 'HMAC', hash: alg },
		false,
		['sign'],
	);
	return hex(new Uint8Array(await crypto.subtle.sign('HMAC', k, new Uint8Array(data))));
}

describe('hashSync', () => {
	it('matches the FIPS 180 "abc" vectors', () => {
		const abc = text('abc');
		expect(hex(hashSync('SHA-1', abc))).toBe('a9993e364706816aba3e25717850c26c9cd0d89d');
		expect(hex(hashSync('SHA-256', abc))).toBe(
			'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
		);
		expect(hex(hashSync('SHA-384', abc))).toBe(
			'cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed' +
				'8086072ba1e7cc2358baeca134c825a7',
		);
		expect(hex(hashSync('SHA-512', abc))).toBe(
			'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a' +
				'2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
		);
	});

	it('matches WebCrypto for every length around the block boundaries', async () => {
		for (const alg of ALGORITHMS) {
			for (let length = 0; length <= 300; length++) {
				const data = bytes(length, length + 1);
				const ours = hex(hashSync(alg, data));
				expect(ours.length).toBe(HASH_SIZE[alg] * 2);
				expect(ours, `${alg} of ${length} bytes`).toBe(await subtleDigest(alg, data));
			}
		}
	});

	it('matches WebCrypto for a large input', async () => {
		const data = bytes(1_000_003, 7);
		for (const alg of ALGORITHMS) {
			expect(hex(hashSync(alg, data))).toBe(await subtleDigest(alg, data));
		}
	});
});

describe('hmacSync', () => {
	it('matches the RFC 2202 / RFC 4231 "Jefe" vectors', () => {
		const key = text('Jefe');
		const data = text('what do ya want for nothing?');
		expect(hex(hmacSync('SHA-1', key, data))).toBe('effcdf6ae5eb2fa2d27416d5f184df9c259a7c79');
		expect(hex(hmacSync('SHA-256', key, data))).toBe(
			'5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843',
		);
	});

	it('matches WebCrypto for short and long keys', async () => {
		for (const alg of ALGORITHMS) {
			for (const keyLength of [1, 20, 64, 128, 200]) {
				const key = bytes(keyLength, keyLength + 3);
				const data = bytes(5000, keyLength + 9);
				expect(hex(hmacSync(alg, key, data))).toBe(await subtleHmac(alg, key, data));
			}
		}
	});
});
