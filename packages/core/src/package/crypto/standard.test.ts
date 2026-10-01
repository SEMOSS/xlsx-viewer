import { createCipheriv, createHash, randomBytes } from 'node:crypto';
import { describe, expect, it } from 'vitest';

import { writeCompoundFile } from '../cfb/write';
import { decryptPackage, encryptionInfoOf } from './package-crypto';

const CSP_NAME = 'Microsoft Enhanced RSA and AES Cryptographic Provider';
const ALG_IDS: Record<number, number> = { 128: 0x660e, 192: 0x660f, 256: 0x6610 };

const sha1 = (...parts: Uint8Array[]) => {
	const h = createHash('sha1');
	parts.forEach((p) => h.update(p));
	return new Uint8Array(h.digest());
};
const u32 = (v: number) => {
	const b = new Uint8Array(4);
	new DataView(b.buffer).setUint32(0, v, true);
	return b;
};
const utf16 = (s: string) => new Uint8Array(Buffer.from(s, 'utf16le'));

interface StandardOptions {
	keyBits?: 128 | 192 | 256;
	major?: number;
	algId?: number;
	flags?: number;
}

/**
 * Encrypt a package with ECMA-376 Standard encryption, written independently
 * from the decryptor and straight from [MS-OFFCRYPTO] 2.3.4.5 to 2.3.4.9 with
 * node:crypto: SHA-1 key derivation (50000 spins, X1 / X2 expansion) and AES-ECB.
 */
function standardEncrypt(zip: Uint8Array, password: string, options: StandardOptions = {}) {
	const keyBits = options.keyBits ?? 128;
	const flags = options.flags ?? 0x24;
	const salt = new Uint8Array(randomBytes(16));
	let h = sha1(salt, utf16(password));
	for (let i = 0; i < 50000; i++) {
		h = sha1(u32(i), h);
	}
	const final = sha1(h, u32(0));
	const expand = (fill: number) => {
		const buf = new Uint8Array(64).fill(fill);
		final.forEach((b, i) => (buf[i] ^= b));
		return sha1(buf);
	};
	const key = Buffer.concat([expand(0x36), expand(0x5c)]).subarray(0, keyBits / 8);
	const ecb = (data: Uint8Array) => {
		const c = createCipheriv(`aes-${keyBits}-ecb`, key, null).setAutoPadding(false);
		return new Uint8Array(Buffer.concat([c.update(data), c.final()]));
	};
	const verifier = new Uint8Array(randomBytes(16));
	const verifierHash = new Uint8Array(32);
	verifierHash.set(sha1(verifier));

	const csp = utf16(`${CSP_NAME}\u0000`);
	const header = Buffer.concat([
		u32(flags),
		u32(0),
		u32(options.algId ?? ALG_IDS[keyBits]),
		u32(0x8004),
		u32(keyBits),
		u32(0x18),
		u32(0),
		u32(0),
		csp,
	]);
	const info = Buffer.concat([
		Uint8Array.of(options.major ?? 4, 0, 2, 0),
		u32(flags),
		u32(header.length),
		header,
		u32(16),
		salt,
		ecb(verifier),
		u32(20),
		ecb(verifierHash),
	]);
	const padded = new Uint8Array(Math.ceil(zip.length / 16) * 16);
	padded.set(zip);
	const encryptedPackage = Buffer.concat([u32(zip.length), u32(0), ecb(padded)]);
	return writeCompoundFile([
		{ path: 'EncryptionInfo', data: new Uint8Array(info) },
		{ path: 'EncryptedPackage', data: new Uint8Array(encryptedPackage) },
	]);
}

const zipLike = (length: number) =>
	Uint8Array.from({ length }, (_, i) => (i < 4 ? [0x50, 0x4b, 3, 4][i] : (i * 31) % 251));

describe('Standard encryption (Excel 2007)', () => {
	it('decrypts AES-128, AES-192 and AES-256 packages', async () => {
		for (const keyBits of [128, 192, 256] as const) {
			const zip = zipLike(5000 + keyBits);
			const file = standardEncrypt(zip, 'Excel 2007', { keyBits });
			const info = encryptionInfoOf(file);
			expect(info).toMatchObject({ type: 'standard', version: '4.2', keyBits, cspName: CSP_NAME });
			expect(await decryptPackage(file, 'Excel 2007')).toEqual(zip);
		}
	});

	it('reads version 3.2 headers and the default algorithm id', async () => {
		const zip = zipLike(100);
		const file = standardEncrypt(zip, 'pw', { major: 3, algId: 0 });
		expect(encryptionInfoOf(file)?.version).toBe('3.2');
		expect(await decryptPackage(file, 'pw')).toEqual(zip);
	});

	it('reports wrong and missing passwords', async () => {
		const file = standardEncrypt(zipLike(64), 'right');
		await expect(decryptPackage(file, 'wrong')).rejects.toMatchObject({ reason: 'incorrect' });
		await expect(decryptPackage(file)).rejects.toMatchObject({ reason: 'required' });
	});

	it('reports RC4 and extensible encryption as unsupported, before any password', async () => {
		const rc4 = standardEncrypt(zipLike(64), 'pw', { algId: 0x6801, flags: 0x04 });
		await expect(decryptPackage(rc4)).rejects.toMatchObject({ reason: 'unsupported' });
		const external = standardEncrypt(zipLike(64), 'pw', { flags: 0x34 });
		await expect(decryptPackage(external, 'pw')).rejects.toMatchObject({
			reason: 'unsupported',
			message: expect.stringMatching(/extensible/),
		});
	});
});
