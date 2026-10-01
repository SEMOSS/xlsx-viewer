import { describe, expect, it } from 'vitest';

import { XlsxFormatError, XlsxPasswordError } from '../../read/errors';
import { readCompoundFile } from '../cfb/read';
import { writeCompoundFile } from '../cfb/write';
import {
	decryptPackage,
	encryptionInfoOf,
	encryptPackage,
	isEncryptedPackageFile,
} from './package-crypto';
import type { HashAlgorithm } from './sha';

const FAST = { spinCount: 1000 };

function zipLike(length: number, seed = 1): Uint8Array {
	const out = new Uint8Array(length);
	let x = seed;
	for (let i = 0; i < length; i++) {
		x = (x * 1103515245 + 12345) >>> 0;
		out[i] = x >>> 24;
	}
	out.set([0x50, 0x4b, 0x03, 0x04]);
	return out;
}

const sameBytes = (a: Uint8Array, b: Uint8Array) =>
	a.length === b.length && a.every((v, i) => v === b[i]);

/** Rewrite the streams of a compound file. */
function rebuild(file: Uint8Array, patch: (streams: Map<string, Uint8Array>) => void): Uint8Array {
	const cfb = readCompoundFile(file);
	const streams = new Map<string, Uint8Array>();
	for (const e of cfb.entries) {
		if (e.type === 'stream') {
			streams.set(e.path, cfb.stream(e.path)!.slice());
		}
	}
	patch(streams);
	return writeCompoundFile([...streams].map(([path, data]) => ({ path, data })));
}

/** Edit the XML descriptor of an agile EncryptionInfo stream. */
function editDescriptor(file: Uint8Array, edit: (xml: string) => string): Uint8Array {
	return rebuild(file, (streams) => {
		const info = streams.get('EncryptionInfo')!;
		const xml = edit(new TextDecoder().decode(info.subarray(8)));
		const body = new TextEncoder().encode(xml);
		const out = new Uint8Array(8 + body.length);
		out.set(info.subarray(0, 8));
		out.set(body, 8);
		streams.set('EncryptionInfo', out);
	});
}

async function failure(promise: Promise<unknown>): Promise<Error> {
	try {
		await promise;
	} catch (e) {
		return e as Error;
	}
	throw new Error('expected a rejection');
}

describe('encryptPackage / decryptPackage', () => {
	it('round-trips every supported key size and hash', async () => {
		const zip = zipLike(9000);
		const hashes: HashAlgorithm[] = ['SHA-1', 'SHA-256', 'SHA-384', 'SHA-512'];
		for (const keyBits of [128, 192, 256] as const) {
			for (const hashAlgorithm of hashes) {
				const file = await encryptPackage(zip, 'Pa$$w0rd', { ...FAST, keyBits, hashAlgorithm });
				const info = encryptionInfoOf(file);
				expect(info).toMatchObject({ type: 'agile', keyData: { keyBits } });
				expect(await decryptPackage(file, 'Pa$$w0rd'), `${keyBits} ${hashAlgorithm}`).toEqual(zip);
			}
		}
	});

	it('writes what Excel 2016 writes by default', async () => {
		const zip = zipLike(5000);
		const file = await encryptPackage(zip, 'secret');
		const info = encryptionInfoOf(file);
		if (info?.type !== 'agile') {
			throw new Error('expected agile encryption');
		}
		const params = {
			saltSize: 16,
			blockSize: 16,
			keyBits: 256,
			hashSize: 64,
			cipherAlgorithm: 'AES',
			cipherChaining: 'ChainingModeCBC',
			hashAlgorithm: 'SHA512',
		};
		expect(info.version).toBe('4.4');
		expect(info.keyData).toMatchObject(params);
		expect(info.passwordKey).toMatchObject({ ...params, spinCount: 100000 });
		expect(info.dataIntegrity?.encryptedHmacKey).toHaveLength(64);
		expect(info.dataIntegrity?.encryptedHmacValue).toHaveLength(64);

		const cfb = readCompoundFile(file);
		expect(cfb.entries.map((e) => `${e.type}:${e.path}`).sort()).toEqual(
			[
				'root:',
				'storage:\u0006DataSpaces',
				'stream:\u0006DataSpaces/Version',
				'stream:\u0006DataSpaces/DataSpaceMap',
				'storage:\u0006DataSpaces/DataSpaceInfo',
				'stream:\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace',
				'storage:\u0006DataSpaces/TransformInfo',
				'storage:\u0006DataSpaces/TransformInfo/StrongEncryptionTransform',
				'stream:\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary',
				'stream:EncryptionInfo',
				'stream:EncryptedPackage',
			].sort(),
		);
		const infoStream = cfb.stream('EncryptionInfo')!;
		expect(Array.from(infoStream.subarray(0, 8))).toEqual([4, 0, 4, 0, 0x40, 0, 0, 0]);
		const text = new TextDecoder().decode(infoStream.subarray(8));
		expect(text).toMatch(/^<\?xml version="1.0" encoding="UTF-8" standalone="yes"\?>\r\n/);
		// 8-byte size, then the data padded to the AES block size.
		const stream = cfb.stream('EncryptedPackage')!;
		const view = new DataView(stream.buffer, stream.byteOffset, 8);
		expect([view.getUint32(0, true), view.getUint32(4, true)]).toEqual([5000, 0]);
		expect(stream.length).toBe(8 + 5008);
		expect(await decryptPackage(file, 'secret')).toEqual(zip);
	});

	it('handles packages of several segments and empty or unicode passwords', async () => {
		const zip = zipLike(3 * 1024 * 1024 + 7, 5);
		const file = await encryptPackage(zip, '', FAST);
		expect(sameBytes(await decryptPackage(file, ''), zip)).toBe(true);
		const unicode = 'éàü世界 🔒';
		const small = zipLike(10);
		const locked = await encryptPackage(small, unicode, FAST);
		expect(await decryptPackage(locked, unicode)).toEqual(small);
	});

	it('reports a missing or wrong password', async () => {
		const file = await encryptPackage(zipLike(100), 'right', FAST);
		const missing = await failure(decryptPackage(file));
		expect(missing).toBeInstanceOf(XlsxPasswordError);
		expect((missing as XlsxPasswordError).reason).toBe('required');
		const wrong = await failure(decryptPackage(file, 'Right'));
		expect(wrong).toBeInstanceOf(XlsxPasswordError);
		expect((wrong as XlsxPasswordError).reason).toBe('incorrect');
	});

	it('detects tampering through the data integrity HMAC', async () => {
		const file = await encryptPackage(zipLike(10_000), 'pw', FAST);
		const flipped = rebuild(file, (streams) => {
			streams.get('EncryptedPackage')![5000] ^= 1;
		});
		const err = await failure(decryptPackage(flipped, 'pw'));
		expect(err).toBeInstanceOf(XlsxFormatError);
		expect(err.message).toMatch(/integrity check/);
		// The size prefix is covered too.
		const resized = rebuild(file, (streams) => {
			streams.get('EncryptedPackage')![0] ^= 1;
		});
		await expect(decryptPackage(resized, 'pw')).rejects.toThrow(/integrity check/);
		const forged = editDescriptor(file, (xml) =>
			xml.replace(/encryptedHmacValue="([^"]+)"/, (_, v: string) => {
				const b = Uint8Array.from(atob(v), (c) => c.charCodeAt(0));
				b[0] ^= 1;
				return `encryptedHmacValue="${btoa(String.fromCharCode(...b))}"`;
			}),
		);
		await expect(decryptPackage(forged, 'pw')).rejects.toThrow(/integrity check/);
	});

	it('reports unsupported agile parameters before asking for a password', async () => {
		const file = await encryptPackage(zipLike(100), 'pw', FAST);
		const variants = [
			(xml: string) => xml.replace(/cipherAlgorithm="AES"/g, 'cipherAlgorithm="DES"'),
			(xml: string) => xml.replace(/ChainingModeCBC/g, 'ChainingModeCFB'),
			(xml: string) => xml.replace(/hashAlgorithm="SHA512"/g, 'hashAlgorithm="MD5"'),
			// Only a certificate key encryptor.
			(xml: string) => xml.replace(/uri="([^"]+)\/password"/, 'uri="$1/certificate"'),
		];
		for (const edit of variants) {
			const err = await failure(decryptPackage(editDescriptor(file, edit)));
			expect(err).toBeInstanceOf(XlsxPasswordError);
			expect((err as XlsxPasswordError).reason).toBe('unsupported');
		}
		const extensible = rebuild(file, (streams) => {
			streams.set('EncryptionInfo', Uint8Array.of(4, 0, 3, 0, 0x10, 0, 0, 0, 0, 0, 0, 0));
		});
		await expect(decryptPackage(extensible, 'pw')).rejects.toMatchObject({
			reason: 'unsupported',
			message: expect.stringMatching(/extensible/),
		});
		const unknown = rebuild(file, (streams) => {
			streams.set('EncryptionInfo', Uint8Array.of(9, 0, 9, 0, 0, 0, 0, 0));
		});
		await expect(decryptPackage(unknown)).rejects.toMatchObject({ reason: 'unsupported' });
	});

	it('reports containers without the encryption streams as unsupported', async () => {
		const empty = writeCompoundFile([{ path: 'Workbook', data: new Uint8Array(10) }]);
		await expect(decryptPackage(empty, 'pw')).rejects.toMatchObject({ reason: 'unsupported' });
		await expect(decryptPackage(new Uint8Array(600), 'pw')).rejects.toMatchObject({
			reason: 'unsupported',
		});
	});
});

describe('isEncryptedPackageFile / encryptionInfoOf', () => {
	it('recognizes encrypted packages only', async () => {
		const file = await encryptPackage(zipLike(100), 'pw', FAST);
		expect(isEncryptedPackageFile(file)).toBe(true);
		expect(isEncryptedPackageFile(file.buffer as ArrayBuffer)).toBe(true);
		expect(isEncryptedPackageFile(zipLike(100))).toBe(false);
		expect(isEncryptedPackageFile(writeCompoundFile([{ path: 'Book', data: zipLike(9) }]))).toBe(
			false,
		);
		expect(encryptionInfoOf(zipLike(100))).toBeUndefined();
		expect(encryptionInfoOf(file)?.type).toBe('agile');
	});
});
