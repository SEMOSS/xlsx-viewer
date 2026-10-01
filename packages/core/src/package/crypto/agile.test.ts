import { describe, expect, it } from 'vitest';

import {
	agileHashAlgorithm,
	agileSecretKey,
	checkCipherParams,
	decryptAgilePackage,
	encryptedPackageSize,
	spinPasswordHash,
} from './agile';
import { parseEncryptionInfo, type AgileEncryptionInfo } from './encryption-info';
import { base64Decode, concatBytes, digest, uint32LE, utf16le } from './primitives';

const hex = (b: Uint8Array) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');

/**
 * Known answer from a third-party encrypted presentation (password "123"):
 * its EncryptionInfo descriptor (AES-128, SHA-1, 100000 spins) and the first
 * four cipher blocks of its EncryptedPackage stream.
 */
const FIXTURE_XML =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
	'<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ' +
	'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
	'<keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20" ' +
	'cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1" ' +
	'saltValue="Z0WLa8YjezJpmDxkc0gzZg=="/>' +
	'<dataIntegrity encryptedHmacKey="jtEDggaAsZ2GVrqW0+JLX70gBB+9HsCTSoJ4kbP6fkE=" ' +
	'encryptedHmacValue="MvGJa/rTUMaTeKUIbVz5A/zvcEHWuwN2mwP4fmn2rPs="/>' +
	'<keyEncryptors>' +
	'<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
	'<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="128" ' +
	'hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1" ' +
	'saltValue="/1xJGUqU6CrsWFViKR+OIw==" ' +
	'encryptedVerifierHashInput="4Ye0dpAo1rxV6MOtlvxpng==" ' +
	'encryptedVerifierHashValue="q89qubgmn/KOFH7i4Ck61DViwT/FouVHmt+vaCBFJZA=" ' +
	'encryptedKeyValue="EiQpj7GIC4NP38c6nnox9A=="/></keyEncryptor></keyEncryptors></encryption>';
const FIXTURE_BLOCKS =
	'TumuzsG4N04D2wdYSJ5YKWSO2XMzbSj25bCHu51H5LCJ8GR8HHdBzrKJ0ADR1HL7FJr2z37243P1hrwe+YWMcA==';
/** ZIP local file header of `_rels/.rels`, the first part of the package. */
const FIXTURE_PLAIN_PREFIX =
	'504b0304140000000800328fe95cae836a37da000000400200000b0014005f72656c732f2e72656c73';

function fixtureInfo(): AgileEncryptionInfo {
	const body = new TextEncoder().encode(FIXTURE_XML);
	const info = parseEncryptionInfo(concatBytes(Uint8Array.of(4, 0, 4, 0, 0x40, 0, 0, 0), body));
	if (info.type !== 'agile') {
		throw new Error('expected agile');
	}
	return info;
}

describe('agile decryption of a third-party file', () => {
	it('verifies the password and decrypts the first package segment', async () => {
		const info = fixtureInfo();
		expect(info.passwordKey?.spinCount).toBe(100000);
		const key = await agileSecretKey(info, '123');
		expect(key && hex(key)).toBe('46e145757c005f51c262d05b54082012');
		const stream = concatBytes(uint32LE(64), uint32LE(0), base64Decode(FIXTURE_BLOCKS));
		const plain = await decryptAgilePackage(info, key!, stream);
		expect(plain).toHaveLength(64);
		expect(hex(plain).startsWith(FIXTURE_PLAIN_PREFIX)).toBe(true);
	});

	it('rejects a wrong password', async () => {
		expect(await agileSecretKey(fixtureInfo(), '1234')).toBeUndefined();
		expect(await agileSecretKey(fixtureInfo(), '')).toBeUndefined();
	});
});

describe('agile key derivation', () => {
	it('spins H(iterator + hash) on top of H(salt + password)', async () => {
		const salt = Uint8Array.from({ length: 16 }, (_, i) => i);
		let expected = await digest('SHA-512', concatBytes(salt, utf16le('pässword')));
		for (let i = 0; i < 50; i++) {
			expected = await digest('SHA-512', concatBytes(uint32LE(i), expected));
		}
		expect(hex(spinPasswordHash('SHA-512', salt, 'pässword', 50))).toBe(hex(expected));
		expect(hex(spinPasswordHash('SHA-1', salt, 'x', 0))).toBe(
			hex(await digest('SHA-1', concatBytes(salt, utf16le('x')))),
		);
	});

	it('names hash algorithms as descriptors spell them', () => {
		expect(agileHashAlgorithm('SHA512')).toBe('SHA-512');
		expect(agileHashAlgorithm('SHA-1')).toBe('SHA-1');
		expect(agileHashAlgorithm('sha256')).toBe('SHA-256');
		expect(agileHashAlgorithm('SHA384')).toBe('SHA-384');
		expect(agileHashAlgorithm('MD5')).toBeUndefined();
		expect(agileHashAlgorithm('WHIRLPOOL')).toBeUndefined();
	});

	it('rejects ciphers, chaining modes and hashes it cannot decrypt', () => {
		const ok = fixtureInfo().keyData;
		expect(checkCipherParams(ok)).toBe('SHA-1');
		const reason = (patch: Partial<typeof ok>) => {
			try {
				checkCipherParams({ ...ok, ...patch });
				return 'accepted';
			} catch (e) {
				return (e as { reason?: string }).reason ?? (e as Error).name;
			}
		};
		expect(reason({ cipherAlgorithm: 'DES' })).toBe('unsupported');
		expect(reason({ cipherChaining: 'ChainingModeCFB' })).toBe('unsupported');
		expect(reason({ hashAlgorithm: 'MD5' })).toBe('unsupported');
		expect(reason({ keyBits: 64 })).toBe('unsupported');
		expect(reason({ blockSize: 8 })).toBe('unsupported');
		expect(reason({ hashSize: 64 })).toBe('XlsxFormatError');
	});
});

describe('encryptedPackageSize', () => {
	const stream = (low: number, high: number) =>
		concatBytes(uint32LE(low), uint32LE(high), new Uint8Array(32));

	it('reads the 8-byte little-endian size and checks it against the data', () => {
		expect(encryptedPackageSize(stream(20, 0))).toBe(20);
		// Garbage in the high word falls back to the low word.
		expect(encryptedPackageSize(stream(20, 7))).toBe(20);
		expect(() => encryptedPackageSize(stream(40, 0))).toThrow(/damaged/);
		expect(() => encryptedPackageSize(new Uint8Array(4))).toThrow(/truncated/);
	});
});
