import { describe, expect, it } from 'vitest';

import {
	buildAgileEncryptionInfo,
	parseEncryptionInfo,
	type AgileEncryptionInfo,
} from './encryption-info';

const bytes = (n: number, fill: number) => new Uint8Array(n).fill(fill);

function sampleAgile(): AgileEncryptionInfo {
	const params = {
		saltSize: 16,
		blockSize: 16,
		keyBits: 256,
		hashSize: 64,
		cipherAlgorithm: 'AES',
		cipherChaining: 'ChainingModeCBC',
		hashAlgorithm: 'SHA512',
	};
	return {
		type: 'agile',
		version: '4.4',
		keyData: { ...params, saltValue: bytes(16, 1) },
		dataIntegrity: { encryptedHmacKey: bytes(64, 2), encryptedHmacValue: bytes(64, 3) },
		passwordKey: {
			...params,
			saltValue: bytes(16, 4),
			spinCount: 100000,
			encryptedVerifierHashInput: bytes(16, 5),
			encryptedVerifierHashValue: bytes(64, 6),
			encryptedKeyValue: bytes(32, 7),
		},
	};
}

describe('agile EncryptionInfo', () => {
	it('round-trips through build and parse', () => {
		const info = sampleAgile();
		const stream = buildAgileEncryptionInfo(info);
		expect(Array.from(stream.subarray(0, 8))).toEqual([4, 0, 4, 0, 0x40, 0, 0, 0]);
		expect(parseEncryptionInfo(stream)).toEqual(info);
	});

	it('writes the descriptor as Excel does', () => {
		const xml = new TextDecoder().decode(buildAgileEncryptionInfo(sampleAgile()).subarray(8));
		expect(xml).toContain(
			'<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ' +
				'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" ' +
				'xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">',
		);
		expect(xml).toContain('<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16"');
		expect(xml).toMatch(/<keyData [^>]*hashAlgorithm="SHA512" saltValue="AQEB[^"]*"\/>/);
	});

	it('tolerates a byte order mark, trailing padding and a missing integrity block', () => {
		const stream = buildAgileEncryptionInfo(sampleAgile());
		const xml = new TextDecoder().decode(stream.subarray(8)).replace(/<dataIntegrity[^>]*\/>/, '');
		const body = new TextEncoder().encode(`\uFEFF${xml}\u0000\u0000\r\n`);
		const padded = new Uint8Array(8 + body.length);
		padded.set(stream.subarray(0, 8));
		padded.set(body, 8);
		const info = parseEncryptionInfo(padded);
		expect(info.type === 'agile' && info.dataIntegrity).toBe(undefined);
		expect(info.type === 'agile' && info.passwordKey?.spinCount).toBe(100000);
	});

	it('reports broken descriptors as damaged', () => {
		const header = [4, 0, 4, 0, 0x40, 0, 0, 0];
		const stream = (xml: string) => Uint8Array.from([...header, ...new TextEncoder().encode(xml)]);
		expect(() => parseEncryptionInfo(stream('<encryption'))).toThrow(/not valid XML/);
		expect(() => parseEncryptionInfo(stream('<encryption/>'))).toThrow(/no key data/);
		expect(() =>
			parseEncryptionInfo(stream('<encryption><keyData saltSize="x"/></encryption>')),
		).toThrow(/saltSize/);
	});
});

describe('parseEncryptionInfo versions', () => {
	const versioned = (major: number, minor: number) => Uint8Array.of(major, 0, minor, 0, 0, 0, 0, 0);

	it('rejects extensible, unknown and truncated streams as unsupported', () => {
		for (const stream of [versioned(3, 3), versioned(4, 3), versioned(1, 1), new Uint8Array(4)]) {
			expect(() => parseEncryptionInfo(stream)).toThrow(
				expect.objectContaining({ reason: 'unsupported' }),
			);
		}
	});

	it('reports truncated standard headers as damaged', () => {
		const header = Uint8Array.of(4, 0, 2, 0, 0x24, 0, 0, 0, 200, 0, 0, 0);
		expect(() => parseEncryptionInfo(header)).toThrow(/truncated/);
	});
});
