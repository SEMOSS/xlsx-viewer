import { describe, expect, it } from 'vitest';

import { dataSpacesStreams } from './data-spaces';

const hex = (b: Uint8Array) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');
const utf16Hex = (s: string) =>
	Array.from(s, (c) => hex(Uint8Array.of(c.charCodeAt(0), 0))).join('');
const u32 = (v: number) => hex(new Uint8Array(Uint32Array.of(v).buffer));

describe('dataSpacesStreams', () => {
	const streams = Object.fromEntries(dataSpacesStreams().map((s) => [s.path, hex(s.data)]));

	it('writes the four streams of the \\u0006DataSpaces storage', () => {
		expect(Object.keys(streams)).toEqual([
			'\u0006DataSpaces/Version',
			'\u0006DataSpaces/DataSpaceMap',
			'\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace',
			'\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary',
		]);
	});

	// Expected bytes match the streams of an encrypted file written by another tool.
	it('writes DataSpaceVersionInfo (2.1.5)', () => {
		expect(streams['\u0006DataSpaces/Version']).toBe(
			u32(60) + utf16Hex('Microsoft.Container.DataSpaces') + '010000000100000001000000',
		);
	});

	it('writes DataSpaceMap (2.1.6) and DataSpaceDefinition (2.1.7)', () => {
		expect(streams['\u0006DataSpaces/DataSpaceMap']).toBe(
			u32(8) +
				u32(1) +
				u32(104) +
				u32(1) +
				u32(0) +
				u32(32) +
				utf16Hex('EncryptedPackage') +
				u32(50) +
				utf16Hex('StrongEncryptionDataSpace') +
				'0000',
		);
		expect(streams['\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace']).toBe(
			u32(8) + u32(1) + u32(50) + utf16Hex('StrongEncryptionTransform') + '0000',
		);
	});

	it('writes TransformInfoHeader (2.1.8) with an EncryptionTransformInfo (2.1.9)', () => {
		const path = '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary';
		const primary = streams[path];
		expect(primary).toBe(
			u32(88) +
				u32(1) +
				u32(76) +
				utf16Hex('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}') +
				u32(78) +
				utf16Hex('Microsoft.Container.EncryptionTransform') +
				'0000' +
				'010000000100000001000000' +
				u32(0) +
				u32(0) +
				u32(0) +
				u32(4),
		);
		expect(primary.length / 2).toBe(200);
	});
});
