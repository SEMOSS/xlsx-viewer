import { describe, expect, it } from 'vitest';

import { compoundFileEntryNames } from './cfb';
import { XlsxFormatError, XlsxPasswordError, readWorkbook } from './index';

/**
 * Build a minimal compound file: header, one FAT sector (sector 0), one
 * directory sector (sector 1) and an optional data sector (sector 2).
 */
function compoundFile(entries: string[], sectorShift = 9, data?: string): Uint8Array {
	const size = 1 << sectorShift;
	const bytes = new Uint8Array(size * (data ? 4 : 3));
	const view = new DataView(bytes.buffer);
	bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
	view.setUint16(0x18, 0x3e, true);
	view.setUint16(0x1a, sectorShift === 9 ? 3 : 4, true);
	view.setUint16(0x1c, 0xfffe, true);
	view.setUint16(0x1e, sectorShift, true);
	view.setUint16(0x20, 6, true);
	view.setUint32(0x2c, 1, true);
	view.setUint32(0x30, 1, true);
	view.setUint32(0x38, 4096, true);
	view.setUint32(0x3c, 0xfffffffe, true);
	view.setUint32(0x44, 0xfffffffe, true);
	for (let i = 0; i < 109; i++) {
		view.setUint32(0x4c + i * 4, i === 0 ? 0 : 0xffffffff, true);
	}
	const fat = size;
	for (let i = 0; i < size / 4; i++) {
		view.setUint32(fat + i * 4, 0xffffffff, true);
	}
	view.setUint32(fat, 0xfffffffd, true);
	view.setUint32(fat + 4, 0xfffffffe, true);
	if (data) {
		view.setUint32(fat + 8, 0xfffffffe, true);
	}
	const dir = size * 2;
	['Root Entry', ...entries].forEach((name, i) => {
		const e = dir + i * 128;
		for (let c = 0; c < name.length; c++) {
			view.setUint16(e + c * 2, name.charCodeAt(c), true);
		}
		view.setUint16(e + 0x40, (name.length + 1) * 2, true);
		bytes[e + 0x42] = i === 0 ? 5 : 2;
	});
	if (data) {
		for (let c = 0; c < data.length; c++) {
			view.setUint16(size * 3 + c * 2, data.charCodeAt(c), true);
		}
	}
	return bytes;
}

describe('compound files', () => {
	it('lists directory entry names', () => {
		expect(compoundFileEntryNames(compoundFile(['Workbook', 'SummaryInformation']))).toEqual([
			'Root Entry',
			'Workbook',
			'SummaryInformation',
		]);
		expect(compoundFileEntryNames(compoundFile(['EncryptionInfo'], 12))).toEqual([
			'Root Entry',
			'EncryptionInfo',
		]);
		expect(compoundFileEntryNames(new Uint8Array(100))).toBeUndefined();
	});

	it('reports encrypted packages as password errors', async () => {
		for (const shift of [9, 12]) {
			const err = await readWorkbook(
				compoundFile(['EncryptionInfo', 'EncryptedPackage', '\u0006DataSpaces'], shift),
			).catch((e: unknown) => e);
			expect(err).toBeInstanceOf(XlsxPasswordError);
			expect((err as XlsxPasswordError).reason).toBe('unsupported');
		}
	});

	it('reports legacy workbooks as format errors, even when their data mentions EncryptionInfo', async () => {
		const legacy = compoundFile(['Workbook'], 9, 'EncryptionInfo');
		await expect(readWorkbook(legacy)).rejects.toBeInstanceOf(XlsxFormatError);
		// The synthetic file has no real BIFF stream: the .xls reader reports it.
		await expect(readWorkbook(legacy)).rejects.toThrow(
			'The file does not contain an Excel workbook stream',
		);
	});
});
