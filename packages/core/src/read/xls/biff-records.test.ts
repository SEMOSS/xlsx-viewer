import { describe, expect, it } from 'vitest';

import { BiffCursor, BiffTruncatedError, decodeChars } from './biff-cursor';
import { BOF_TYPE, RT } from './biff-opcodes';
import { readRecord, readSubstream } from './biff-records';
import { bof, continued, eof } from './test-utils/biff-builder';
import { chars, concat, record, u16 } from './test-utils/bytes';

describe('BiffCursor', () => {
	it('reads little-endian values and fails past the end', () => {
		const cur = new BiffCursor(
			Uint8Array.from([0x34, 0x12, 0xff, 0xff, 0x78, 0x56, 0x34, 0x12, 0x01]),
		);
		expect(cur.u16()).toBe(0x1234);
		expect(cur.i16()).toBe(-1);
		expect(cur.u32()).toBe(0x12345678);
		expect(cur.remaining).toBe(1);
		expect(() => cur.u16()).toThrow(BiffTruncatedError);
	});

	it('decodes 8-bit characters as Latin-1, not windows-1252', () => {
		expect(decodeChars(Uint8Array.from([0x41, 0x80, 0xe9]), 0, 3, false)).toBe('A\u0080é');
		expect(decodeChars(Uint8Array.from([0xac, 0x20]), 0, 1, true)).toBe('€');
	});

	it('switches character width at CONTINUE breaks inside a string', () => {
		// "abc" compressed, then a break with option byte 1 and "€d" as UTF-16.
		const data = Uint8Array.from([...chars('abc', false), 0x01, ...chars('€d', true)]);
		const cur = new BiffCursor(data, [3]);
		expect(cur.chars(5, false)).toBe('abc€d');
		expect(cur.remaining).toBe(0);
	});

	it('reads an option byte when the characters start exactly at a break', () => {
		const data = Uint8Array.from([0x00, ...chars('xy', false)]);
		expect(new BiffCursor(data, [0]).chars(2, true)).toBe('xy');
	});

	it('reads other values straight across breaks', () => {
		const cur = new BiffCursor(Uint8Array.from([0x01, 0x02, 0x03, 0x04]), [1, 3]);
		expect(cur.u32()).toBe(0x04030201);
	});

	it('reads the string variants', () => {
		const cur = new BiffCursor(
			Uint8Array.from([
				...u16(2),
				0,
				...chars('hi', false),
				3,
				1,
				...chars('abc', true),
				0,
				...chars('z', false),
			]),
		);
		expect(cur.unicodeString()).toBe('hi');
		expect(cur.shortString()).toBe('abc');
		expect(cur.stringNoCch(1)).toBe('z');
	});
});

describe('records', () => {
	it('merges CONTINUE records and records the breaks', () => {
		const stream = concat([continued(RT.SST, [[1, 2], [3], [4, 5, 6]]), record(RT.EOF, [])]);
		const read = readRecord(stream, 0)!;
		expect(read.record.type).toBe(RT.SST);
		expect([...read.record.data]).toEqual([1, 2, 3, 4, 5, 6]);
		expect(read.record.breaks).toEqual([2, 3]);
		expect(read.next).toBe(stream.length - 4);
	});

	it('cuts a record whose size runs past the end of the stream', () => {
		const stream = Uint8Array.from([...u16(RT.NUMBER), ...u16(14), 1, 2, 3]);
		const read = readRecord(stream, 0)!;
		expect(read.truncated).toBe(true);
		expect([...read.record.data]).toEqual([1, 2, 3]);
		expect(readRecord(stream, stream.length - 2)).toBeUndefined();
	});

	it('reads a substream and skips nested chart substreams', () => {
		const stream = concat([
			bof(BOF_TYPE.worksheet),
			record(RT.NUMBER, [1]),
			bof(BOF_TYPE.chart),
			record(RT.NUMBER, [2]),
			eof(),
			record(RT.NUMBER, [3]),
			eof(),
			bof(BOF_TYPE.worksheet),
		]);
		const sub = readSubstream(stream, 0);
		expect(sub.records.map((r) => r.data[0])).toEqual([0x00, 1, 3]);
		expect(sub.nested).toEqual([BOF_TYPE.chart]);
		expect(sub.truncated).toBe(false);
	});

	it('stops at the next sheet BOF when an EOF is missing', () => {
		const stream = concat([
			bof(BOF_TYPE.worksheet),
			record(RT.NUMBER, [1]),
			bof(BOF_TYPE.worksheet),
			eof(),
		]);
		const sub = readSubstream(stream, 0);
		expect(sub.records).toHaveLength(2);
		expect(sub.truncated).toBe(true);
	});

	it('reports a stream that ends inside a substream', () => {
		const sub = readSubstream(concat([bof(BOF_TYPE.worksheet), record(RT.NUMBER, [1])]), 0);
		expect(sub.records).toHaveLength(2);
		expect(sub.truncated).toBe(true);
		expect(readSubstream(Uint8Array.from(record(RT.NUMBER, [1])), 0).truncated).toBe(true);
	});
});
