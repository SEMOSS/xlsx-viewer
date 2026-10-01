import { describe, expect, it } from 'vitest';

import { RT } from './biff-opcodes';
import { readRecord } from './biff-records';
import { parseSst, richTextRuns } from './sst';
import { continued } from './test-utils/biff-builder';
import { chars, concat, u16, u32 } from './test-utils/bytes';

function sstRecord(chunks: number[][]) {
	return readRecord(concat([continued(RT.SST, chunks)]), 0)!.record;
}

describe('shared string table', () => {
	it('reads plain, 16-bit, rich and phonetic strings', () => {
		const record = sstRecord([
			[
				...u32(4),
				...u32(4),
				...u16(5),
				0x00,
				...chars('Hello', false),
				...u16(3),
				0x01,
				...chars('é€!', true),
				// Rich string: 2 runs.
				...u16(4),
				0x08,
				...u16(2),
				...chars('Bold', false),
				...u16(0),
				...u16(5),
				...u16(2),
				...u16(6),
				// Phonetic data (ExtRst) is skipped.
				...u16(2),
				0x04,
				...u32(6),
				...chars('ab', false),
				1,
				2,
				3,
				4,
				5,
				6,
			],
		]);
		const sst = parseSst(record);
		expect(sst.strings).toEqual(['Hello', 'é€!', 'Bold', 'ab']);
		expect(sst.runs.get(2)).toEqual([
			{ ich: 0, ifnt: 5 },
			{ ich: 2, ifnt: 6 },
		]);
		expect(sst.truncated).toBe(false);
	});

	it('continues strings across CONTINUE records, switching character width', () => {
		const record = sstRecord([
			[...u32(3), ...u32(3), ...u16(6), 0x00, ...chars('abc', false)],
			// Continuation of "abc..." stored as 16-bit, then the next string starts in the same record.
			[0x01, ...chars('dĀf', true), ...u16(3), 0x01, ...chars('xyz', true)],
			// A 16-bit string continued as compressed characters.
			[...u16(4), 0x01, ...chars('中', true)],
			[0x00, ...chars('123', false)],
		]);
		const sst = parseSst(record);
		expect(sst.strings).toEqual(['abcdĀf', 'xyz', '中123']);
		expect(sst.truncated).toBe(false);
	});

	it('reads formatting runs and phonetic data that cross a boundary', () => {
		const record = sstRecord([
			[
				...u32(2),
				...u32(2),
				...u16(2),
				0x0c,
				...u16(1),
				...u32(4),
				...chars('ab', false),
				...u16(1),
			],
			[...u16(5), 9, 9, 9, 9, ...u16(1), 0x00, ...chars('c', false)],
		]);
		const sst = parseSst(record);
		expect(sst.strings).toEqual(['ab', 'c']);
		expect(sst.runs.get(0)).toEqual([{ ich: 1, ifnt: 5 }]);
	});

	it('keeps the strings read before a truncation', () => {
		const record = sstRecord([
			[
				...u32(3),
				...u32(3),
				...u16(1),
				0x00,
				...chars('a', false),
				...u16(10),
				0x00,
				...chars('bc', false),
			],
		]);
		const sst = parseSst(record);
		expect(sst.strings).toEqual(['a']);
		expect(sst.truncated).toBe(true);
	});

	it('caps an absurd string count by the record size', () => {
		const sst = parseSst(sstRecord([[...u32(0xffffffff), ...u32(0xffffffff), ...u16(1), 0, 65]]));
		expect(sst.strings).toEqual(['A']);
		expect(sst.truncated).toBe(true);
	});
});

describe('rich text runs', () => {
	const fonts = [{ name: 'Arial' }, { name: 'Arial', bold: true }, { name: 'Arial', italic: true }];

	it('splits the text at run starts and keeps the text before the first run plain', () => {
		const runs = richTextRuns(
			'Hello world',
			[
				{ ich: 6, ifnt: 1 },
				{ ich: 11, ifnt: 0 },
			],
			(i) => fonts[i],
		);
		expect(runs).toEqual([
			{ text: 'Hello ' },
			{ text: 'world', font: { name: 'Arial', bold: true } },
		]);
	});

	it('returns undefined when no run applies', () => {
		expect(richTextRuns('abc', [{ ich: 3, ifnt: 1 }], (i) => fonts[i])).toBeUndefined();
	});
});
