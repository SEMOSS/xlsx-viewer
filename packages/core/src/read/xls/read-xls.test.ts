import { describe, expect, it } from 'vitest';

import { XlsxFormatError, XlsxPasswordError } from '../errors';
import { BOF_TYPE, RT } from './biff-opcodes';
import { readXlsWorkbook } from './index';
import {
	bof,
	dateMode,
	eof,
	filepass,
	sst,
	standardStyles,
	workbookStream,
	xlsFile,
} from './test-utils/biff-builder';
import { chars, concat, record, u16, u32 } from './test-utils/bytes';
import { writeCompoundFileForTest } from './test-utils/cfb-writer';
import { ptg, rgce } from './test-utils/ptg-builder';
import {
	colInfo,
	formula,
	hlink,
	labelSst,
	mergedCells,
	note,
	number,
	obj,
	row,
	shrfmla,
	txo,
	window2,
} from './test-utils/sheet-builder';

function sampleStream(): Uint8Array {
	return workbookStream(
		[...standardStyles(), sst(['alpha', 'beta', 'gamma'])],
		[
			{
				name: 'Sheet1',
				records: [
					colInfo(0, 1, 3000),
					row(0, { height: 300, custom: true }),
					number(0, 0, 15, 1.5),
					labelSst(0, 1, 15, 2),
					formula(1, 0, 15, 3, rgce(ptg.exp(1, 0))),
					shrfmla(1, 2, 0, 0, rgce(ptg.refN(-1, 0), ptg.int(2), ptg.mul)),
					formula(2, 0, 15, 6, rgce(ptg.exp(1, 0))),
					window2(),
					mergedCells([[3, 4, 0, 1]]),
					hlink([5, 5, 0, 0], { url: 'https://example.com' }),
					obj(0x19, 1),
					txo('note text'),
					note(0, 0, 1, 'me'),
				],
			},
			{ name: 'Sheet2', records: [labelSst(0, 0, 15, 0)] },
		],
	);
}

describe('format detection', () => {
	it('reads a workbook from a compound file or a bare BIFF stream', () => {
		const stream = sampleStream();
		for (const input of [xlsFile(stream), stream]) {
			const wb = readXlsWorkbook(input);
			expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1', 'Sheet2']);
			expect(wb.sheets[0].cells.get(1)?.value).toBe('gamma');
			expect(wb.warnings).toEqual([]);
		}
	});

	it('reports encrypted workbooks as unsupported password errors', () => {
		const stream = workbookStream([filepass(), ...standardStyles()], [{ name: 'S', records: [] }]);
		const err = (() => {
			try {
				readXlsWorkbook(xlsFile(stream));
			} catch (e) {
				return e;
			}
			return undefined;
		})();
		expect(err).toBeInstanceOf(XlsxPasswordError);
		expect((err as XlsxPasswordError).reason).toBe('unsupported');
		expect((err as Error).message).toMatch(/^This workbook is encrypted/);
		const ooxml = writeCompoundFileForTest([{ name: 'EncryptionInfo', data: new Uint8Array(200) }]);
		expect(() => readXlsWorkbook(ooxml)).toThrow(XlsxPasswordError);
	});

	it('rejects Excel 5.0/95 workbooks', () => {
		const biff5 = concat([bof(BOF_TYPE.globals, 0x0500), eof()]);
		for (const file of [xlsFile(biff5, [], 'Book'), xlsFile(biff5)]) {
			expect(() => readXlsWorkbook(file)).toThrow(
				new XlsxFormatError('Excel 5.0/95 workbooks are not supported'),
			);
		}
	});

	it('rejects older and unrelated formats', () => {
		expect(() => readXlsWorkbook(Uint8Array.from([0x09, 0x00, 0x04, 0x00, 2, 0, 0x10, 0]))).toThrow(
			/Excel 2.x-4.x/,
		);
		expect(() => readXlsWorkbook(Uint8Array.from([0x50, 0x4b, 3, 4, 0, 0]))).toThrow(
			XlsxFormatError,
		);
		expect(() => readXlsWorkbook(new Uint8Array(0))).toThrow(XlsxFormatError);
		const other = writeCompoundFileForTest([{ name: 'WordDocument', data: new Uint8Array(100) }]);
		expect(() => readXlsWorkbook(other)).toThrow(/workbook stream/);
		const broken = xlsFile(sampleStream()).slice(0, 600);
		expect(() => readXlsWorkbook(broken)).toThrow(XlsxFormatError);
		const notGlobals = concat([bof(BOF_TYPE.worksheet), eof()]);
		expect(() => readXlsWorkbook(xlsFile(notGlobals))).toThrow(/globals/);
		expect(() => readXlsWorkbook(xlsFile(Uint8Array.from(record(RT.NUMBER, [1, 2]))))).toThrow(
			/BOF/,
		);
	});

	it('reports VBA projects', () => {
		const stream = workbookStream([], [{ name: 'S', records: [] }]);
		const wb = readXlsWorkbook(
			xlsFile(stream, [{ name: '_VBA_PROJECT_CUR', data: new Uint8Array(10) }]),
		);
		expect(wb.warnings.map((w) => w.code)).toEqual(['macros-dropped']);
	});

	it('reads the 1904 date system', () => {
		const wb = readXlsWorkbook(
			xlsFile(workbookStream([dateMode(true)], [{ name: 'S', records: [] }])),
		);
		expect(wb.date1904).toBe(true);
	});
});

describe('robustness', () => {
	const allowed = (e: unknown): boolean =>
		e instanceof XlsxFormatError || e instanceof XlsxPasswordError;

	it('never crashes on a stream cut at any byte', () => {
		const stream = sampleStream();
		let read = 0;
		for (let cut = 0; cut <= stream.length; cut++) {
			try {
				const wb = readXlsWorkbook(stream.slice(0, cut));
				expect(wb.sheets.length).toBeGreaterThan(0);
				read++;
			} catch (e) {
				if (!allowed(e)) {
					throw new Error(`cut at ${cut}: ${String(e)}`);
				}
			}
		}
		expect(read).toBeGreaterThan(stream.length / 2);
	});

	it('never crashes on corrupted bytes', () => {
		const stream = sampleStream();
		for (let i = 0; i < stream.length; i++) {
			for (const value of [0x00, 0xff, stream[i] ^ 0x5a]) {
				const copy = stream.slice();
				copy[i] = value;
				try {
					readXlsWorkbook(copy);
				} catch (e) {
					if (!allowed(e)) {
						throw new Error(`byte ${i} = ${value}: ${String(e)}`);
					}
				}
			}
		}
	});

	it('reports truncated sheets and keeps the cells read so far', () => {
		const stream = sampleStream();
		const wb = readXlsWorkbook(stream.slice(0, stream.length - 30));
		expect(wb.sheets[0].cells.size).toBeGreaterThan(0);
		expect(wb.warnings.map((w) => w.code)).toContain('truncated-record');
	});
});

describe('document properties', () => {
	function propertySet(props: Array<[number, number[]]>): Uint8Array {
		const values: number[] = [];
		const offsets: number[] = [];
		const head = 8 + props.length * 8;
		for (const [, bytes] of props) {
			offsets.push(head + values.length);
			values.push(...bytes, ...new Array<number>((4 - (bytes.length % 4)) % 4).fill(0));
		}
		const set = [
			...u32(head + values.length),
			...u32(props.length),
			...props.flatMap(([id], i) => [...u32(id), ...u32(offsets[i])]),
			...values,
		];
		return Uint8Array.from([
			...u16(0xfffe),
			...u16(0),
			...u32(0x20006),
			...new Array<number>(16).fill(0),
			...u32(1),
			...new Array<number>(16).fill(1),
			...u32(48),
			...set,
		]);
	}
	const lpstr = (s: string): number[] => [
		...u16(0x1e),
		...u16(0),
		...u32(s.length + 1),
		...chars(s, false),
		0,
	];
	const lpwstr = (s: string): number[] => [
		...u16(0x1f),
		...u16(0),
		...u32(s.length + 1),
		...chars(`${s}\u0000`, true),
	];
	const filetime = (date: Date): number[] => {
		const ticks = (date.getTime() + 11644473600000) * 10000;
		return [...u16(0x40), ...u16(0), ...u32(ticks % 2 ** 32), ...u32(Math.floor(ticks / 2 ** 32))];
	};

	it('reads the summary information streams', () => {
		const summary = propertySet([
			[1, [...u16(2), ...u16(0), ...u16(1252), ...u16(0)]],
			[2, lpstr('Quarterly report')],
			[4, lpstr('Zoë')],
			[8, lpwstr('Ünal')],
			[12, filetime(new Date('2003-04-05T06:07:08Z'))],
		]);
		const docSummary = propertySet([[15, lpstr('ACME')]]);
		const wb = readXlsWorkbook(
			xlsFile(workbookStream([], [{ name: 'S', records: [] }]), [
				{ name: '\u0005SummaryInformation', data: summary },
				{ name: '\u0005DocumentSummaryInformation', data: docSummary },
			]),
		);
		expect(wb.properties).toEqual({
			title: 'Quarterly report',
			creator: 'Zoë',
			lastModifiedBy: 'Ünal',
			created: '2003-04-05T06:07:08Z',
			company: 'ACME',
		});
	});

	it('ignores damaged property streams', () => {
		const wb = readXlsWorkbook(
			xlsFile(workbookStream([], [{ name: 'S', records: [] }]), [
				{ name: '\u0005SummaryInformation', data: Uint8Array.from([0xfe, 0xff, 1, 2, 3]) },
			]),
		);
		expect(wb.properties).toEqual({});
	});
});
