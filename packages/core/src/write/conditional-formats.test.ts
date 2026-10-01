import { describe, expect, it } from 'vitest';

import { StylePool } from '../styles/style-pool';
import { createDefaultStyleSheet } from '../styles/defaults';
import type { ConditionalFormat } from '../types/conditional-format';
import type { DifferentialFormat } from '../types/style';
import { serializeElement } from '../xml';
import {
	buildConditionalFormats,
	CF_EXT_URI,
	type CfStyles,
	DATABAR_EXT_URI,
} from './conditional-formats';
import { buildDataValidations, DV_EXT_URI } from './data-validations';
import { buildSparklineExt } from './sparklines';
import { childNames, pick, pickAll } from './test-utils/package-check';

const range = { top: 1, left: 1, bottom: 9, right: 1 };

/** Styles accessor over a plain list (interning through StylePool, like the writer). */
function cfStyles(dxfs: DifferentialFormat[] = []): CfStyles & { list: DifferentialFormat[] } {
	const sheet = { ...createDefaultStyleSheet(), dxfs };
	const pool = new StylePool(sheet);
	return { list: dxfs, dxfs: () => dxfs, internDxf: (d) => pool.internDxf(d) };
}

describe('conditional formatting', () => {
	it('writes 2007 rules with formulas, operators and validated dxf ids', () => {
		const cfs: ConditionalFormat[] = [
			{
				ranges: [range, { top: 0, left: 4, bottom: 0, right: 4 }],
				rules: [
					{
						type: 'cellIs',
						priority: 2,
						operator: 'between',
						formulas: ['1', 'MAXIFS(C:C,D:D,1)'],
						dxfId: 0,
						stopIfTrue: true,
					},
					{ type: 'expression', priority: 1, formulas: ['MOD(ROW(),2)=0'], dxfId: 7 },
					{ type: 'top10', priority: 3, rank: 5, bottom: true, percent: true, dxfId: 0 },
					{ type: 'aboveAverage', priority: 4, aboveAverage: false, stdDev: 1, dxfId: 0 },
				],
			},
		];
		const out = buildConditionalFormats(cfs, cfStyles([{}]));
		expect(out.ext).toBeNull();
		const cf = out.elements[0];
		expect(cf.attrs.sqref).toBe('B2:B10 E1');
		const rules = pickAll(cf, 'cfRule');
		expect(rules[0].attrs).toEqual({
			type: 'cellIs',
			dxfId: '0',
			priority: '2',
			stopIfTrue: '1',
			operator: 'between',
		});
		expect(pickAll(rules[0], 'formula').map((f) => f.children[0])).toEqual([
			'1',
			'_xlfn.MAXIFS(C:C,D:D,1)',
		]);
		expect(rules[1].attrs.dxfId).toBeUndefined();
		expect(rules[2].attrs).toMatchObject({ type: 'top10', rank: '5', bottom: '1', percent: '1' });
		expect(rules[3].attrs).toMatchObject({ aboveAverage: '0', stdDev: '1' });
	});

	it('interns inline formats of 2007 rules into dxfs (deduplicated) and prefers them over dxfId', () => {
		const styles = cfStyles([{ font: { bold: true } }]);
		const red = {
			fill: {
				type: 'pattern' as const,
				patternType: 'solid' as const,
				bgColor: { rgb: 'FFFFC7CE' },
			},
		};
		const out = buildConditionalFormats(
			[
				{
					ranges: [range],
					rules: [
						{ type: 'expression', priority: 1, formulas: ['A1>1'], dxf: red },
						{ type: 'expression', priority: 2, formulas: ['A1>2'], dxf: { ...red } },
						{
							type: 'expression',
							priority: 3,
							formulas: ['A1>3'],
							dxf: { font: { bold: true } },
							dxfId: 9,
						},
					],
				},
			],
			styles,
		);
		const ids = pickAll(out.elements[0], 'cfRule').map((r) => r.attrs.dxfId);
		expect(ids).toEqual(['1', '1', '0']);
		expect(styles.list).toHaveLength(2);
		expect(styles.list[1]).toEqual(red);
	});

	it('generates the formulas Excel requires for text, blank, error and period rules', () => {
		const out = buildConditionalFormats(
			[
				{
					ranges: [range],
					rules: [
						{ type: 'containsText', priority: 1, text: 'a"b' },
						{ type: 'beginsWith', priority: 2, text: 'x' },
						{ type: 'containsBlanks', priority: 3 },
						{ type: 'timePeriod', priority: 4, timePeriod: 'last7Days' },
					],
				},
			],
			cfStyles(),
		);
		const rules = pickAll(out.elements[0], 'cfRule');
		expect(rules[0].attrs.operator).toBe('containsText');
		expect(pick(rules[0], 'formula')?.children[0]).toBe('NOT(ISERROR(SEARCH("a""b",B2)))');
		expect(pick(rules[1], 'formula')?.children[0]).toBe('LEFT(B2,LEN("x"))="x"');
		expect(pick(rules[2], 'formula')?.children[0]).toBe('LEN(TRIM(B2))=0');
		expect(pick(rules[3], 'formula')?.children[0]).toBe(
			'AND(TODAY()-FLOOR(B2,1)<=6,FLOOR(B2,1)<=TODAY())',
		);
	});

	it('writes color scales and drops incomplete rules', () => {
		const out = buildConditionalFormats(
			[
				{
					ranges: [range],
					rules: [
						{
							type: 'colorScale',
							priority: 1,
							colorScale: {
								cfvos: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
								colors: [{ rgb: 'FFF8696B' }, { rgb: 'FFFFEB84' }, { rgb: 'FF63BE7B' }],
							},
						},
						{ type: 'expression', priority: 2 },
						{ type: 'dataBar', priority: 3 },
					],
				},
			],
			cfStyles(),
		);
		expect(out.skipped).toBe(2);
		const scale = pick(out.elements[0], 'cfRule', 'colorScale')!;
		expect(childNames(scale)).toEqual(['cfvo', 'cfvo', 'cfvo', 'color', 'color', 'color']);
		expect(pickAll(scale, 'cfvo')[1].attrs).toEqual({ type: 'percentile', val: '50' });
	});

	it('links Excel 2010 data bar properties through an x14 extension', () => {
		const out = buildConditionalFormats(
			[
				{
					ranges: [range],
					rules: [
						{
							type: 'dataBar',
							priority: 1,
							extId: '{ABCDEF01-0000-0000-0000-000000000001}',
							dataBar: {
								cfvos: [{ type: 'autoMin' }, { type: 'num', value: '100' }],
								color: { rgb: 'FF638EC6' },
								gradient: false,
								negativeFillColor: { rgb: 'FFFF0000' },
								axisPosition: 'middle',
							},
						},
					],
				},
			],
			cfStyles(),
		);
		const rule = pick(out.elements[0], 'cfRule')!;
		expect(childNames(rule)).toEqual(['dataBar', 'extLst']);
		expect(pickAll(pick(rule, 'dataBar'), 'cfvo').map((c) => c.attrs)).toEqual([
			{ type: 'min' },
			{ type: 'num', val: '100' },
		]);
		const link = pick(rule, 'extLst', 'ext')!;
		expect(link.attrs.uri).toBe(DATABAR_EXT_URI);
		expect(pick(link, 'id')?.children[0]).toBe('{ABCDEF01-0000-0000-0000-000000000001}');
		expect(out.ext?.attrs.uri).toBe(CF_EXT_URI);
		const x14 = pick(out.ext!, 'conditionalFormattings', 'conditionalFormatting')!;
		expect(childNames(x14)).toEqual(['cfRule', 'sqref']);
		const x14Rule = pick(x14, 'cfRule')!;
		expect(x14Rule.attrs).toEqual({
			type: 'dataBar',
			id: '{ABCDEF01-0000-0000-0000-000000000001}',
		});
		const bar = pick(x14Rule, 'dataBar')!;
		expect(bar.attrs).toEqual({ gradient: '0', axisPosition: 'middle' });
		expect(childNames(bar)).toEqual(['cfvo', 'cfvo', 'negativeFillColor']);
		expect(pick(bar, 'cfvo')?.attrs.type).toBe('autoMin');
		expect(pickAll(bar, 'cfvo')[1].children.length).toBe(1);
	});

	it('writes x14-only rules (custom icon sets, extension rules with inline dxf) only in the extension', () => {
		const out = buildConditionalFormats(
			[
				{
					ranges: [range],
					rules: [
						{
							type: 'iconSet',
							priority: 1,
							iconSet: {
								iconSet: '3Stars',
								cfvos: [
									{ type: 'percent', value: '0' },
									{ type: 'percent', value: '33' },
									{ type: 'percent', value: '67' },
								],
								customIcons: [{ iconSet: '3Arrows', iconId: 0 }],
							},
						},
						{
							type: 'cellIs',
							priority: 2,
							operator: 'equal',
							formulas: ['Other!$A$1'],
							fromExtension: true,
							dxf: { font: { bold: true } },
						},
					],
				},
			],
			cfStyles(),
		);
		expect(out.elements).toHaveLength(0);
		const x14 = pick(out.ext!, 'conditionalFormattings', 'conditionalFormatting')!;
		const [icons, cellIs] = pickAll(x14, 'cfRule');
		expect(icons.attrs).toMatchObject({ type: 'iconSet', priority: '1' });
		expect(icons.attrs.id).toMatch(/^\{[0-9A-F-]{36}\}$/);
		const set = pick(icons, 'iconSet')!;
		expect(set.attrs).toEqual({ iconSet: '3Stars', custom: '1' });
		expect(childNames(set)).toEqual(['cfvo', 'cfvo', 'cfvo', 'cfIcon']);
		expect(childNames(cellIs)).toEqual(['f', 'dxf']);
		expect(serializeElement(pick(cellIs, 'dxf')!)).toBe('<x14:dxf><font><b/></font></x14:dxf>');
		expect(pick(x14, 'sqref')?.children[0]).toBe('B2:B10');
	});
});

describe('data validations', () => {
	it('writes 2007 validations and x14 ones for cross-sheet sources', () => {
		const out = buildDataValidations([
			{
				ranges: [range],
				type: 'whole',
				operator: 'greaterThan',
				formula1: '0',
				showErrorMessage: true,
				errorStyle: 'warning',
				error: 'e'.repeat(300),
			},
			{
				ranges: [range],
				type: 'list',
				formula1: 'Lists!$A$1:$A$3',
				fromExtension: true,
				allowBlank: true,
			},
			{ ranges: [], type: 'list', formula1: '"a"' },
		]);
		const dv = pick(out.element!, 'dataValidation')!;
		expect(out.element?.attrs.count).toBe('1');
		expect(dv.attrs).toMatchObject({
			type: 'whole',
			errorStyle: 'warning',
			operator: 'greaterThan',
			showErrorMessage: '1',
			sqref: 'B2:B10',
		});
		expect(dv.attrs.error).toHaveLength(255);
		expect(out.ext?.attrs.uri).toBe(DV_EXT_URI);
		const x14 = pick(out.ext!, 'dataValidations', 'dataValidation')!;
		expect(childNames(x14)).toEqual(['formula1', 'sqref']);
		expect(pick(x14, 'formula1', 'f')?.children[0]).toBe('Lists!$A$1:$A$3');
	});
});

describe('sparklines', () => {
	it('writes groups with default colors and sparkline references', () => {
		const ext = buildSparklineExt([
			{
				type: 'line',
				sparklines: [{ formula: 'Data!A2:F2', location: { row: 1, col: 6 } }],
				high: true,
				seriesColor: { rgb: 'FF376092' },
			},
			{ type: 'column', sparklines: [] },
		])!;
		const groups = pickAll(pick(ext, 'sparklineGroups'), 'sparklineGroup');
		expect(groups).toHaveLength(1);
		expect(groups[0].attrs).toEqual({ high: '1' });
		expect(childNames(groups[0])).toEqual([
			'colorSeries',
			'colorNegative',
			'colorAxis',
			'colorMarkers',
			'colorFirst',
			'colorLast',
			'colorHigh',
			'colorLow',
			'sparklines',
		]);
		expect(pick(groups[0], 'colorSeries')?.attrs.rgb).toBe('FF376092');
		const line = pick(groups[0], 'sparklines', 'sparkline')!;
		expect(childNames(line)).toEqual(['f', 'sqref']);
		expect(pick(line, 'sqref')?.children[0]).toBe('G2');
		expect(buildSparklineExt([])).toBeNull();
	});
});
